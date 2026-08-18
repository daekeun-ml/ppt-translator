#!/usr/bin/env python3
"""PowerPoint Translator CLI using Click"""

import click
import functools
import logging
import multiprocessing
import os
import re
import sys
import threading
from zipfile import BadZipFile, ZipFile
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Optional

from .cache import build_cache
from .config import Config
from .glossary import find_default_glossary, get_glossary_for_language, load_glossary
from .logging_utils import quiet_batch_worker_logs, quiet_dependency_logs
from .markdown_exporter import MarkdownExporter
from .post_processing import PowerPointPostProcessor
from .ppt_handler import PowerPointTranslator
from .pricing import estimate_cost, estimate_tokens
from .progress import make_progress, make_slide_progress_callback, render_metrics_line


def _configure_logging() -> None:
    """Route all logging through rich so it doesn't collide with the progress bar.

    Plain `logging.basicConfig` emits to stderr without any awareness of the
    rich Live display — so when a log line fires while the bar is rendering
    you get `$0.00562026-04-25 ...`. RichHandler cooperates with the Live
    instance and prints each log line cleanly above the bar.
    """
    try:
        from rich.logging import RichHandler
        logging.basicConfig(
            level=logging.INFO,
            format='%(message)s',
            datefmt='%H:%M:%S',
            handlers=[RichHandler(rich_tracebacks=True, show_path=False, markup=False)],
            force=True,  # replace any handlers installed by imports
        )
    except ImportError:
        # Fallback: plain stream logging (will visually collide with the bar,
        # but the tool still works).
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        )
    quiet_dependency_logs()


_configure_logging()
logger = logging.getLogger(__name__)


def _common_options(func):
    """CLI flags shared by every translate-* command.

    Kept as a decorator so the options stay in sync across all commands — adding
    a flag here automatically makes it available to translate / translate-slides
    / batch-translate without copy-paste drift.
    """
    @click.option('--source-language', default=None,
                  help='Source language code (e.g., en, ko). Auto-detected if omitted.')
    @click.option('--no-detect-source', is_flag=True,
                  help='Skip source language auto-detection (let the model infer from context)')
    @click.option('-g', '--glossary', 'glossary_path', type=click.Path(exists=True, dir_okay=False),
                  default=None, help='Glossary YAML file (default: ./glossary.yaml)')
    @click.option('--cache-backend', type=click.Choice(['sqlite', 'memory', 'none']),
                  default=lambda: os.getenv('CACHE_BACKEND', 'sqlite'),
                  help='Translation cache backend (default: sqlite)')
    @click.option('--cache-path', default=lambda: os.getenv('CACHE_PATH', str(Path.home() / '.ppt-translator' / 'cache.db')),
                  help='SQLite cache path (ignored for memory/none)')
    @click.option('--no-cache', is_flag=True, help='Disable translation cache (overrides --cache-backend)')
    @click.option('--dry-run', is_flag=True, help='Estimate cost without translating or saving')
    @click.option('--no-charts', is_flag=True, help='Skip chart title/axis/category translation')
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)
    return wrapper


def _slide_parallel_options(func):
    """CLI flags for parallelizing slides within one presentation."""
    @click.option(
        '--slide-workers',
        default=Config.SLIDE_WORKERS,
        type=click.IntRange(min=1),
        show_default=True,
        help='Maximum slide translation workers per presentation; use 1 to disable',
    )
    @click.option(
        '--slides-per-worker',
        default=Config.SLIDES_PER_WORKER,
        type=click.IntRange(min=1),
        show_default=True,
        help='Slides assigned to each worker before creating another chunk',
    )
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs)
    return wrapper


def _resolve_glossary(glossary_path: Optional[str], target_language: str):
    """Resolve the glossary path (explicit flag → default search) and load it."""
    path = glossary_path
    if path is None:
        default = find_default_glossary()
        if default is not None:
            path = str(default)
            click.echo(f"📖 Using glossary: {path}")
    elif path:
        click.echo(f"📖 Using glossary: {path}")
    glossary_map = load_glossary(path) if path else {}
    return get_glossary_for_language(glossary_map, target_language)


def _resolve_cache_backend(cache_backend: str, no_cache: bool) -> str:
    if no_cache:
        return 'none'
    return cache_backend


def _validate_markdown_language(language: str) -> None:
    if language != "source" and language not in Config.LANGUAGE_MAP:
        available = ", ".join(["source", *sorted(Config.LANGUAGE_MAP)])
        raise click.BadParameter(
            f"Unsupported output language '{language}'. Available: {available}",
            param_hint="--language",
        )


def _markdown_metrics_line(metrics) -> str:
    parts = [
        f"calls={getattr(metrics, 'api_calls', 0)}",
        (
            f"cache={getattr(metrics, 'cache_hits', 0)}/"
            f"{getattr(metrics, 'cache_hits', 0) + getattr(metrics, 'cache_misses', 0)}"
        ),
        (
            f"tokens={getattr(metrics, 'tokens_in', 0)}+"
            f"{getattr(metrics, 'tokens_out', 0)}"
        ),
    ]
    web_queries = getattr(metrics, "web_queries", 0)
    if web_queries:
        parts.append(
            f"web={web_queries} queries/"
            f"{getattr(metrics, 'web_results', 0)} results"
        )
    return "  ".join(parts)


def _resolve_batch_workers(
    file_workers: int,
    slide_workers: int,
    file_count: int,
) -> tuple[int, int]:
    """Resolve independent file and per-presentation slide concurrency."""
    active_file_workers = max(1, min(int(file_workers), max(1, file_count)))
    per_file_slide_workers = max(1, int(slide_workers))
    return active_file_workers, per_file_slide_workers


def _slide_chunk_plan(
    slide_count: int,
    slide_workers: int,
    slides_per_worker: int,
) -> tuple[list[int], int]:
    """Return chunk sizes and active workers for one presentation."""
    if slide_workers <= 1 or slide_count <= slides_per_worker:
        return [slide_count], 1
    chunks = [
        min(slides_per_worker, slide_count - start)
        for start in range(0, slide_count, slides_per_worker)
    ]
    return chunks, min(slide_workers, len(chunks))


def _presentation_slide_count(path: Path) -> int:
    """Read a PPTX slide count without loading the full presentation model."""
    if path.suffix.lower() != ".pptx":
        return 0
    prefix = "ppt/slides/slide"
    suffix = ".xml"
    try:
        with ZipFile(path) as archive:
            return sum(
                1
                for name in archive.namelist()
                if name.startswith(prefix)
                and name.endswith(suffix)
                and name[len(prefix):-len(suffix)].isdigit()
            )
    except (BadZipFile, OSError):
        return 0


def _discover_presentation_files(
    input_path: Path,
    output_path: Path,
    recursive: bool,
) -> list[Path]:
    """Find source presentations while excluding the output directory tree."""
    input_root = input_path.expanduser().resolve()
    output_root = output_path.expanduser().resolve()
    if input_root == output_root:
        raise ValueError(
            "Batch output folder must be different from the input folder."
        )

    extensions = {'.pptx', '.ppt'}
    if not recursive:
        return sorted(
            path
            for path in input_root.iterdir()
            if path.is_file() and path.suffix.lower() in extensions
        )

    presentations: list[Path] = []
    for root, directories, filenames in os.walk(input_root):
        root_path = Path(root)
        directories[:] = [
            directory
            for directory in directories
            if (root_path / directory).resolve() != output_root
        ]
        for filename in filenames:
            path = root_path / filename
            if path.suffix.lower() not in extensions:
                continue
            if path.resolve().is_relative_to(output_root):
                continue
            presentations.append(path)
    return sorted(presentations)


def _discover_markdown_source_files(
    input_path: Path,
    output_path: Path,
    recursive: bool,
) -> list[Path]:
    """Find source decks without re-exporting generated translation folders."""
    input_root = input_path.expanduser().resolve()
    files = _discover_presentation_files(input_root, output_path, recursive)
    generated_file = re.compile(
        r"_translated_[a-z]{2,3}(?:-[a-z0-9]+)?$",
        re.IGNORECASE,
    )
    sources = []
    for path in files:
        if path.suffix.lower() != ".pptx":
            continue
        relative = path.resolve().relative_to(input_root)
        generated_parent = any(
            part.lower().startswith(("translated_", "markdown_"))
            for part in relative.parts[:-1]
        )
        if generated_parent or generated_file.search(path.stem):
            continue
        sources.append(path)
    return sources


def _collect_for_dry_run(translator: PowerPointTranslator, input_file: str,
                         target_language: str, model_id: str,
                         slide_numbers: Optional[list] = None,
                         detect_source: bool = True) -> None:
    """Print a dry-run summary and return without translating."""
    stats = translator.collect_all_texts(input_file, slide_numbers=slide_numbers,
                                         detect_source=detect_source)
    src_lang = stats.get('source_language') or 'en'
    tokens_in = estimate_tokens(stats['total_chars'], src_lang)
    tokens_out = estimate_tokens(stats['total_chars'], target_language)
    cost = estimate_cost(tokens_in, tokens_out, model_id)

    same_lang = src_lang and src_lang.split('-')[0].lower() == target_language.split('-')[0].lower()

    click.echo()
    click.echo(f"📊 Dry-Run Report: {input_file} → {target_language}")
    click.echo(f"  Source language:     {stats.get('source_language') or '(not detected)'}")
    click.echo(f"  Total slides:        {stats['slide_count']}")
    click.echo(f"  Translatable items:  {stats['translatable_items']}")
    click.echo(f"  Total characters:    {stats['total_chars']:,}"
               + (f" (avg {stats['total_chars'] // max(stats['translatable_items'], 1)}/item)"
                  if stats['translatable_items'] else ""))

    if same_lang:
        click.echo()
        click.echo(f"⏭️  Source '{src_lang}' already matches target '{target_language}'. "
                   "No API calls would be made.")
    else:
        click.echo()
        click.echo(f"💰 Cost Estimate ({model_id}):")
        click.echo(f"  Input tokens (est.):  ~{tokens_in:,}")
        click.echo(f"  Output tokens (est.): ~{tokens_out:,}")
        if cost > 0:
            click.echo(f"  Estimated cost:       ${cost:.4f}")
        else:
            click.echo(f"  Estimated cost:       (pricing unavailable for this model)")

    if stats.get('sample_texts'):
        click.echo()
        click.echo("📝 Sample texts:")
        for i, sample in enumerate(stats['sample_texts'][:5], 1):
            click.echo(f"  [{i}] {sample[:80]}{'…' if len(sample) > 80 else ''}")
    click.echo()
    click.echo("ℹ️  Run without --dry-run to perform actual translation.")


@click.group()
@click.version_option()
def cli():
    """PowerPoint Translator using Amazon Bedrock Mantle"""
    pass


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('-t', '--target-language', default=Config.DEFAULT_TARGET_LANGUAGE, help='Target language')
@click.option('-o', '--output-file', help='Output file path')
@click.option('-m', '--model-id', default=Config.DEFAULT_MODEL_ID, help='Bedrock Mantle model ID')
@click.option('--no-polishing', is_flag=True, help='Disable natural language polishing')
@_slide_parallel_options
@_common_options
def translate(input_file, target_language, output_file, model_id, no_polishing,
              slide_workers, slides_per_worker,
              source_language, no_detect_source,
              glossary_path, cache_backend, cache_path, no_cache, dry_run, no_charts):
    """Translate entire PowerPoint presentation"""
    glossary = _resolve_glossary(glossary_path, target_language)
    backend = _resolve_cache_backend(cache_backend, no_cache)
    auto_detect = not no_detect_source

    if not output_file:
        input_path = Path(input_file)
        output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{input_path.suffix}")

    if dry_run:
        translator = PowerPointTranslator(model_id, not no_polishing, glossary=glossary,
                                          translate_charts=not no_charts,
                                          source_language=source_language,
                                          auto_detect_source=auto_detect,
                                          slide_workers=slide_workers,
                                          slides_per_worker=slides_per_worker)
        _collect_for_dry_run(translator, input_file, target_language, model_id,
                             detect_source=auto_detect)
        return

    click.echo(f"🚀 Starting translation: {input_file} -> {target_language}")
    with build_cache(backend, cache_path) as cache:
        translator = PowerPointTranslator(model_id, not no_polishing, cache=cache,
                                          glossary=glossary, translate_charts=not no_charts,
                                          source_language=source_language,
                                          auto_detect_source=auto_detect,
                                          slide_workers=slide_workers,
                                          slides_per_worker=slides_per_worker)
        slide_count = translator.get_slide_count(input_file)
        chunk_sizes, active_slide_workers = _slide_chunk_plan(
            slide_count,
            slide_workers,
            slides_per_worker,
        )
        if len(chunk_sizes) > 1:
            chunk_text = "+".join(str(size) for size in chunk_sizes)
            click.echo(
                f"🧵 Translation plan: {slide_count} slides → "
                f"{len(chunk_sizes)} chunks ({chunk_text}), "
                f"{active_slide_workers} slide workers"
            )
            task_description = (
                f"Translating ({active_slide_workers} slide workers)"
            )
        else:
            click.echo(
                f"📄 Translation plan: {slide_count} slides → sequential"
            )
            task_description = "Translating"

        with make_progress() as progress:
            task = progress.add_task(
                task_description,
                total=slide_count,
                extra="",
            )
            callback = make_slide_progress_callback(progress, task, translator.engine.metrics, model_id)
            result = translator.translate_presentation(
                input_file, output_file, target_language, progress_callback=callback,
            )

    metrics_line = render_metrics_line(translator.engine.metrics, model_id)
    if result:
        click.echo(f"✅ Translation completed: {output_file}")
        click.echo(f"   {metrics_line}")
    else:
        click.echo("❌ Translation failed", err=True)
        sys.exit(1)


def parse_slide_numbers(slides_str):
    """Parse slide numbers string like '1,3,5' or '2-4' into list of integers"""
    slide_numbers = []
    for part in slides_str.split(','):
        part = part.strip()
        if '-' in part:
            start, end = map(int, part.split('-'))
            slide_numbers.extend(range(start, end + 1))
        else:
            slide_numbers.append(int(part))
    return slide_numbers


@cli.command('translate-slides')
@click.argument('input_file', type=click.Path(exists=True))
@click.option('-s', '--slides', required=True, help='Slide numbers (e.g., "1,3,5" or "2-4")')
@click.option('-t', '--target-language', default=Config.DEFAULT_TARGET_LANGUAGE, help='Target language')
@click.option('-o', '--output-file', help='Output file path')
@click.option('-m', '--model-id', default=Config.DEFAULT_MODEL_ID, help='Bedrock Mantle model ID')
@click.option('--no-polishing', is_flag=True, help='Disable natural language polishing')
@_slide_parallel_options
@_common_options
def translate_slides(input_file, slides, target_language, output_file, model_id, no_polishing,
                     slide_workers, slides_per_worker,
                     source_language, no_detect_source,
                     glossary_path, cache_backend, cache_path, no_cache, dry_run, no_charts):
    """Translate specific slides in PowerPoint presentation"""
    try:
        slide_numbers = parse_slide_numbers(slides)
    except ValueError:
        click.echo(f"❌ Invalid slide numbers format: {slides}", err=True)
        sys.exit(1)

    glossary = _resolve_glossary(glossary_path, target_language)
    backend = _resolve_cache_backend(cache_backend, no_cache)
    auto_detect = not no_detect_source

    if not output_file:
        input_path = Path(input_file)
        output_file = str(input_path.parent / f"{input_path.stem}_slides_{slides.replace(',', '_').replace('-', 'to')}_{target_language}{input_path.suffix}")

    if dry_run:
        translator = PowerPointTranslator(model_id, not no_polishing, glossary=glossary,
                                          translate_charts=not no_charts,
                                          source_language=source_language,
                                          auto_detect_source=auto_detect,
                                          slide_workers=slide_workers,
                                          slides_per_worker=slides_per_worker)
        _collect_for_dry_run(translator, input_file, target_language, model_id, slide_numbers,
                             detect_source=auto_detect)
        return

    click.echo(f"🚀 Starting translation of slides {slides}: {input_file} -> {target_language}")
    with build_cache(backend, cache_path) as cache:
        translator = PowerPointTranslator(model_id, not no_polishing, cache=cache,
                                          glossary=glossary, translate_charts=not no_charts,
                                          source_language=source_language,
                                          auto_detect_source=auto_detect,
                                          slide_workers=slide_workers,
                                          slides_per_worker=slides_per_worker)

        with make_progress() as progress:
            task = progress.add_task("Translating", total=len(slide_numbers), extra="")
            callback = make_slide_progress_callback(progress, task, translator.engine.metrics, model_id)
            result = translator.translate_specific_slides(
                input_file, output_file, target_language, slide_numbers,
                progress_callback=callback,
            )

    metrics_line = render_metrics_line(translator.engine.metrics, model_id)
    if result:
        click.echo(f"✅ Translation completed: {output_file}")
        click.echo(f"   {metrics_line}")
    else:
        click.echo("❌ Translation failed", err=True)
        sys.exit(1)


@cli.command("export-markdown")
@click.argument(
    "input_file",
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
)
@click.option("-o", "--output-file", type=click.Path(dir_okay=False, path_type=Path))
@click.option(
    "-l",
    "--language",
    default=Config.DEFAULT_TARGET_LANGUAGE,
    show_default=True,
    help="Markdown language code, or 'source' to preserve the source language",
)
@click.option(
    "-m",
    "--model-id",
    default=Config.DEFAULT_MODEL_ID,
    show_default=True,
    help="Bedrock Mantle model ID",
)
@click.option(
    "--mode",
    type=click.Choice(["structured", "extract"]),
    default="structured",
    show_default=True,
    help="AI-structured notes or deterministic content extraction",
)
@click.option(
    "--web-verify",
    is_flag=True,
    help="Verify selected current or external claims with client-side web search",
)
@click.option(
    "--max-web-queries",
    default=Config.MARKDOWN_MAX_WEB_QUERIES,
    type=click.IntRange(min=1, max=10),
    show_default=True,
)
@click.option(
    "--workers",
    default=Config.MARKDOWN_WORKERS,
    type=click.IntRange(min=1),
    show_default=True,
    help="Maximum concurrent Markdown slide chunks",
)
@click.option(
    "--slides-per-chunk",
    default=Config.MARKDOWN_SLIDES_PER_CHUNK,
    type=click.IntRange(min=1),
    show_default=True,
)
@click.option(
    "--cache-backend",
    type=click.Choice(["sqlite", "memory", "none"]),
    default=lambda: os.getenv("CACHE_BACKEND", "sqlite"),
    show_default="sqlite",
)
@click.option(
    "--cache-path",
    default=lambda: os.getenv(
        "CACHE_PATH",
        str(Path.home() / ".ppt-translator" / "cache.db"),
    ),
    show_default="~/.ppt-translator/cache.db",
)
@click.option("--no-cache", is_flag=True, help="Disable Markdown summary cache")
def export_markdown(
    input_file,
    output_file,
    language,
    model_id,
    mode,
    web_verify,
    max_web_queries,
    workers,
    slides_per_chunk,
    cache_backend,
    cache_path,
    no_cache,
):
    """Export one PowerPoint presentation as structured Markdown."""
    _validate_markdown_language(language)
    if web_verify and mode != "structured":
        raise click.UsageError("--web-verify requires --mode structured.")

    input_path = input_file.expanduser().resolve()
    if input_path.suffix.lower() != ".pptx":
        raise click.BadParameter(
            "Markdown export currently supports .pptx files only.",
            param_hint="input_file",
        )
    if output_file is None:
        suffix = "" if language == "source" else f"_{language}"
        output_file = input_path.with_name(f"{input_path.stem}{suffix}.md")

    slide_count = _presentation_slide_count(input_path)
    chunks = max(1, (slide_count + slides_per_chunk - 1) // slides_per_chunk)
    active_workers = min(workers, chunks)
    click.echo(f"📝 Markdown export: {input_path.name}")
    click.echo(f"📄 Slides: {slide_count}")
    click.echo(f"🌐 Output language: {language}")
    if mode == "structured":
        click.echo(
            f"🧵 Plan: {chunks} chunk(s), {slides_per_chunk} slides/chunk, "
            f"{active_workers} worker(s)"
        )
        click.echo(f"🔎 Web verification: {'ON' if web_verify else 'OFF'}")
    else:
        click.echo("📋 Mode: deterministic extraction (no model calls)")

    backend = _resolve_cache_backend(cache_backend, no_cache)
    try:
        with build_cache(backend, cache_path) as cache:
            exporter = MarkdownExporter(
                model_id=model_id,
                output_language=language,
                cache=cache,
                chunk_size=slides_per_chunk,
                workers=workers,
            )
            if mode == "structured":
                with make_progress() as progress:
                    task = progress.add_task(
                        "Structuring Markdown",
                        total=slide_count or 1,
                        extra="",
                    )

                    def _progress(completed, total):
                        progress.update(task, completed=completed, total=total or 1)

                    result = exporter.export(
                        input_path,
                        output_file,
                        mode=mode,
                        web_verify=web_verify,
                        max_web_queries=max_web_queries,
                        progress_callback=_progress,
                    )
            else:
                result = exporter.export(
                    input_path,
                    output_file,
                    mode=mode,
                    web_verify=False,
                )
    except Exception as exc:
        click.echo(f"❌ Markdown export failed: {exc}", err=True)
        raise click.ClickException(str(exc)) from exc

    click.echo(f"✅ Markdown created: {result.output_file}")
    click.echo(f"   {_markdown_metrics_line(result.metrics)}")


def _export_markdown_single_file(args):
    """Process worker for batch Markdown export."""
    quiet_batch_worker_logs()
    (
        input_file,
        output_file,
        language,
        model_id,
        mode,
        web_verify,
        max_web_queries,
        chunk_workers,
        slides_per_chunk,
        cache_backend,
        cache_path,
    ) = args
    try:
        with build_cache(cache_backend, cache_path) as cache:
            exporter = MarkdownExporter(
                model_id=model_id,
                output_language=language,
                cache=cache,
                chunk_size=slides_per_chunk,
                workers=chunk_workers,
            )
            result = exporter.export(
                input_file,
                output_file,
                mode=mode,
                web_verify=web_verify,
                max_web_queries=max_web_queries,
            )
        return (
            Path(input_file).name,
            str(result.output_file),
            True,
            None,
            result.metrics,
        )
    except Exception as exc:
        return (Path(input_file).name, None, False, str(exc), None)


@cli.command("batch-export-markdown")
@click.argument(
    "input_folder",
    type=click.Path(exists=True, file_okay=False, path_type=Path),
)
@click.option("-o", "--output-folder", type=click.Path(file_okay=False, path_type=Path))
@click.option(
    "-l",
    "--language",
    default=Config.DEFAULT_TARGET_LANGUAGE,
    show_default=True,
    help="Markdown language code, or 'source'",
)
@click.option(
    "-m",
    "--model-id",
    default=Config.DEFAULT_MODEL_ID,
    show_default=True,
    help="Bedrock Mantle model ID",
)
@click.option(
    "--mode",
    type=click.Choice(["structured", "extract"]),
    default="structured",
    show_default=True,
)
@click.option("--web-verify", is_flag=True)
@click.option(
    "--max-web-queries",
    default=Config.MARKDOWN_MAX_WEB_QUERIES,
    type=click.IntRange(min=1, max=10),
    show_default=True,
)
@click.option(
    "-w",
    "--workers",
    default=Config.BATCH_WORKERS,
    type=click.IntRange(min=1),
    show_default=True,
    help="Concurrent PowerPoint files",
)
@click.option(
    "--chunk-workers",
    default=Config.MARKDOWN_WORKERS,
    type=click.IntRange(min=1),
    show_default=True,
    help="Concurrent Markdown chunks per presentation",
)
@click.option(
    "--slides-per-chunk",
    default=Config.MARKDOWN_SLIDES_PER_CHUNK,
    type=click.IntRange(min=1),
    show_default=True,
)
@click.option("-r/-R", "--recursive/--no-recursive", default=True, show_default=True)
@click.option(
    "--cache-backend",
    type=click.Choice(["sqlite", "memory", "none"]),
    default=lambda: os.getenv("CACHE_BACKEND", "sqlite"),
    show_default="sqlite",
)
@click.option(
    "--cache-path",
    default=lambda: os.getenv(
        "CACHE_PATH",
        str(Path.home() / ".ppt-translator" / "cache.db"),
    ),
    show_default="~/.ppt-translator/cache.db",
)
@click.option("--no-cache", is_flag=True)
def batch_export_markdown(
    input_folder,
    output_folder,
    language,
    model_id,
    mode,
    web_verify,
    max_web_queries,
    workers,
    chunk_workers,
    slides_per_chunk,
    recursive,
    cache_backend,
    cache_path,
    no_cache,
):
    """Export every PowerPoint in a folder as Markdown."""
    _validate_markdown_language(language)
    if web_verify and mode != "structured":
        raise click.UsageError("--web-verify requires --mode structured.")

    input_path = input_folder.expanduser().resolve()
    output_path = (
        output_folder.expanduser().resolve()
        if output_folder is not None
        else input_path / f"markdown_{language}"
    )
    try:
        ppt_files = _discover_markdown_source_files(
            input_path,
            output_path,
            recursive,
        )
    except ValueError as exc:
        raise click.ClickException(str(exc)) from exc
    if not ppt_files:
        raise click.ClickException(f"No .pptx files found in {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    file_workers = min(workers, len(ppt_files))
    backend = _resolve_cache_backend(cache_backend, no_cache)
    tasks = []
    for ppt_file in ppt_files:
        relative = ppt_file.relative_to(input_path)
        output_file = output_path / relative.parent / f"{relative.stem}.md"
        output_file.parent.mkdir(parents=True, exist_ok=True)
        tasks.append((
            ppt_file,
            output_file,
            language,
            model_id,
            mode,
            web_verify,
            max_web_queries,
            chunk_workers,
            slides_per_chunk,
            backend,
            cache_path,
        ))

    click.echo(f"📁 Found {len(tasks)} PowerPoint file(s)")
    click.echo(f"📂 Output folder: {output_path}")
    click.echo(f"⚡ File workers: {file_workers}")
    click.echo(
        f"🧵 Chunk workers per presentation: {chunk_workers} "
        f"({slides_per_chunk} slides/chunk)"
    )
    click.echo(f"🔎 Web verification: {'ON' if web_verify else 'OFF'}")

    successes = 0
    failures: list[str] = []
    aggregate = {
        "api_calls": 0,
        "cache_hits": 0,
        "cache_misses": 0,
        "tokens_in": 0,
        "tokens_out": 0,
        "web_queries": 0,
        "web_results": 0,
    }
    with ProcessPoolExecutor(max_workers=file_workers) as executor:
        futures = {
            executor.submit(_export_markdown_single_file, task): task
            for task in tasks
        }
        with make_progress() as progress:
            task_id = progress.add_task(
                "Batch Markdown export",
                total=len(tasks),
                extra="",
            )
            for future in as_completed(futures):
                try:
                    filename, _, success, error, metrics = future.result()
                except Exception as exc:
                    filename = Path(futures[future][0]).name
                    success = False
                    error = f"Worker process failed: {exc}"
                    metrics = None
                if success:
                    successes += 1
                    if metrics is not None:
                        for key in aggregate:
                            aggregate[key] += getattr(metrics, key, 0)
                else:
                    failures.append(f"{filename}: {error}")
                progress.update(
                    task_id,
                    advance=1,
                    extra=(
                        f"success={successes}  failed={len(failures)}  "
                        f"calls={aggregate['api_calls']}"
                    ),
                )

    click.echo(f"✅ Markdown export completed: {successes}/{len(tasks)}")
    click.echo(f"   Output: {output_path}")
    click.echo(f"   {_markdown_metrics_line(type('Metrics', (), aggregate)())}")
    if failures:
        for failure in failures:
            click.echo(f"❌ {failure}", err=True)
        raise click.ClickException(f"{len(failures)} file(s) failed")


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
def info(input_file):
    """Show slide information and previews"""
    translator = PowerPointTranslator()

    try:
        slide_count = translator.get_slide_count(input_file)
        click.echo(f"📊 Presentation: {input_file}")
        click.echo(f"📄 Total slides: {slide_count}")
        click.echo()

        for i in range(1, min(slide_count + 1, 6)):  # Show first 5 slides
            preview = translator.get_slide_preview(input_file, i, max_chars=100)
            click.echo(f"Slide {i}:")
            if preview.strip():
                click.echo(f"  • {preview}")
            else:
                click.echo(f"  • (No text content)")
            click.echo()

        if slide_count > 5:
            click.echo(f"... and {slide_count - 5} more slides")

    except Exception as e:
        click.echo(f"❌ Error reading presentation: {e}", err=True)
        sys.exit(1)


def _batch_progress_listener(
    queue,
    progress,
    overall_task_id,
    per_file_task_ids=None,
):
    """Drain progress events from workers and reflect them in the parent UI.

    Runs in a daemon thread in the parent process. A `None` sentinel means
    "no more events, exit". File rows are pre-created before the live display
    starts so its height remains stable throughout the batch.
    """
    per_file_task_ids = dict(per_file_task_ids or {})
    while True:
        try:
            event = queue.get()
        except (EOFError, OSError):
            # Manager shut down — nothing more is coming.
            return
        if event is None:
            return
        try:
            etype = event.get('type')
            filename = event.get('file', '?')
            if etype == 'file_start':
                total = event.get('total_slides') or 0
                tid = per_file_task_ids.get(filename)
                if tid is None:
                    tid = progress.add_task(
                        filename,
                        total=total or 1,
                        extra="",
                    )
                    per_file_task_ids[filename] = tid
                elif total:
                    progress.update(tid, total=total)
            elif etype == 'slide_done':
                tid = per_file_task_ids.get(filename)
                if tid is not None:
                    # Render metrics straight from the event dict (supported by
                    # render_metrics_line's duck-typing).
                    extra = render_metrics_line(event, event.get('model_id', ''))
                    progress.update(tid, advance=1, extra=extra)
            elif etype == 'file_done':
                pass
            elif etype == 'file_complete':
                tid = per_file_task_ids.get(filename)
                if tid is not None:
                    status = "✅" if event.get('success') else "❌"
                    progress.update(tid, description=f"{status} {filename}")
                progress.update(overall_task_id, advance=1)
        except Exception as e:
            logger.debug(f"Progress listener event ignored: {e}")


def _translate_single_file(args):
    """Worker: build its own cache inside the child process.

    Each worker opens an independent SQLite connection. Sharing a connection
    across a fork boundary corrupts the DB.

    If a `progress_queue` is supplied (last tuple element), the worker emits
    file_start / slide_done / file_done events so the parent can render
    per-file slide progress bars.
    """
    quiet_batch_worker_logs()

    # Tuple layouts we support (grown over time, oldest first):
    #   9:  no source-lang, no queue       (legacy)
    #   11: +source_language, +auto_detect (no queue)
    #   12: +progress_queue                (legacy batch path)
    #   13: +slide parallel settings       (MCP batch path)
    #   14: +slide settings, +queue        (current CLI batch path)
    progress_queue = None
    if len(args) == 14:
        (ppt_file, output_file, target_language, model_id, enable_polishing,
         cache_backend, cache_path, glossary_path, translate_charts,
         source_language, auto_detect_source, slide_workers,
         slides_per_worker, progress_queue) = args
    elif len(args) == 13:
        (ppt_file, output_file, target_language, model_id, enable_polishing,
         cache_backend, cache_path, glossary_path, translate_charts,
         source_language, auto_detect_source, slide_workers,
         slides_per_worker) = args
    elif len(args) == 12:
        (ppt_file, output_file, target_language, model_id, enable_polishing,
         cache_backend, cache_path, glossary_path, translate_charts,
         source_language, auto_detect_source, progress_queue) = args
        slide_workers, slides_per_worker = 1, Config.SLIDES_PER_WORKER
    elif len(args) == 11:
        (ppt_file, output_file, target_language, model_id, enable_polishing,
         cache_backend, cache_path, glossary_path, translate_charts,
         source_language, auto_detect_source) = args
        slide_workers, slides_per_worker = 1, Config.SLIDES_PER_WORKER
    else:
        (ppt_file, output_file, target_language, model_id, enable_polishing,
         cache_backend, cache_path, glossary_path, translate_charts) = args
        source_language, auto_detect_source = None, True
        slide_workers, slides_per_worker = 1, Config.SLIDES_PER_WORKER

    filename = ppt_file.name

    def _emit(event: dict) -> None:
        # Queue ops are best-effort: a broken queue must not derail translation.
        if progress_queue is None:
            return
        try:
            progress_queue.put(event)
        except Exception:
            pass

    started = False
    try:
        glossary_map = load_glossary(glossary_path) if glossary_path else {}
        glossary = get_glossary_for_language(glossary_map, target_language)
        with build_cache(cache_backend, cache_path) as cache:
            translator = PowerPointTranslator(
                model_id, enable_polishing, cache=cache,
                glossary=glossary, translate_charts=translate_charts,
                source_language=source_language, auto_detect_source=auto_detect_source,
                slide_workers=slide_workers,
                slides_per_worker=slides_per_worker,
            )

            # Announce the file as soon as we know how many slides it has.
            try:
                total_slides = translator.get_slide_count(str(ppt_file))
            except Exception:
                total_slides = 0
            chunk_sizes, active_slide_workers = _slide_chunk_plan(
                total_slides,
                slide_workers,
                slides_per_worker,
            )
            _emit({
                'type': 'file_start',
                'file': filename,
                'total_slides': total_slides,
                'chunk_sizes': chunk_sizes,
                'active_slide_workers': active_slide_workers,
            })
            started = True

            def _slide_cb(slide_idx, metrics):
                # Convert the live metrics object into a pickle-safe dict.
                _emit({
                    'type': 'slide_done',
                    'file': filename,
                    'slide': slide_idx,
                    'cache_hits': getattr(metrics, 'cache_hits', 0),
                    'cache_misses': getattr(metrics, 'cache_misses', 0),
                    'tokens_in': getattr(metrics, 'tokens_in', 0),
                    'tokens_out': getattr(metrics, 'tokens_out', 0),
                    'model_id': model_id,
                })

            result = translator.translate_presentation(
                str(ppt_file), str(output_file), target_language,
                progress_callback=_slide_cb if progress_queue is not None else None,
            )
        return (filename, output_file.name, result, None)
    except Exception as e:
        return (filename, None, False, str(e))
    finally:
        # Always clear the file's row so the parent UI doesn't leak on errors.
        if started:
            _emit({'type': 'file_done', 'file': filename})


@cli.command('batch-translate')
@click.argument('input_folder', type=click.Path(exists=True, file_okay=False, dir_okay=True))
@click.option('-t', '--target-language', default=Config.DEFAULT_TARGET_LANGUAGE, help='Target language')
@click.option('-o', '--output-folder', help='Output folder path')
@click.option('-m', '--model-id', default=Config.DEFAULT_MODEL_ID, help='Bedrock Mantle model ID')
@click.option('--no-polishing', is_flag=True, help='Disable natural language polishing')
@click.option(
    '-w',
    '--workers',
    default=Config.BATCH_WORKERS,
    type=click.IntRange(min=1),
    show_default=True,
    help='Number of PowerPoint files translated concurrently',
)
@click.option('-r/-R', '--recursive/--no-recursive', default=True,
              help='Recursively process subfolders (default: enabled). Use -R/--no-recursive to limit to top level.')
@_slide_parallel_options
@_common_options
def batch_translate(input_folder, target_language, output_folder, model_id, no_polishing,
                    workers, recursive, slide_workers, slides_per_worker,
                    source_language, no_detect_source,
                    glossary_path, cache_backend, cache_path, no_cache, dry_run, no_charts):
    """Translate all PowerPoint files in a folder (parallel processing)"""
    input_path = Path(input_folder).expanduser().resolve()
    output_path = (
        Path(output_folder).expanduser().resolve()
        if output_folder
        else input_path / f"translated_{target_language}"
    )
    try:
        ppt_files = _discover_presentation_files(
            input_path,
            output_path,
            recursive,
        )
    except ValueError as exc:
        click.echo(f"❌ {exc}", err=True)
        sys.exit(1)

    if not ppt_files:
        search_type = "recursively" if recursive else "(top level only)"
        click.echo(f"❌ No PowerPoint files found {search_type} in {input_folder}", err=True)
        sys.exit(1)

    output_path.mkdir(parents=True, exist_ok=True)

    file_worker_count, per_file_slide_workers = _resolve_batch_workers(
        workers,
        slide_workers,
        len(ppt_files),
    )

    click.echo(f"📁 Found {len(ppt_files)} PowerPoint file(s)")
    click.echo(f"🌍 Target language: {target_language}")
    click.echo(f"📂 Output folder: {output_path}")
    click.echo(f"⚡ File workers: {file_worker_count}")
    click.echo(f"🔄 Recursive: {'ON' if recursive else 'OFF (top level only)'}")
    click.echo(
        f"🧵 Slide workers per presentation: {per_file_slide_workers} "
        f"({slides_per_worker} slides/chunk)"
    )
    click.echo()

    backend = _resolve_cache_backend(cache_backend, no_cache)

    # Resolve glossary path once so every worker loads the same file.
    resolved_glossary_path = glossary_path
    if resolved_glossary_path is None:
        default = find_default_glossary()
        if default is not None:
            resolved_glossary_path = str(default)

    auto_detect = not no_detect_source

    if dry_run:
        # Dry-run for batch: aggregate stats across all files, no parallelism needed.
        from .pricing import estimate_cost, estimate_tokens  # local to avoid reordering
        glossary_map = load_glossary(resolved_glossary_path) if resolved_glossary_path else {}
        glossary = get_glossary_for_language(glossary_map, target_language)
        total_chars = 0
        total_items = 0
        translator = PowerPointTranslator(model_id, not no_polishing,
                                          glossary=glossary, translate_charts=not no_charts,
                                          source_language=source_language,
                                          auto_detect_source=auto_detect)
        for ppt_file in ppt_files:
            # Only detect once (on the first file) — same source language is assumed across the folder.
            stats = translator.collect_all_texts(
                str(ppt_file),
                detect_source=auto_detect and translator.source_language is None,
            )
            total_chars += stats['total_chars']
            total_items += stats['translatable_items']
            if translator.source_language is None and stats.get('source_language'):
                translator.source_language = stats['source_language']
                translator.engine.source_language = stats['source_language']
        src_for_estimate = translator.source_language or source_language or 'en'
        tokens_in = estimate_tokens(total_chars, src_for_estimate)
        tokens_out = estimate_tokens(total_chars, target_language)
        cost = estimate_cost(tokens_in, tokens_out, model_id)
        click.echo(f"  Source language:     {translator.source_language or '(not detected)'}")
        click.echo(f"📊 Batch Dry-Run Report")
        click.echo(f"  Files:               {len(ppt_files)}")
        click.echo(f"  Translatable items:  {total_items}")
        click.echo(f"  Total characters:    {total_chars:,}")
        click.echo(f"  Estimated tokens in: ~{tokens_in:,}")
        click.echo(f"  Estimated tokens out:~{tokens_out:,}")
        if cost > 0:
            click.echo(f"  Estimated cost:      ${cost:.4f}")
        return

    success_count = 0
    failed_files = []

    file_plans = {}
    click.echo("📋 Translation plan:")
    for ppt_file in ppt_files:
        total_slides = _presentation_slide_count(ppt_file)
        chunk_sizes, active_slide_workers = _slide_chunk_plan(
            total_slides,
            per_file_slide_workers,
            slides_per_worker,
        )
        file_plans[ppt_file.name] = {
            "total_slides": total_slides,
            "chunk_sizes": chunk_sizes,
            "active_slide_workers": active_slide_workers,
        }
        if total_slides and len(chunk_sizes) > 1:
            chunk_text = "+".join(str(size) for size in chunk_sizes)
            click.echo(
                f"  🧵 {ppt_file.name}: {total_slides} slides → "
                f"{len(chunk_sizes)} chunks ({chunk_text}), "
                f"{active_slide_workers} slide workers"
            )
        elif total_slides:
            click.echo(
                f"  📄 {ppt_file.name}: {total_slides} slides → sequential"
            )
        else:
            click.echo(f"  📄 {ppt_file.name}: slide count pending")
    click.echo()

    progress = make_progress()
    overall_task_id = progress.add_task(
        "Batch translating",
        total=len(ppt_files),
        extra="",
    )
    per_file_task_ids = {
        filename: progress.add_task(
            filename,
            total=plan["total_slides"] or 1,
            extra="",
        )
        for filename, plan in file_plans.items()
    }

    # A Manager-backed queue is the cleanest way to stream per-slide events
    # from worker processes back into the parent's rich Progress. The Manager
    # owns a small helper process; that overhead is negligible next to Bedrock
    # latency and earns us a live multi-row UI.
    with multiprocessing.Manager() as mgr:
        progress_queue = mgr.Queue()
        with ProcessPoolExecutor(max_workers=file_worker_count) as executor:
            tasks = []
            for ppt_file in ppt_files:
                relative_path = ppt_file.relative_to(input_path)
                output_file = output_path / relative_path.parent / f"{relative_path.stem}_{target_language}{relative_path.suffix}"
                output_file.parent.mkdir(parents=True, exist_ok=True)
                tasks.append((
                    ppt_file, output_file, target_language, model_id, not no_polishing,
                    backend, cache_path, resolved_glossary_path, not no_charts,
                    source_language, auto_detect, per_file_slide_workers,
                    slides_per_worker, progress_queue,
                ))

            # Submitting starts the forked worker processes. Do this before
            # Rich starts its refresh thread or our listener starts, since
            # forking a multi-threaded parent can leave child locks corrupted.
            futures = {executor.submit(_translate_single_file, t): t for t in tasks}

            with progress:
                listener = threading.Thread(
                    target=_batch_progress_listener,
                    args=(
                        progress_queue,
                        progress,
                        overall_task_id,
                        per_file_task_ids,
                    ),
                    daemon=True,
                )
                listener.start()

                try:
                    for future in as_completed(futures):
                        task = futures[future]
                        try:
                            filename, output_name, result, error = future.result()
                        except Exception as exc:
                            filename = task[0].name
                            output_name = None
                            result = False
                            error = f"Worker process failed: {exc}"
                        if result:
                            success_count += 1
                        else:
                            failed_files.append(filename)
                        # Bump the overall bar via the listener so all Progress
                        # mutations happen on the same thread.
                        try:
                            progress_queue.put({
                                'type': 'file_complete',
                                'file': filename,
                                'success': bool(result),
                                'output': output_name,
                                'error': error,
                            })
                        except Exception:
                            # Listener gone (very unlikely) — fall back to direct update.
                            progress.update(overall_task_id, advance=1)
                finally:
                    # Signal the listener to drain and stop before Progress shuts down.
                    try:
                        progress_queue.put(None)
                    except Exception:
                        pass
                    listener.join(timeout=5)

    click.echo()
    click.echo("=" * 60)
    click.echo(f"✨ Batch translation completed!")
    click.echo(f"   Success: {success_count}/{len(ppt_files)}")
    if failed_files:
        click.echo(f"   Failed: {len(failed_files)}")
        for failed in failed_files:
            click.echo(f"     - {failed}")


if __name__ == '__main__':
    cli()
