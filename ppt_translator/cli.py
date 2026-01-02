#!/usr/bin/env python3
"""PowerPoint Translator CLI using Click"""

import click
import sys
import logging
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

from .config import Config
from .ppt_handler import PowerPointTranslator
from .post_processing import PowerPointPostProcessor

# Configure logging to show detailed INFO messages
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)


@click.group()
@click.version_option()
def cli():
    """PowerPoint Translator using Amazon Bedrock"""
    pass


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('-t', '--target-language', default=Config.DEFAULT_TARGET_LANGUAGE, help='Target language')
@click.option('-o', '--output-file', help='Output file path')
@click.option('-m', '--model-id', default=Config.DEFAULT_MODEL_ID, help='Bedrock model ID')
@click.option('--no-polishing', is_flag=True, help='Disable natural language polishing')
def translate(input_file, target_language, output_file, model_id, no_polishing):
    """Translate entire PowerPoint presentation"""
    if not output_file:
        input_path = Path(input_file)
        output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{input_path.suffix}")
    
    click.echo(f"🚀 Starting translation: {input_file} -> {target_language}")
    
    translator = PowerPointTranslator(model_id, not no_polishing)
    result = translator.translate_presentation(input_file, output_file, target_language)
    
    if result:
        click.echo(f"✅ Translation completed: {output_file}")
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


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('-s', '--slides', required=True, help='Slide numbers (e.g., "1,3,5" or "2-4")')
@click.option('-t', '--target-language', default=Config.DEFAULT_TARGET_LANGUAGE, help='Target language')
@click.option('-o', '--output-file', help='Output file path')
@click.option('-m', '--model-id', default=Config.DEFAULT_MODEL_ID, help='Bedrock model ID')
@click.option('--no-polishing', is_flag=True, help='Disable natural language polishing')
def translate_slides(input_file, slides, target_language, output_file, model_id, no_polishing):
    """Translate specific slides in PowerPoint presentation"""
    try:
        slide_numbers = parse_slide_numbers(slides)
    except ValueError as e:
        click.echo(f"❌ Invalid slide numbers format: {slides}", err=True)
        sys.exit(1)
    
    if not output_file:
        input_path = Path(input_file)
        output_file = str(input_path.parent / f"{input_path.stem}_slides_{slides.replace(',', '_').replace('-', 'to')}_{target_language}{input_path.suffix}")
    
    click.echo(f"🚀 Starting translation of slides {slides}: {input_file} -> {target_language}")
    
    translator = PowerPointTranslator(model_id, not no_polishing)
    result = translator.translate_specific_slides(input_file, output_file, target_language, slide_numbers)
    
    if result:
        click.echo(f"✅ Translation completed: {output_file}")
    else:
        click.echo("❌ Translation failed", err=True)
        sys.exit(1)


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


def _translate_single_file(args):
    """Helper function for parallel processing"""
    ppt_file, output_file, target_language, model_id, enable_polishing = args
    try:
        translator = PowerPointTranslator(model_id, enable_polishing)
        result = translator.translate_presentation(str(ppt_file), str(output_file), target_language)
        return (ppt_file.name, output_file.name, result, None)
    except Exception as e:
        return (ppt_file.name, None, False, str(e))


@cli.command()
@click.argument('input_folder', type=click.Path(exists=True, file_okay=False, dir_okay=True))
@click.option('-t', '--target-language', default=Config.DEFAULT_TARGET_LANGUAGE, help='Target language')
@click.option('-o', '--output-folder', help='Output folder path')
@click.option('-m', '--model-id', default=Config.DEFAULT_MODEL_ID, help='Bedrock model ID')
@click.option('--no-polishing', is_flag=True, help='Disable natural language polishing')
@click.option('-w', '--workers', default=4, type=int, help='Number of parallel workers (default: 4)')
@click.option('-r', '--recursive', is_flag=True, help='Recursively process subfolders')
def batch_translate(input_folder, target_language, output_folder, model_id, no_polishing, workers, recursive):
    """Translate all PowerPoint files in a folder (parallel processing)"""
    input_path = Path(input_folder)
    output_path = Path(output_folder) if output_folder else input_path / f"translated_{target_language}"
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Find PowerPoint files (recursive or non-recursive)
    if recursive:
        ppt_files = list(input_path.rglob("*.pptx")) + list(input_path.rglob("*.ppt"))
    else:
        ppt_files = list(input_path.glob("*.pptx")) + list(input_path.glob("*.ppt"))
    
    if not ppt_files:
        search_type = "recursively" if recursive else ""
        click.echo(f"❌ No PowerPoint files found {search_type} in {input_folder}", err=True)
        sys.exit(1)
    
    click.echo(f"📁 Found {len(ppt_files)} PowerPoint file(s)")
    click.echo(f"🌍 Target language: {target_language}")
    click.echo(f"📂 Output folder: {output_path}")
    click.echo(f"⚡ Workers: {workers}")
    if recursive:
        click.echo("🔄 Recursive mode: ON")
    click.echo()
    
    # Prepare tasks with relative path preservation
    tasks = []
    for ppt_file in ppt_files:
        # Preserve folder structure in output
        relative_path = ppt_file.relative_to(input_path)
        output_file = output_path / relative_path.parent / f"{relative_path.stem}_{target_language}{relative_path.suffix}"
        output_file.parent.mkdir(parents=True, exist_ok=True)
        tasks.append((ppt_file, output_file, target_language, model_id, not no_polishing))
    
    success_count = 0
    failed_files = []
    completed = 0
    
    # Process with continuous batching
    with ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {}
        task_iter = iter(tasks)
        
        # Submit initial batch
        for _ in range(min(workers, len(tasks))):
            task = next(task_iter, None)
            if task:
                futures[executor.submit(_translate_single_file, task)] = task
        
        # Process as completed and submit new tasks
        while futures:
            done, _ = as_completed(futures), None
            for future in done:
                completed += 1
                task = futures.pop(future)
                filename, output_name, result, error = future.result()
                
                if result:
                    click.echo(f"[{completed}/{len(ppt_files)}] ✅ Completed: {output_name}")
                    success_count += 1
                else:
                    error_msg = f" - {error}" if error else ""
                    click.echo(f"[{completed}/{len(ppt_files)}] ❌ Failed: {filename}{error_msg}")
                    failed_files.append(filename)
                
                # Submit next task
                next_task = next(task_iter, None)
                if next_task:
                    futures[executor.submit(_translate_single_file, next_task)] = next_task
                
                break  # Process one at a time
    
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
