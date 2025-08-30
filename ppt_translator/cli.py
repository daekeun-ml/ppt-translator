#!/usr/bin/env python3
"""PowerPoint Translator CLI using Click"""

import click
import sys
import logging
from pathlib import Path

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


if __name__ == '__main__':
    cli()
