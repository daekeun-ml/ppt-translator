#!/usr/bin/env python3
"""
FastMCP Server implementation for PowerPoint Translator
"""
import os
import sys
import logging
from pathlib import Path
from typing import Optional

# Add current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fastmcp import FastMCP
from config import Config
from ppt_handler import PowerPointTranslator
from post_processing import PowerPointPostProcessor

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastMCP server
mcp = FastMCP("PowerPoint Translator")

@mcp.tool()
def translate_powerpoint(
    input_file: str,
    target_language: str = Config.DEFAULT_TARGET_LANGUAGE,
    output_file: Optional[str] = None,
    model_id: str = Config.DEFAULT_MODEL_ID,
    enable_polishing: bool = True
) -> str:
    """
    Translate a PowerPoint presentation to the specified language.
    
    Args:
        input_file: Path to the input PowerPoint file (.pptx)
        target_language: Target language code (e.g., 'ko', 'ja', 'es', 'fr', 'de')
        output_file: Path to save the translated file (optional, auto-generated if not provided)
        model_id: AWS Bedrock model ID to use for translation
        enable_polishing: Enable natural language polishing for more fluent translation
    
    Returns:
        Success message with translation details
    """
    try:
        # Validate input file
        input_path = Path(input_file)
        if not input_path.exists():
            return f"❌ Error: File not found: {input_file}"
        
        if not input_path.suffix.lower() == '.pptx':
            return f"❌ Error: File must be a PowerPoint (.pptx) file: {input_file}"
        
        # Validate target language
        if target_language not in Config.LANGUAGE_MAP:
            available_langs = ', '.join(Config.LANGUAGE_MAP.keys())
            return f"❌ Error: Unsupported language '{target_language}'. Available: {available_langs}"
        
        # Generate output filename if not provided
        if not output_file:
            output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{input_path.suffix}")
        
        # Create translator and translate
        logger.info(f"Starting translation: {input_file} -> {target_language}")
        translator = PowerPointTranslator(model_id, enable_polishing)
        result = translator.translate_presentation(input_file, output_file, target_language)
        
        # Apply post-processing if enabled
        config = Config()
        post_processing_applied = False
        if config.get_bool('ENABLE_TEXT_AUTOFIT', True):
            try:
                verbose = config.get_bool('DEBUG', False)
                post_processor = PowerPointPostProcessor(config, verbose=verbose)
                # Overwrite the original output file instead of creating a new one
                final_output = post_processor.process_presentation(output_file, output_file)
                post_processing_applied = True
                logger.info("Post-processing applied: Text auto-fitting enabled")
            except Exception as e:
                logger.warning(f"Post-processing failed: {e}")
        
        # Format success message
        lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        translation_mode = "Natural/Polished" if enable_polishing else "Literal"
        post_processing_status = "✅ Applied" if post_processing_applied else "⚠️ Skipped"
        
        return f"""✅ PowerPoint translation completed successfully!

📁 Input file: {input_file}
📁 Output file: {output_file}
🌐 Target language: {target_language} ({lang_name})
🎨 Translation mode: {translation_mode}
🤖 Model: {model_id}
📝 Translated texts: {result.translated_count}
📋 Translated notes: {result.translated_notes_count}
📊 Total shapes processed: {result.total_shapes}
🔧 Post-processing: {post_processing_status}

💡 Translation features used:
• Intelligent batch processing for efficiency
• Context-aware translation for coherence
• Unified text frame processing
• Formatting preservation
• {'Natural language polishing for fluent output' if enable_polishing else 'Literal translation for accuracy'}"""
        
    except Exception as e:
        logger.error(f"Translation failed: {str(e)}")
        return f"❌ Translation failed: {str(e)}"

@mcp.tool()
def translate_specific_slides(
    input_file: str,
    slide_numbers: str,
    target_language: str = Config.DEFAULT_TARGET_LANGUAGE,
    output_file: Optional[str] = None,
    model_id: str = Config.DEFAULT_MODEL_ID,
    enable_polishing: bool = True
) -> str:
    """
    Translate specific slides in a PowerPoint presentation.
    
    Args:
        input_file: Path to the input PowerPoint file (.pptx)
        slide_numbers: Comma-separated slide numbers to translate (e.g., "1,3,5" or "2-4,7")
        target_language: Target language code (e.g., 'ko', 'ja', 'es', 'fr', 'de')
        output_file: Path to save the translated file (optional, auto-generated if not provided)
        model_id: AWS Bedrock model ID to use for translation
        enable_polishing: Enable natural language polishing for more fluent translation
    
    Returns:
        Success message with translation details
    """
    try:
        # Validate input file
        input_path = Path(input_file)
        if not input_path.exists():
            return f"❌ Error: File not found: {input_file}"
        
        if not input_path.suffix.lower() == '.pptx':
            return f"❌ Error: File must be a PowerPoint (.pptx) file: {input_file}"
        
        # Validate target language
        if target_language not in Config.LANGUAGE_MAP:
            available_langs = ', '.join(Config.LANGUAGE_MAP.keys())
            return f"❌ Error: Unsupported language '{target_language}'. Available: {available_langs}"
        
        # Parse slide numbers
        try:
            slide_list = []
            for part in slide_numbers.split(','):
                part = part.strip()
                if '-' in part:
                    # Handle range like "2-4"
                    start, end = map(int, part.split('-'))
                    slide_list.extend(range(start, end + 1))
                else:
                    # Handle single number
                    slide_list.append(int(part))
        except ValueError:
            return f"❌ Error: Invalid slide numbers format. Use comma-separated numbers or ranges (e.g., '1,3,5' or '2-4,7')"
        
        # Generate output filename if not provided
        if not output_file:
            # Create slides suffix with range format
            sorted_slides = sorted(set(slide_list))
            if len(sorted_slides) > 1 and sorted_slides[-1] - sorted_slides[0] == len(sorted_slides) - 1:
                # Consecutive range
                slides_suffix = f"_slides_range_{sorted_slides[0]}_{sorted_slides[-1]}"
            else:
                # Individual slides or non-consecutive
                slides_suffix = f"_slides_{'_'.join(map(str, sorted_slides))}"
            output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{slides_suffix}{input_path.suffix}")
        
        # Create translator and translate specific slides
        logger.info(f"Starting specific slides translation: {input_file} -> {target_language}")
        translator = PowerPointTranslator(model_id, enable_polishing)
        result = translator.translate_specific_slides(input_file, output_file, target_language, slide_list)
        
        # Check for errors
        if result.errors:
            return f"❌ Translation failed: {'; '.join(result.errors)}"
        
        # Apply post-processing if enabled
        config = Config()
        post_processing_applied = False
        if config.get_bool('ENABLE_TEXT_AUTOFIT', True):
            try:
                verbose = config.get_bool('DEBUG', False)
                post_processor = PowerPointPostProcessor(config, verbose=verbose)
                # Overwrite the original output file instead of creating a new one
                final_output = post_processor.process_presentation(output_file, output_file)
                post_processing_applied = True
                logger.info("Post-processing applied: Text auto-fitting enabled")
            except Exception as e:
                logger.warning(f"Post-processing failed: {e}")
        
        # Format success message
        lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        translation_mode = "Natural/Polished" if enable_polishing else "Literal"
        post_processing_status = "✅ Applied" if post_processing_applied else "⚠️ Skipped"
        
        return f"""✅ Specific slides translation completed successfully!

📁 Input file: {input_file}
📁 Output file: {output_file}
📄 Translated slides: {sorted(set(slide_list))}
🌐 Target language: {target_language} ({lang_name})
🎨 Translation mode: {translation_mode}
🤖 Model: {model_id}
📝 Translated texts: {result.translated_count}
📋 Translated notes: {result.translated_notes_count}
📊 Total shapes processed: {result.total_shapes}
🔧 Post-processing: {post_processing_status}

💡 Translation features used:
• Intelligent batch processing for efficiency
• Context-aware translation for coherence
• Unified text frame processing
• Formatting preservation
• {'Natural language polishing for fluent output' if enable_polishing else 'Literal translation for accuracy'}"""
        
    except Exception as e:
        logger.error(f"Specific slides translation failed: {str(e)}")
        return f"❌ Translation failed: {str(e)}"

@mcp.tool()
def get_slide_info(input_file: str) -> str:
    """
    Get information about slides in a PowerPoint presentation.
    
    Args:
        input_file: Path to the PowerPoint file (.pptx)
    
    Returns:
        Information about the presentation including slide count and preview of each slide
    """
    try:
        # Validate input file
        input_path = Path(input_file)
        if not input_path.exists():
            return f"❌ Error: File not found: {input_file}"
        
        if not input_path.suffix.lower() == '.pptx':
            return f"❌ Error: File must be a PowerPoint (.pptx) file: {input_file}"
        
        # Create translator to access slide info methods
        translator = PowerPointTranslator()
        slide_count = translator.get_slide_count(input_file)
        
        info_text = f"""📊 PowerPoint Presentation Information

📁 File: {input_file}
📄 Total slides: {slide_count}

📋 Slide previews:
"""
        
        # Get preview for each slide (limit to first 10 slides for readability)
        max_preview_slides = min(slide_count, 10)
        for i in range(1, max_preview_slides + 1):
            try:
                preview = translator.get_slide_preview(input_file, i, max_chars=150)
                info_text += f"\n🔸 Slide {i}: {preview}"
            except Exception as e:
                info_text += f"\n🔸 Slide {i}: [Error getting preview: {str(e)}]"
        
        if slide_count > 10:
            info_text += f"\n\n... and {slide_count - 10} more slides"
        
        info_text += f"""

💡 Usage examples:
• Translate all slides: translate_powerpoint("{input_file}")
• Translate specific slides: translate_specific_slides("{input_file}", "1,3,5")
• Translate slide range: translate_specific_slides("{input_file}", "2-4")"""
        
        return info_text
        
    except Exception as e:
        logger.error(f"Failed to get slide info: {str(e)}")
        return f"❌ Failed to get slide info: {str(e)}"

@mcp.tool()
def get_slide_preview(input_file: str, slide_number: int) -> str:
    """
    Get a detailed preview of a specific slide's content.
    
    Args:
        input_file: Path to the PowerPoint file (.pptx)
        slide_number: Slide number to preview (1-based indexing)
    
    Returns:
        Detailed preview of the slide content
    """
    try:
        # Validate input file
        input_path = Path(input_file)
        if not input_path.exists():
            return f"❌ Error: File not found: {input_file}"
        
        if not input_path.suffix.lower() == '.pptx':
            return f"❌ Error: File must be a PowerPoint (.pptx) file: {input_file}"
        
        # Create translator and get preview
        translator = PowerPointTranslator()
        slide_count = translator.get_slide_count(input_file)
        
        if slide_number < 1 or slide_number > slide_count:
            return f"❌ Error: Invalid slide number {slide_number}. Valid range: 1-{slide_count}"
        
        preview = translator.get_slide_preview(input_file, slide_number, max_chars=500)
        
        return f"""📄 Slide {slide_number} Preview

📁 File: {input_file}
📊 Total slides: {slide_count}

📝 Content preview:
{preview}

💡 To translate this slide:
translate_specific_slides("{input_file}", "{slide_number}")"""
        
    except Exception as e:
        logger.error(f"Failed to get slide preview: {str(e)}")
        return f"❌ Failed to get slide preview: {str(e)}"

@mcp.tool()
def list_supported_languages() -> str:
    """
    List all supported target languages for translation.
    
    Returns:
        List of supported language codes and names
    """
    languages_text = "🌐 Supported target languages:\n\n"
    for code, name in sorted(Config.LANGUAGE_MAP.items()):
        languages_text += f"• {code}: {name}\n"
    
    return languages_text

@mcp.tool()
def list_supported_models() -> str:
    """
    List all supported AWS Bedrock models for translation.
    
    Returns:
        List of supported model IDs
    """
    models_text = "🤖 Supported AWS Bedrock models:\n\n"
    for model in Config.SUPPORTED_MODELS:
        models_text += f"• {model}\n"
    
    return models_text

@mcp.tool()
def get_translation_help() -> str:
    """
    Get help information about using the PowerPoint translator.
    
    Returns:
        Help text with usage examples
    """
    return """📖 PowerPoint Translator Help

🎯 Main Functions:
• translate_powerpoint() - Translate entire PowerPoint presentation
• translate_specific_slides() - Translate only specific slides
• get_slide_info() - Get presentation overview and slide previews
• get_slide_preview() - Get detailed preview of a specific slide

📋 Required Parameters:
• input_file: Path to your .pptx file

🔧 Optional Parameters:
• target_language: Language code (default: 'ko' for Korean)
• output_file: Output path (auto-generated if not specified)
• model_id: Bedrock model (default: Claude 3.7 Sonnet)
• enable_polishing: Natural translation vs literal (default: true)

💡 Usage Examples:

1. Get presentation information:
   get_slide_info("presentation.pptx")

2. Preview specific slide:
   get_slide_preview("presentation.pptx", 3)

3. Translate entire presentation:
   translate_powerpoint("presentation.pptx")

4. Translate specific slides (individual):
   translate_specific_slides("slides.pptx", "1,3,5")

5. Translate slide range:
   translate_specific_slides("slides.pptx", "2-4")

6. Translate mixed (individual + range):
   translate_specific_slides("slides.pptx", "1,3-5,8")

7. Translate to Spanish with custom output:
   translate_specific_slides("slides.pptx", "1-3", "es", "spanish_slides.pptx")

8. Literal translation (no polishing):
   translate_specific_slides("doc.pptx", "2,4", "ja", enable_polishing=False)

🌐 Get supported languages:
   list_supported_languages()

🤖 Get supported models:
   list_supported_models()

⚙️ Configuration:
• AWS credentials must be configured (aws configure)
• Bedrock access required in your AWS account
• Supported file format: .pptx only

📄 Slide Number Format:
• Individual slides: "1,3,5"
• Ranges: "2-4" (translates slides 2, 3, 4)
• Mixed: "1,3-5,8" (translates slides 1, 3, 4, 5, 8)"""


@mcp.tool()
def post_process_powerpoint(
    input_file: str,
    output_file: Optional[str] = None,
    text_threshold: Optional[int] = None,
    enable_autofit: bool = True
) -> str:
    """
    Apply post-processing to a PowerPoint presentation to optimize text boxes.
    
    This function enables text wrapping and shrink text on overflow for text boxes
    that contain text longer than the specified threshold.
    
    Args:
        input_file: Path to the input PowerPoint file (.pptx)
        output_file: Path to save the processed file (optional, auto-generated if not provided)
        text_threshold: Text length threshold for enabling auto-fit (overrides .env setting)
        enable_autofit: Enable text auto-fitting (default: True)
    
    Returns:
        Success message with post-processing details
    """
    try:
        # Validate input file
        input_path = Path(input_file)
        if not input_path.exists():
            return f"❌ Error: File not found: {input_file}"
        
        if not input_path.suffix.lower() == '.pptx':
            return f"❌ Error: Only .pptx files are supported. Got: {input_path.suffix}"
        
        # Create configuration
        config = Config()
        if text_threshold is not None:
            config.set('TEXT_LENGTH_THRESHOLD', str(text_threshold))
        if not enable_autofit:
            config.set('ENABLE_TEXT_AUTOFIT', 'false')
        
        # Generate output filename if not provided
        if not output_file:
            output_file = input_file  # Overwrite the original file
        
        # Apply post-processing
        logger.info(f"Starting post-processing: {input_file}")
        verbose = config.get_bool('DEBUG', False)
        post_processor = PowerPointPostProcessor(config, verbose=verbose)
        final_output = post_processor.process_presentation(input_file, output_file)
        
        threshold = config.get_int('TEXT_LENGTH_THRESHOLD', 10)
        autofit_enabled = config.get_bool('ENABLE_TEXT_AUTOFIT', True)
        
        return f"""✅ PowerPoint post-processing completed successfully!

📁 Input file: {input_file}
📁 Output file: {final_output}
🔧 Text auto-fitting: {'✅ Enabled' if autofit_enabled else '❌ Disabled'}
📏 Text length threshold: {threshold} characters
📝 Processing applied to text boxes longer than {threshold} characters

💡 Post-processing features applied:
• Text wrapping in shape enabled
• Shrink text on overflow enabled
• Text box margins optimized
• Formatting preservation maintained"""
        
    except Exception as e:
        logger.error(f"Post-processing failed: {str(e)}")
        return f"❌ Post-processing failed: {str(e)}"


def main():
    """Main entry point for the FastMCP server."""
    # Run the FastMCP server
    mcp.run()

if __name__ == "__main__":
    main()
