#!/usr/bin/env python3
"""
PowerPoint Translator Server
Supports both standalone translation and FastMCP server mode
"""
import argparse
import sys
import logging
import os
from pathlib import Path

# # Ensure the script can find its modules regardless of where it's called from
# script_dir = Path(__file__).parent.absolute()
# os.chdir(script_dir)
# sys.path.insert(0, str(script_dir))

from config import Config
from ppt_handler import PowerPointTranslator

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def translate_standalone(input_file: str, target_language: str, output_file: str = None, 
                        model_id: str = Config.DEFAULT_MODEL_ID, enable_polishing: bool = True):
    """Standalone translation function"""
    try:
        # Validate input file
        if not Path(input_file).exists():
            logger.error(f"Input file not found: {input_file}")
            return False
        
        # Generate output filename if not provided
        if not output_file:
            input_path = Path(input_file)
            polishing_suffix = "_polished" if enable_polishing else "_literal"
            output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{polishing_suffix}{input_path.suffix}")
        
        # Create translator and translate
        translator = PowerPointTranslator(model_id, enable_polishing)
        result = translator.translate_presentation(input_file, output_file, target_language)
        
        # Print results
        lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        translation_mode = "Natural/Polished" if enable_polishing else "Literal"
        
        print(f"""
✅ Translation completed successfully!

📁 Input file: {input_file}
📁 Output file: {output_file}
🌐 Target language: {target_language} ({lang_name})
🎨 Translation mode: {translation_mode}
🤖 Model: {model_id}
📝 Translated texts: {result.translated_count}
📋 Translated notes: {result.translated_notes_count}
📊 Total shapes processed: {result.total_shapes}
        """)
        
        return True
        
    except Exception as e:
        logger.error(f"Translation failed: {str(e)}")
        return False


def translate_specific_slides_standalone(input_file: str, slide_numbers: str, target_language: str, 
                                        output_file: str = None, model_id: str = Config.DEFAULT_MODEL_ID, 
                                        enable_polishing: bool = True):
    """Standalone specific slides translation function"""
    try:
        # Validate input file
        if not Path(input_file).exists():
            logger.error(f"Input file not found: {input_file}")
            return False
        
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
            logger.error("Invalid slide numbers format. Use comma-separated numbers or ranges (e.g., '1,3,5' or '2-4,7')")
            return False
        
        # Generate output filename if not provided
        if not output_file:
            input_path = Path(input_file)
            polishing_suffix = "_polished" if enable_polishing else "_literal"
            slides_suffix = f"_slides_{'_'.join(map(str, sorted(set(slide_list))))}"
            output_file = str(input_path.parent / f"{input_path.stem}_translated_{target_language}{slides_suffix}{polishing_suffix}{input_path.suffix}")
        
        # Create translator and translate specific slides
        translator = PowerPointTranslator(model_id, enable_polishing)
        result = translator.translate_specific_slides(input_file, output_file, target_language, slide_list)
        
        # Check for errors
        if result.errors:
            logger.error(f"Translation failed: {'; '.join(result.errors)}")
            return False
        
        # Print results
        lang_name = Config.LANGUAGE_MAP.get(target_language, target_language)
        translation_mode = "Natural/Polished" if enable_polishing else "Literal"
        
        print(f"""
✅ Specific slides translation completed successfully!

📁 Input file: {input_file}
📁 Output file: {output_file}
📄 Translated slides: {sorted(set(slide_list))}
🌐 Target language: {target_language} ({lang_name})
🎨 Translation mode: {translation_mode}
🤖 Model: {model_id}
📝 Translated texts: {result.translated_count}
📋 Translated notes: {result.translated_notes_count}
📊 Total shapes processed: {result.total_shapes}
        """)
        
        return True
        
    except Exception as e:
        logger.error(f"Translation failed: {str(e)}")
        return False


def show_slide_info(input_file: str):
    """Show slide information"""
    try:
        # Validate input file
        if not Path(input_file).exists():
            logger.error(f"Input file not found: {input_file}")
            return False
        
        # Create translator and get slide info
        translator = PowerPointTranslator()
        slide_count = translator.get_slide_count(input_file)
        
        print(f"""
📊 PowerPoint Presentation Information

📁 File: {input_file}
📄 Total slides: {slide_count}

📋 Slide previews:
        """)
        
        # Get preview for each slide (limit to first 10 slides for readability)
        max_preview_slides = min(slide_count, 10)
        for i in range(1, max_preview_slides + 1):
            try:
                preview = translator.get_slide_preview(input_file, i, max_chars=150)
                print(f"🔸 Slide {i}: {preview}")
            except Exception as e:
                print(f"🔸 Slide {i}: [Error getting preview: {str(e)}]")
        
        if slide_count > 10:
            print(f"\n... and {slide_count - 10} more slides")
        
        print(f"""
💡 Usage examples:
• Translate all slides: python server.py --translate -i "{input_file}"
• Translate specific slides: python server.py --translate-slides "1,3,5" -i "{input_file}"
• Translate slide range: python server.py --translate-slides "2-4" -i "{input_file}"
        """)
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to get slide info: {str(e)}")
        return False


def run_fastmcp_server():
    """Run FastMCP server"""
    try:
        # Import and run the FastMCP server
        import fastmcp_server
        # The server will run when imported due to the if __name__ == "__main__" block
        
    except ImportError as e:
        logger.error(f"FastMCP server dependencies not available: {e}")
        logger.error("Please install fastmcp: pip install fastmcp")
        sys.exit(1)
    except Exception as e:
        logger.error(f"FastMCP server failed: {e}")
        sys.exit(1)


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(description="PowerPoint Translator")
    
    # Mode selection
    mode_group = parser.add_mutually_exclusive_group(required=True)
    mode_group.add_argument("--translate", action="store_true", help="Run standalone translation (all slides)")
    mode_group.add_argument("--translate-slides", metavar="SLIDES", help="Translate specific slides (e.g., '1,3,5' or '2-4')")
    mode_group.add_argument("--slide-info", action="store_true", help="Show slide information and previews")
    mode_group.add_argument("--mcp", action="store_true", help="Run FastMCP server")
    
    # Translation arguments
    parser.add_argument("--input-file", "-i", help="Input PowerPoint file")
    parser.add_argument("--output-file", "-o", help="Output PowerPoint file")
    parser.add_argument("--target-language", "-t", default=Config.DEFAULT_TARGET_LANGUAGE,
                       help=f"Target language (default: {Config.DEFAULT_TARGET_LANGUAGE})")
    parser.add_argument("--model-id", "-m", default=Config.DEFAULT_MODEL_ID,
                       help=f"Bedrock model ID (default: {Config.DEFAULT_MODEL_ID})")
    parser.add_argument("--no-polishing", action="store_true", 
                       help="Disable natural language polishing (use literal translation)")
    
    # Debug options
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    parser.add_argument("--test", action="store_true", help="Test FastMCP server")
    
    args = parser.parse_args()
    
    # Set debug logging
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Handle modes
    if args.translate:
        if not args.input_file:
            parser.error("--input-file is required for translation mode")
        
        success = translate_standalone(
            args.input_file,
            args.target_language,
            args.output_file,
            args.model_id,
            not args.no_polishing
        )
        sys.exit(0 if success else 1)
    
    elif args.translate_slides:
        if not args.input_file:
            parser.error("--input-file is required for specific slides translation mode")
        
        success = translate_specific_slides_standalone(
            args.input_file,
            args.translate_slides,
            args.target_language,
            args.output_file,
            args.model_id,
            not args.no_polishing
        )
        sys.exit(0 if success else 1)
    
    elif args.slide_info:
        if not args.input_file:
            parser.error("--input-file is required for slide info mode")
        
        success = show_slide_info(args.input_file)
        sys.exit(0 if success else 1)
    
    elif args.mcp:
        if args.test:
            print("Testing FastMCP server...")
            # Simple test - just try to import fastmcp
            try:
                import fastmcp
                print("✅ FastMCP server test passed")
                sys.exit(0)
            except Exception as e:
                print(f"❌ FastMCP server test failed: {e}")
                sys.exit(1)
        else:
            run_fastmcp_server()


if __name__ == "__main__":
    main()
