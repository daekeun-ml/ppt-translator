"""
PowerPoint document handling and text frame updates
"""
import logging
import re
from typing import List, Dict, Any, Tuple
from dataclasses import dataclass
from pathlib import Path
from pptx.dml.color import RGBColor
from config import Config
from dependencies import DependencyManager
from translation_engine import TranslationEngine
from text_utils import SlideTextCollector

logger = logging.getLogger(__name__)


@dataclass
class TranslationResult:
    """Data class for translation results"""
    translated_count: int = 0
    translated_notes_count: int = 0
    total_shapes: int = 0
    errors: List[str] = None
    
    def __post_init__(self):
        if self.errors is None:
            self.errors = []


class TextFrameUpdater:
    """Handles updating PowerPoint text frames with translations"""
    
    @staticmethod
    def update_text_frame(text_frame, new_text: str):
        """Update text frame while preserving formatting safely"""
        try:
            if not text_frame.paragraphs:
                text_frame.text = new_text
                return
            
            # Check for hyperlinks first
            has_hyperlinks = TextFrameUpdater._has_hyperlinks(text_frame)
            
            if has_hyperlinks:
                logger.debug("Hyperlinks detected, using safe hyperlink preservation")
                TextFrameUpdater._update_with_hyperlinks_safe(text_frame, new_text)
                return
            
            # For single paragraph, preserve formatting
            if len(text_frame.paragraphs) == 1 and '\n' not in new_text.strip():
                TextFrameUpdater._update_single_paragraph_safe(text_frame.paragraphs[0], new_text.strip())
                return
            
            # For multiple paragraphs, try to preserve structure
            new_lines = new_text.strip().split('\n')
            
            if len(new_lines) == len(text_frame.paragraphs):
                # Same number of lines - update in place
                for paragraph, new_line in zip(text_frame.paragraphs, new_lines):
                    if new_line.strip():
                        TextFrameUpdater._update_single_paragraph_safe(paragraph, new_line.strip())
            else:
                # Different structure - preserve first paragraph's formatting
                TextFrameUpdater._update_with_preserved_formatting_safe(text_frame, new_text)
                
        except Exception as e:
            logger.error(f"Formatting error: {str(e)}")
            text_frame.text = new_text
    
    @staticmethod
    def _has_hyperlinks(text_frame):
        """Check if text frame contains hyperlinks"""
        try:
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    if hasattr(run, 'hyperlink') and run.hyperlink:
                        if hasattr(run.hyperlink, 'address') and run.hyperlink.address:
                            return True
        except Exception:
            pass
        return False
    
    @staticmethod
    def _update_with_hyperlinks_safe(text_frame, new_text: str):
        """Update text frame while preserving hyperlinks safely"""
        try:
            # Extract hyperlink information safely
            hyperlink_info = []
            
            for paragraph in text_frame.paragraphs:
                para_info = {
                    'text': paragraph.text,
                    'hyperlinks': []
                }
                
                for run in paragraph.runs:
                    if run.text.strip():
                        run_info = {
                            'text': run.text,
                            'hyperlink': None,
                            'formatting': TextFrameUpdater._extract_run_formatting_safe(run)
                        }
                        
                        try:
                            if hasattr(run, 'hyperlink') and run.hyperlink:
                                if hasattr(run.hyperlink, 'address') and run.hyperlink.address:
                                    run_info['hyperlink'] = run.hyperlink.address
                        except Exception:
                            pass
                        
                        para_info['hyperlinks'].append(run_info)
                
                hyperlink_info.append(para_info)
            
            # Apply new text while preserving hyperlinks
            new_lines = new_text.strip().split('\n')
            
            for i, line in enumerate(new_lines):
                if i < len(text_frame.paragraphs):
                    paragraph = text_frame.paragraphs[i]
                else:
                    paragraph = text_frame.add_paragraph()
                
                # Clear paragraph
                paragraph.clear()
                
                # Apply hyperlinks if available
                if i < len(hyperlink_info):
                    TextFrameUpdater._apply_hyperlinks_to_paragraph_safe(paragraph, line.strip(), hyperlink_info[i])
                else:
                    # No hyperlink info, just add text
                    run = paragraph.add_run()
                    run.text = line.strip()
                    
        except Exception as e:
            logger.error(f"Safe hyperlink preservation failed: {e}")
            text_frame.text = new_text
    
    @staticmethod
    def _apply_hyperlinks_to_paragraph_safe(paragraph, line: str, para_info):
        """Apply hyperlinks to paragraph safely"""
        try:
            # Find runs with hyperlinks
            hyperlink_runs = [run for run in para_info['hyperlinks'] if run['hyperlink']]
            
            if not hyperlink_runs:
                # No hyperlinks, just add text
                run = paragraph.add_run()
                run.text = line
                if para_info['hyperlinks']:
                    TextFrameUpdater._apply_run_formatting_safe(run, para_info['hyperlinks'][0]['formatting'])
                return
            
            # Try to preserve hyperlinks
            remaining_text = line
            
            for hyperlink_run in hyperlink_runs:
                original_text = hyperlink_run['text'].strip()
                hyperlink_url = hyperlink_run['hyperlink']
                
                # Find hyperlink text in translated line
                hyperlink_text = TextFrameUpdater._find_hyperlink_text_safe(remaining_text, original_text)
                
                if hyperlink_text and hyperlink_text in remaining_text:
                    parts = remaining_text.split(hyperlink_text, 1)
                    
                    # Text before hyperlink
                    if parts[0]:
                        run = paragraph.add_run()
                        run.text = parts[0]
                        # Use default formatting
                        default_formatting = next((r['formatting'] for r in para_info['hyperlinks'] if not r['hyperlink']), 
                                                para_info['hyperlinks'][0]['formatting'] if para_info['hyperlinks'] else {})
                        TextFrameUpdater._apply_run_formatting_safe(run, default_formatting)
                    
                    # Hyperlink text
                    run = paragraph.add_run()
                    run.text = hyperlink_text
                    TextFrameUpdater._apply_run_formatting_safe(run, hyperlink_run['formatting'])
                    
                    # Apply hyperlink safely
                    try:
                        run.hyperlink.address = hyperlink_url
                        logger.debug(f"Applied hyperlink safely: '{hyperlink_text}' -> {hyperlink_url}")
                    except Exception as e:
                        logger.debug(f"Could not apply hyperlink: {e}")
                    
                    # Update remaining text
                    remaining_text = parts[1] if len(parts) > 1 else ""
                    break
            
            # Add remaining text
            if remaining_text:
                run = paragraph.add_run()
                run.text = remaining_text
                default_formatting = next((r['formatting'] for r in para_info['hyperlinks'] if not r['hyperlink']), 
                                        para_info['hyperlinks'][0]['formatting'] if para_info['hyperlinks'] else {})
                TextFrameUpdater._apply_run_formatting_safe(run, default_formatting)
                
        except Exception as e:
            logger.error(f"Error applying hyperlinks safely: {e}")
            # Fallback
            if not paragraph.runs:
                run = paragraph.add_run()
                run.text = line
    
    @staticmethod
    def _find_hyperlink_text_safe(translated_text: str, original_text: str):
        """Find text that should be hyperlinked safely"""
        # First try exact match
        if original_text in translated_text:
            return original_text
        
        # Common hyperlink patterns
        patterns = [
            'Boto3', 'Code samples', 'Starter Toolkit', 'samples', 'toolkit',
            '코드 샘플', '샘플', '툴킷', '스타터', 'Boto3', '코드'
        ]
        
        words = translated_text.split()
        for pattern in patterns:
            for word in words:
                if pattern.lower() in word.lower() or word.lower() in pattern.lower():
                    return word
        
        # Return first meaningful word
        meaningful_words = [word for word in words if len(word) > 2]
        return meaningful_words[0] if meaningful_words else None
    
    @staticmethod
    def _update_single_paragraph_safe(paragraph, new_text: str):
        """Update single paragraph preserving formatting safely"""
        try:
            # Extract formatting from first run
            original_formatting = None
            special_runs = []
            
            if paragraph.runs:
                original_formatting = TextFrameUpdater._extract_run_formatting_safe(paragraph.runs[0])
                
                # Look for special formatting (italic, colors)
                for run in paragraph.runs:
                    if run.text.strip():
                        formatting = TextFrameUpdater._extract_run_formatting_safe(run)
                        if (formatting.get('font_italic') or 
                            (formatting.get('font_color') and 
                             formatting['font_color'] not in [('rgb', '000000'), ('rgb', 'FFFFFF')])):
                            special_runs.append({
                                'text': run.text.strip(),
                                'formatting': formatting
                            })
            
            # Clear and rebuild
            paragraph.clear()
            
            # Apply special formatting if found
            if special_runs:
                TextFrameUpdater._apply_special_formatting_safe(paragraph, new_text, special_runs, original_formatting)
            else:
                # Simple formatting
                run = paragraph.add_run()
                run.text = new_text
                if original_formatting:
                    TextFrameUpdater._apply_run_formatting_safe(run, original_formatting)
                
        except Exception as e:
            logger.error(f"Single paragraph update failed: {e}")
            paragraph.clear()
            paragraph.add_run().text = new_text
    
    @staticmethod
    def _apply_special_formatting_safe(paragraph, new_text: str, special_runs: List[Dict], default_formatting):
        """Apply special formatting safely"""
        try:
            # Look for technical terms that should keep special formatting
            remaining_text = new_text
            
            for special_run in special_runs:
                original_term = special_run['text']
                
                # Check if original term exists in translation
                if original_term in remaining_text:
                    parts = remaining_text.split(original_term, 1)
                    
                    # Text before special term
                    if parts[0]:
                        run = paragraph.add_run()
                        run.text = parts[0]
                        if default_formatting:
                            TextFrameUpdater._apply_run_formatting_safe(run, default_formatting)
                    
                    # Special term
                    run = paragraph.add_run()
                    run.text = original_term
                    TextFrameUpdater._apply_run_formatting_safe(run, special_run['formatting'])
                    
                    # Update remaining text
                    remaining_text = parts[1] if len(parts) > 1 else ""
                    break
                
                # Check for technical patterns
                elif TextFrameUpdater._is_technical_term_safe(original_term):
                    # Find similar technical terms in translated text
                    found_terms = TextFrameUpdater._find_technical_terms_safe(remaining_text)
                    if found_terms:
                        tech_term = found_terms[0]
                        if tech_term in remaining_text:
                            parts = remaining_text.split(tech_term, 1)
                            
                            # Text before special term
                            if parts[0]:
                                run = paragraph.add_run()
                                run.text = parts[0]
                                if default_formatting:
                                    TextFrameUpdater._apply_run_formatting_safe(run, default_formatting)
                            
                            # Special term
                            run = paragraph.add_run()
                            run.text = tech_term
                            TextFrameUpdater._apply_run_formatting_safe(run, special_run['formatting'])
                            
                            # Update remaining text
                            remaining_text = parts[1] if len(parts) > 1 else ""
                            break
            
            # Add remaining text
            if remaining_text:
                run = paragraph.add_run()
                run.text = remaining_text
                if default_formatting:
                    TextFrameUpdater._apply_run_formatting_safe(run, default_formatting)
            
            # If no runs were added, add the whole text
            if not paragraph.runs:
                run = paragraph.add_run()
                run.text = new_text
                if default_formatting:
                    TextFrameUpdater._apply_run_formatting_safe(run, default_formatting)
                    
        except Exception as e:
            logger.error(f"Special formatting application failed: {e}")
            # Fallback
            if not paragraph.runs:
                run = paragraph.add_run()
                run.text = new_text
                if default_formatting:
                    TextFrameUpdater._apply_run_formatting_safe(run, default_formatting)
    
    @staticmethod
    def _is_technical_term_safe(term: str) -> bool:
        """Check if term is technical and should preserve formatting"""
        patterns = [r'.*Request$', r'.*Response$', r'.*ID$', r'^x-.*', r'^[A-Z]{2,}$']
        return any(re.match(pattern, term) for pattern in patterns)
    
    @staticmethod
    def _find_technical_terms_safe(text: str) -> List[str]:
        """Find technical terms in text"""
        words = text.split()
        patterns = [
            r'.*요청$', r'.*응답$', r'.*세션.*', r'.*Request$', r'.*Response$', r'.*ID$'
        ]
        
        return [word for word in words 
                for pattern in patterns 
                if re.match(pattern, word)]
    
    @staticmethod
    def _update_with_preserved_formatting_safe(text_frame, new_text: str):
        """Update text frame preserving first paragraph's formatting safely"""
        try:
            # Extract formatting from first paragraph/run
            original_formatting = None
            
            if text_frame.paragraphs and text_frame.paragraphs[0].runs:
                original_formatting = TextFrameUpdater._extract_run_formatting_safe(text_frame.paragraphs[0].runs[0])
            
            # Clear and rebuild
            text_frame.clear()
            new_lines = new_text.strip().split('\n')
            
            for i, line in enumerate(new_lines):
                if i > 0:
                    paragraph = text_frame.add_paragraph()
                else:
                    paragraph = text_frame.paragraphs[0]
                
                # Add text with formatting
                run = paragraph.add_run()
                run.text = line.strip()
                if original_formatting:
                    TextFrameUpdater._apply_run_formatting_safe(run, original_formatting)
                    
        except Exception as e:
            logger.error(f"Preserved formatting update failed: {e}")
            text_frame.text = new_text
    
    @staticmethod
    def _extract_run_formatting_safe(run) -> Dict[str, Any]:
        """Extract formatting from a run safely"""
        formatting = {
            'font_name': None,
            'font_size': None,
            'font_bold': None,
            'font_italic': None,
            'font_color': None
        }
        
        try:
            if hasattr(run, 'font') and run.font:
                font = run.font
                
                if hasattr(font, 'name') and font.name:
                    formatting['font_name'] = font.name
                if hasattr(font, 'size') and font.size:
                    formatting['font_size'] = font.size
                if hasattr(font, 'bold') and font.bold is not None:
                    formatting['font_bold'] = font.bold
                if hasattr(font, 'italic') and font.italic is not None:
                    formatting['font_italic'] = font.italic
                
                # Extract color safely
                if hasattr(font, 'color') and font.color:
                    try:
                        color_obj = font.color
                        if hasattr(color_obj, 'type'):
                            if color_obj.type == 1 and hasattr(color_obj, 'rgb'):  # RGB
                                formatting['font_color'] = ('rgb', str(color_obj.rgb))
                            elif color_obj.type == 2 and hasattr(color_obj, 'theme_color'):  # Theme
                                formatting['font_color'] = ('theme', color_obj.theme_color)
                    except Exception:
                        pass
                        
        except Exception as e:
            logger.debug(f"Could not extract run formatting safely: {e}")
        
        return formatting
    
    @staticmethod
    def _apply_run_formatting_safe(run, formatting: Dict[str, Any]):
        """Apply formatting to a run safely"""
        try:
            if hasattr(run, 'font') and run.font:
                font = run.font
                
                if formatting.get('font_name'):
                    font.name = formatting['font_name']
                if formatting.get('font_size'):
                    font.size = formatting['font_size']
                if formatting.get('font_bold') is not None:
                    font.bold = formatting['font_bold']
                if formatting.get('font_italic') is not None:
                    font.italic = formatting['font_italic']
                
                # Apply color safely
                if formatting.get('font_color'):
                    try:
                        color_info = formatting['font_color']
                        if isinstance(color_info, tuple) and len(color_info) == 2:
                            color_type, color_value = color_info
                            if color_type == 'rgb' and color_value:
                                # Convert hex string to RGBColor
                                if isinstance(color_value, str) and len(color_value) == 6:
                                    rgb_int = int(color_value, 16)
                                    font.color.rgb = RGBColor(
                                        (rgb_int >> 16) & 0xFF,
                                        (rgb_int >> 8) & 0xFF,
                                        rgb_int & 0xFF
                                    )
                            elif color_type == 'theme':
                                font.color.theme_color = color_value
                    except Exception:
                        pass
                        
        except Exception as e:
            logger.debug(f"Could not apply run formatting safely: {e}")


class PowerPointTranslator:
    """Main PowerPoint translation class"""
    
    def __init__(self, model_id: str = Config.DEFAULT_MODEL_ID, enable_polishing: bool = Config.ENABLE_POLISHING):
        self.model_id = model_id
        self.enable_polishing = enable_polishing
        self.engine = TranslationEngine(model_id, enable_polishing)
        self.text_collector = SlideTextCollector()
        self.text_updater = TextFrameUpdater()
        self.deps = DependencyManager()
    
    def translate_presentation(self, input_file: str, output_file: str, target_language: str) -> TranslationResult:
        """Translate entire PowerPoint presentation"""
        try:
            Presentation = self.deps.require('pptx')
            prs = Presentation(input_file)
            result = TranslationResult()
            
            total_slides = len(prs.slides)
            logger.info(f"🎯 Starting translation of {total_slides} slides...")
            logger.info(f"🎨 Translation mode: {'Natural/Polished' if self.enable_polishing else 'Literal'}")
            
            for slide_idx, slide in enumerate(prs.slides):
                logger.info(f"📄 Processing slide {slide_idx + 1}/{total_slides}")
                
                translated_count, notes_translated = self._translate_slide(slide, target_language)
                
                result.translated_count += translated_count
                if notes_translated:
                    result.translated_notes_count += 1
                result.total_shapes += len(slide.shapes)
                
                logger.info(f"✅ Slide {slide_idx + 1}: {translated_count} texts translated")
            
            # Save translated presentation
            prs.save(output_file)
            logger.info(f"🎉 Translation completed: {output_file}")
            logger.info(f"📊 Summary: {result.translated_count} texts, {result.translated_notes_count} notes")
            
            return result
            
        except Exception as e:
            logger.error(f"❌ Translation failed: {str(e)}")
            raise

    def translate_specific_slides(self, input_file: str, output_file: str, target_language: str, slide_numbers: List[int]) -> TranslationResult:
        """Translate specific slides in PowerPoint presentation
        
        Args:
            input_file: Path to input PowerPoint file
            output_file: Path to output PowerPoint file
            target_language: Target language code
            slide_numbers: List of slide numbers to translate (1-based indexing)
            
        Returns:
            TranslationResult with translation statistics
        """
        try:
            Presentation = self.deps.require('pptx')
            prs = Presentation(input_file)
            result = TranslationResult()
            
            total_slides = len(prs.slides)
            
            # Validate slide numbers
            invalid_slides = [num for num in slide_numbers if num < 1 or num > total_slides]
            if invalid_slides:
                error_msg = f"Invalid slide numbers: {invalid_slides}. Valid range: 1-{total_slides}"
                logger.error(error_msg)
                result.errors.append(error_msg)
                return result
            
            # Remove duplicates and sort
            slide_numbers = sorted(list(set(slide_numbers)))
            
            logger.info(f"🎯 Starting translation of {len(slide_numbers)} specific slides: {slide_numbers}")
            logger.info(f"🎨 Translation mode: {'Natural/Polished' if self.enable_polishing else 'Literal'}")
            
            for slide_num in slide_numbers:
                slide_idx = slide_num - 1  # Convert to 0-based index
                slide = prs.slides[slide_idx]
                
                logger.info(f"📄 Processing slide {slide_num}/{total_slides}")
                
                translated_count, notes_translated = self._translate_slide(slide, target_language)
                
                result.translated_count += translated_count
                if notes_translated:
                    result.translated_notes_count += 1
                result.total_shapes += len(slide.shapes)
                
                logger.info(f"✅ Slide {slide_num}: {translated_count} texts translated")
            
            # Save translated presentation
            prs.save(output_file)
            logger.info(f"🎉 Translation completed: {output_file}")
            logger.info(f"📊 Summary: {result.translated_count} texts, {result.translated_notes_count} notes from {len(slide_numbers)} slides")
            
            return result
            
        except Exception as e:
            logger.error(f"❌ Translation failed: {str(e)}")
            raise

    def get_slide_count(self, input_file: str) -> int:
        """Get total number of slides in PowerPoint presentation
        
        Args:
            input_file: Path to PowerPoint file
            
        Returns:
            Number of slides in the presentation
        """
        try:
            Presentation = self.deps.require('pptx')
            prs = Presentation(input_file)
            return len(prs.slides)
        except Exception as e:
            logger.error(f"❌ Failed to get slide count: {str(e)}")
            raise

    def get_slide_preview(self, input_file: str, slide_number: int, max_chars: int = 200) -> str:
        """Get a preview of text content from a specific slide
        
        Args:
            input_file: Path to PowerPoint file
            slide_number: Slide number (1-based indexing)
            max_chars: Maximum characters to return in preview
            
        Returns:
            Preview text from the slide
        """
        try:
            Presentation = self.deps.require('pptx')
            prs = Presentation(input_file)
            
            if slide_number < 1 or slide_number > len(prs.slides):
                raise ValueError(f"Invalid slide number: {slide_number}. Valid range: 1-{len(prs.slides)}")
            
            slide = prs.slides[slide_number - 1]  # Convert to 0-based index
            text_items, notes_text = self.text_collector.collect_slide_texts(slide)
            
            # Collect all text content
            all_texts = []
            for item in text_items:
                if item['text'].strip():
                    all_texts.append(item['text'].strip())
            
            if notes_text and notes_text.strip():
                all_texts.append(f"[Notes: {notes_text.strip()}]")
            
            # Join and truncate if necessary
            preview = " | ".join(all_texts)
            if len(preview) > max_chars:
                preview = preview[:max_chars] + "..."
            
            return preview if preview else "[No text content found]"
            
        except Exception as e:
            logger.error(f"❌ Failed to get slide preview: {str(e)}")
            raise
    
    def _translate_slide(self, slide, target_language: str) -> Tuple[int, bool]:
        """Translate a single slide"""
        text_items, notes_text = self.text_collector.collect_slide_texts(slide)
        
        translated_count = 0
        notes_translated = False
        
        # Translate notes if present
        if notes_text:
            notes_translated = self._translate_notes(slide, notes_text, target_language)
        
        # Check if slide has complex formatting
        if self._slide_has_complex_formatting(text_items):
            logger.info("🎨 Complex formatting detected, using individual translation")
            translated_count = self._translate_individually(text_items, target_language)
        # Choose translation strategy based on text count
        elif len(text_items) > Config.CONTEXT_THRESHOLD:
            translated_count = self._translate_with_context(text_items, target_language)
        else:
            translated_count = self._translate_with_batch(text_items, target_language)
        
        return translated_count, notes_translated
    
    def _slide_has_complex_formatting(self, text_items: List[Dict]) -> bool:
        """Check if slide has complex formatting"""
        for item in text_items:
            if item['type'] == 'text_frame_unified':
                text_frame = item['text_frame']
                for paragraph in text_frame.paragraphs:
                    # Check for indentation (lists)
                    if paragraph.level and paragraph.level > 0:
                        return True
                    
                    # Check for multiple runs with different formatting
                    if len(paragraph.runs) > 1:
                        colors = []
                        italic_states = []
                        
                        for run in paragraph.runs:
                            try:
                                # Check colors
                                if hasattr(run.font, 'color') and run.font.color:
                                    color = run.font.color
                                    if hasattr(color, 'type') and color.type == 1:  # RGB
                                        colors.append(str(color.rgb))
                                    elif hasattr(color, 'type') and color.type == 2:  # Theme
                                        colors.append(f"theme_{color.theme_color}")
                                
                                # Check italic
                                italic_states.append(run.font.italic if hasattr(run.font, 'italic') else None)
                                
                            except Exception:
                                pass
                        
                        # If we have different colors or italic states
                        if len(set(colors)) > 1 or len(set(italic_states)) > 1:
                            return True
        
        return False
    
    def _translate_notes(self, slide, notes_text: str, target_language: str) -> bool:
        """Translate slide notes"""
        try:
            translated_notes = self.engine.translate_text(notes_text, target_language)
            if translated_notes != notes_text:
                slide.notes_slide.notes_text_frame.text = translated_notes
                return True
        except Exception as e:
            logger.error(f"Error translating slide notes: {str(e)}")
        return False
    
    def _translate_individually(self, text_items: List[Dict], target_language: str) -> int:
        """Translate each text individually to preserve formatting"""
        translated_count = 0
        
        for item in text_items:
            try:
                original_text = item['text']
                translation = self.engine.translate_text(original_text, target_language)
                
                if original_text != translation:
                    if self._apply_translations([item], [translation]) > 0:
                        translated_count += 1
                        
            except Exception as e:
                logger.error(f"Individual translation failed: {str(e)}")
        
        return translated_count
    
    def _translate_with_context(self, text_items: List[Dict], target_language: str) -> int:
        """Translate using context-aware approach"""
        if not text_items:
            return 0
        
        try:
            translations = self.engine.translate_with_context(text_items, target_language)
            return self._apply_translations(text_items, translations)
        except Exception as e:
            logger.error(f"Context translation failed: {str(e)}")
            return self._translate_with_batch(text_items, target_language)
    
    def _translate_with_batch(self, text_items: List[Dict], target_language: str) -> int:
        """Translate using batch approach"""
        if not text_items:
            return 0
        
        texts_to_translate = [item['text'] for item in text_items]
        translated_count = 0
        
        # Process in batches
        for i in range(0, len(texts_to_translate), Config.BATCH_SIZE):
            batch_items = text_items[i:i + Config.BATCH_SIZE]
            batch_texts = texts_to_translate[i:i + Config.BATCH_SIZE]
            
            try:
                batch_translations = self.engine.translate_batch(batch_texts, target_language)
                translated_count += self._apply_translations(batch_items, batch_translations)
            except Exception as e:
                logger.error(f"Batch translation failed: {str(e)}")
                # Individual fallback
                for item in batch_items:
                    try:
                        translation = self.engine.translate_text(item['text'], target_language)
                        if self._apply_translations([item], [translation]) > 0:
                            translated_count += 1
                    except Exception:
                        pass
        
        return translated_count
    
    def _apply_translations(self, text_items: List[Dict], translations: List[str]) -> int:
        """Apply translations back to the original shapes"""
        if len(text_items) != len(translations):
            logger.error(f"Translation count mismatch: {len(text_items)} items, {len(translations)} translations")
            return 0
        
        translated_count = 0
        
        for item, translation in zip(text_items, translations):
            try:
                if item['text'] == translation:
                    continue
                
                item_type = item['type']
                
                if item_type == 'table_cell':
                    cell = item['cell']
                    if hasattr(cell, 'text_frame') and cell.text_frame:
                        self.text_updater.update_text_frame(cell.text_frame, translation)
                    else:
                        cell.text = translation
                    translated_count += 1
                    
                elif item_type == 'text_frame_unified':
                    text_frame = item['text_frame']
                    self.text_updater.update_text_frame(text_frame, translation)
                    translated_count += 1
                    
                elif item_type == 'direct_text':
                    item['shape'].text = translation
                    translated_count += 1
                
            except Exception as e:
                logger.error(f"Error applying translation: {str(e)}")
        
        return translated_count
