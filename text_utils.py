"""
Text processing utilities for translation
"""
import re
import logging
from typing import List, Dict, Any, Tuple
from config import Config

logger = logging.getLogger(__name__)


class TextProcessor:
    """Handles text processing and validation logic"""
    
    @staticmethod
    def should_skip_translation(text: str) -> bool:
        """Determine if text should be skipped from translation"""
        if not text or not text.strip():
            return True
        
        text = text.strip()
        
        # Check against skip patterns
        for pattern in Config.SKIP_PATTERNS:
            if re.match(pattern, text):
                return True
        
        # Skip very short text that's likely not translatable
        if len(text) <= 2 and not any(c.isalpha() for c in text):
            return True
        
        return False
    
    @staticmethod
    def clean_translation_response(response: str) -> str:
        """Clean up translation response by removing unwanted prefixes/suffixes"""
        unwanted_prefixes = [
            "Here are the translations:",
            "Here is the translation:",
            "Translated texts:",
            "Translated text:",
            "Translations:",
            "Translation:",
            "번역:",
            "번역 결과:",
            "다음은 번역입니다:",
            "翻译:",
            "翻訳:",
            "Traductions:",
            "Übersetzungen:",
            "Traducciones:",
            "The translations are:",
            "Translation results:"
        ]
        
        unwanted_suffixes = [
            "End of translations.",
            "Translation complete.",
            "번역 완료.",
            "翻译完成。",
            "翻訳完了。"
        ]
        
        cleaned = response.strip()
        
        # Remove prefixes
        for prefix in unwanted_prefixes:
            if cleaned.lower().startswith(prefix.lower()):
                cleaned = cleaned[len(prefix):].strip()
                logger.debug(f"🧹 Removed prefix: '{prefix}'")
                break
        
        # Remove suffixes
        for suffix in unwanted_suffixes:
            if cleaned.lower().endswith(suffix.lower()):
                cleaned = cleaned[:-len(suffix)].strip()
                logger.debug(f"🧹 Removed suffix: '{suffix}'")
                break
        
        return cleaned
    
    @staticmethod
    def clean_translation_part(part: str) -> str:
        """Clean individual translation part"""
        cleaned = part.strip()
        
        # Remove quotes if wrapped
        if (cleaned.startswith('"') and cleaned.endswith('"')) or \
           (cleaned.startswith("'") and cleaned.endswith("'")):
            cleaned = cleaned[1:-1].strip()
        
        # Remove numbered prefixes like "1. ", "2. ", etc.        
        cleaned = re.sub(r'^\d+\.\s+', '', cleaned)
                
        # Remove bullet points
        cleaned = re.sub(r'^[•\-\*]\s*', '', cleaned)
        
        return cleaned.strip()
    
    @staticmethod
    def parse_batch_response(response: str, expected_count: int) -> List[str]:
        """Parse batch translation response"""
        cleaned_response = TextProcessor.clean_translation_response(response)
        parts = cleaned_response.split("---SEPARATOR---")
        
        # Clean each part
        cleaned_parts = [TextProcessor.clean_translation_part(part) for part in parts]
        
        if len(cleaned_parts) != expected_count:
            logger.warning(f"⚠️ Batch translation count mismatch. Expected {expected_count}, got {len(cleaned_parts)}")
        
        return cleaned_parts
    
    @staticmethod
    def parse_context_response(response: str) -> List[str]:
        """Parse context-aware translation response"""
        logger.debug(f"🔍 Parsing translation response: {response[:200]}...")
        translations = []
        lines = response.strip().split('\n')
        
        current_translation = ""
        current_number = None
        
        for line in lines:
            line = line.strip()
            if line.startswith('[') and ']' in line:
                # Save previous translation
                if current_translation and current_number is not None:
                    translations.append(current_translation.strip())
                    logger.debug(f"🔍 Parsed translation {current_number}: '{current_translation[:50]}{'...' if len(current_translation) > 50 else ''}'")
                
                # Start new translation
                bracket_end = line.find(']')
                if bracket_end != -1:
                    current_number = line[1:bracket_end]
                    current_translation = line[bracket_end + 1:].strip()
            else:
                # Continue current translation (multi-line)
                if current_translation:
                    current_translation += "\n" + line
        
        # Don't forget the last translation
        if current_translation and current_number is not None:
            translations.append(current_translation.strip())
            logger.debug(f"🔍 Parsed translation {current_number}: '{current_translation[:50]}{'...' if len(current_translation) > 50 else ''}'")
        
        logger.debug(f"🔍 Total parsed translations: {len(translations)}")
        return translations


class SlideTextCollector:
    """Collects texts from PowerPoint slides"""
    
    @staticmethod
    def collect_slide_texts(slide) -> Tuple[List[Dict], str]:
        """Collect all translatable texts from a slide"""
        text_items = []
        notes_text = ""
        
        # Collect notes text
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes_text = slide.notes_slide.notes_text_frame.text.strip()
        except Exception as e:
            logger.error(f"Error collecting notes: {str(e)}")
        
        # Collect shape texts
        for shape_idx, shape in enumerate(slide.shapes):
            SlideTextCollector._collect_shape_texts(shape, text_items, shape_idx)
        
        return text_items, notes_text
    
    @staticmethod
    def _collect_shape_texts(shape, text_items: List[Dict], shape_idx: int, parent_path: str = ""):
        """Recursively collect texts from shapes"""
        current_path = f"{parent_path}.{shape_idx}" if parent_path else str(shape_idx)
        
        try:
            # Handle GROUP shapes recursively
            if hasattr(shape, 'shapes'):
                for sub_idx, sub_shape in enumerate(shape.shapes):
                    SlideTextCollector._collect_shape_texts(sub_shape, text_items, sub_idx, current_path)
                return
            
            # Handle table shapes
            if hasattr(shape, 'table'):
                SlideTextCollector._collect_table_texts(shape, text_items, current_path)
                return
            
            # Handle text frames
            if hasattr(shape, 'text_frame') and shape.text_frame:
                full_text = shape.text_frame.text.strip()
                if full_text and not TextProcessor.should_skip_translation(full_text):
                    text_items.append({
                        'type': 'text_frame_unified',
                        'path': f"{current_path}.text_frame",
                        'text': full_text,
                        'shape': shape,
                        'text_frame': shape.text_frame
                    })
                return
            
            # Handle shapes with direct text property
            if hasattr(shape, "text"):
                original_text = shape.text.strip()
                if original_text and not TextProcessor.should_skip_translation(original_text):
                    text_items.append({
                        'type': 'direct_text',
                        'path': f"{current_path}.text",
                        'text': original_text,
                        'shape': shape
                    })
                        
        except Exception as e:
            logger.error(f"Error collecting shape texts: {str(e)}")
    
    @staticmethod
    def _collect_table_texts(shape, text_items: List[Dict], current_path: str):
        """Collect texts from table cells"""
        try:
            table = shape.table
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if cell_text and not TextProcessor.should_skip_translation(cell_text):
                        text_items.append({
                            'type': 'table_cell',
                            'path': f"{current_path}.table.{row_idx}.{cell_idx}",
                            'text': cell_text,
                            'shape': shape,
                            'cell': cell,
                            'row_idx': row_idx,
                            'cell_idx': cell_idx
                        })
        except Exception as e:
            logger.error(f"Error collecting table texts: {str(e)}")
    
    @staticmethod
    def build_slide_context(text_items: List[Dict], notes_text: str) -> str:
        """Build context information for the slide"""
        context_parts = ["SLIDE CONTENT:"]
        
        for i, item in enumerate(text_items):
            item_type = item['type']
            text = item['text']
            
            if item_type == 'table_cell':
                context_parts.append(f"[{i+1}] Table Cell: {text}")
            elif item_type == 'text_frame_unified':
                context_parts.append(f"[{i+1}] Text Frame: {text}")
            elif item_type == 'direct_text':
                context_parts.append(f"[{i+1}] Direct Text: {text}")
            else:
                context_parts.append(f"[{i+1}] {text}")
        
        if notes_text:
            context_parts.append(f"\nSLIDE NOTES: {notes_text}")
        
        return "\n".join(context_parts)
