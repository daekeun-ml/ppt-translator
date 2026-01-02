"""
Text processing utilities for translation
"""
import re
import logging
from typing import List, Dict, Any, Tuple
from .config import Config

logger = logging.getLogger(__name__)


class TextProcessor:
    """Handles text processing and validation logic"""
    
    @staticmethod
    def should_skip_translation(text: str) -> bool:
        """Determine if text should be skipped from translation"""
        if not text or not text.strip():
            return True
        
        text = text.strip()
        
        # Skip code blocks (enclosed in triple backticks)
        if text.startswith('```') or text.endswith('```'):
            return True
        
        # Skip JSON-like structures
        if (text.startswith('{') and text.endswith('}')) or (text.startswith('[') and text.endswith(']')):
            return True
        
        # Skip if contains JSON key-value patterns
        json_patterns = [
            r'"[^"]+"\s*:\s*"[^"]*"',  # "key": "value"
            r'"[^"]+"\s*:\s*\{',       # "key": {
            r'"[^"]+"\s*:\s*\[',       # "key": [
            r'\{\s*"[^"]+"\s*:',       # {"key":
        ]
        
        if any(re.search(pattern, text) for pattern in json_patterns):
            return True
        
        # Comprehensive code patterns for multiple languages
        code_patterns = [
            # Python
            r'\bdef\s+\w+\s*\(',
            r'\bclass\s+\w+\s*[\(:]',
            r'\bimport\s+\w+',
            r'\bfrom\s+\w+\s+import',
            r'\bprint\s*\(',
            r'\b__\w+__\b',
            r'\bself\.\w+',
            
            # JavaScript/TypeScript
            r'\bfunction\s+\w+\s*\(',
            r'\bvar\s+\w+\s*=',
            r'\blet\s+\w+\s*=',
            r'\bconst\s+\w+\s*=',
            r'\bconsole\.\w+\s*\(',
            r'=>\s*\{',
            r'\$\{\w+\}',
            
            # Java/C#/C++
            r'\bpublic\s+\w+',
            r'\bprivate\s+\w+',
            r'\bprotected\s+\w+',
            r'\bstatic\s+\w+',
            r'\bvoid\s+\w+\s*\(',
            r'\bint\s+\w+\s*[=;]',
            r'\bString\s+\w+\s*[=;]',
            r'System\.out\.print',
            
            # General programming patterns
            r'\bif\s*\([^)]+\)\s*\{',
            r'\bfor\s*\([^)]+\)\s*\{',
            r'\bwhile\s*\([^)]+\)\s*\{',
            r'\btry\s*\{',
            r'\bcatch\s*\([^)]+\)\s*\{',
            r'\breturn\s+[^;]+;',
            r'\w+\s*=\s*new\s+\w+\s*\(',
            
            # Common code symbols and structures
            r'\w+\.\w+\s*\(',  # method calls
            r'\w+\[\w*\]\s*=',  # array assignments
            r'//.*$',  # single line comments
            r'/\*.*?\*/',  # multi-line comments
            r'#.*$',  # Python/shell comments
        ]
        
        # Count matches
        code_matches = sum(1 for pattern in code_patterns if re.search(pattern, text, re.MULTILINE))
        
        # If multiple code patterns match, likely code
        if code_matches >= 2:
            return True
        
        # Skip if text has high ratio of special characters (likely code)
        special_chars = sum(1 for c in text if c in '{}[]()":,;=<>+-*/%&|!^~')
        total_chars = len(text)
        if total_chars > 10 and (special_chars / total_chars) > 0.25:
            return True
        
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
        cleaned = response.strip()
        
        # If response contains "I'd be happy to help" or similar, it's not a translation
        if any(phrase in cleaned for phrase in [
            "I'd be happy to help",
            "I don't see any text",
            "Could you please provide",
            "appears to be a question",
            "Once you share it"
        ]):
            # Return empty string to trigger fallback
            return ""
        
        # Remove markdown headers
        cleaned = re.sub(r'^#+\s*', '', cleaned, flags=re.MULTILINE)
        
        # Split by lines and filter out prompt-like content
        lines = cleaned.split('\n')
        translation_lines = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Skip lines that are clearly prompts or instructions
            skip_patterns = [
                'translate this exact text',
                'translate each text',
                'keep same order',
                'separate with',
                'format:',
                '---separator---'
            ]
            
            skip_line = any(pattern.lower() in line.lower() for pattern in skip_patterns)
            
            if not skip_line and line:
                translation_lines.append(line)
        
        # Join the clean lines
        if translation_lines:
            cleaned = '\n'.join(translation_lines)
        
        return cleaned.strip()
    
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
        
    @staticmethod
    def clean_translation_part(part: str) -> str:
        """Clean individual translation part with stricter rules"""
        cleaned = part.strip()
        
        # Remove quotes if wrapped
        if (cleaned.startswith('"') and cleaned.endswith('"')) or \
           (cleaned.startswith("'") and cleaned.endswith("'")):
            cleaned = cleaned[1:-1].strip()
        
        # Remove "Translation to [Language]:" prefixes
        cleaned = re.sub(r'^Translation to \w+:\s*', '', cleaned, flags=re.IGNORECASE)
        
        # Remove numbered prefixes like "1. ", "2. ", etc.        
        cleaned = re.sub(r'^\d+\.\s+', '', cleaned)
                
        # Remove bullet points
        cleaned = re.sub(r'^[•\-\*]\s*', '', cleaned)
        
        # Remove markdown formatting
        cleaned = re.sub(r'\*\*(.*?)\*\*', r'\1', cleaned)  # **text** -> text
        cleaned = re.sub(r'\*(.*?)\*', r'\1', cleaned)      # *text* -> text
        
        # Split by lines and extract only the main translation
        lines = cleaned.split('\n')
        main_translation = ""
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Skip explanation lines
            if any(skip_phrase in line.lower() for skip_phrase in [
                'alternative translations', 'depending on context', 'raw source', 
                'if referring to', 'the most common', 'translation is', '---',
                'unprocessed', 'original material', 'emphasizing'
            ]):
                continue
                
            # Handle lines with arrows (→) - extract the translation part
            if '→' in line:
                parts = line.split('→')
                if len(parts) > 1:
                    main_translation = parts[-1].strip()
                    break
                continue
            
            # Take the first clean line as the main translation
            if line and not main_translation:
                main_translation = line
                break
        
        # Final cleanup: remove any remaining language prefixes
        main_translation = re.sub(r'^(Korean|Japanese|English|Chinese|Spanish|French|German|Italian|Portuguese|Russian|Arabic|Hindi|한국어|일본어|영어|중국어):\s*', '', main_translation, flags=re.IGNORECASE)
        
        return main_translation.strip()
    
    @staticmethod
    def parse_batch_response(response: str, expected_count: int) -> List[str]:
        """Parse batch translation response with improved error handling"""
        cleaned_response = TextProcessor.clean_translation_response(response)
        parts = cleaned_response.split("---SEPARATOR---")
        
        # Clean each part
        cleaned_parts = [TextProcessor.clean_translation_part(part) for part in parts if part.strip()]
        
        # If count mismatch, try alternative parsing methods
        if len(cleaned_parts) != expected_count:
            logger.warning(f"⚠️ Batch translation count mismatch. Expected {expected_count}, got {len(cleaned_parts)}")
            
            # Try parsing with numbered format [1], [2], etc.
            numbered_parts = TextProcessor.parse_numbered_response(response, expected_count)
            if len(numbered_parts) == expected_count:
                logger.info("✅ Successfully parsed using numbered format")
                return numbered_parts
            
            # Try parsing with line breaks
            line_parts = TextProcessor._parse_line_response(response, expected_count)
            if len(line_parts) == expected_count:
                logger.info("✅ Successfully parsed using line format")
                return line_parts
            
            # If still mismatch, pad or truncate to match expected count
            if len(cleaned_parts) < expected_count:
                # Pad with empty strings
                cleaned_parts.extend([''] * (expected_count - len(cleaned_parts)))
                logger.warning(f"⚠️ Padded response to match expected count")
            elif len(cleaned_parts) > expected_count:
                # Truncate to expected count
                cleaned_parts = cleaned_parts[:expected_count]
                logger.warning(f"⚠️ Truncated response to match expected count")
        
        return cleaned_parts
    
    @staticmethod
    def parse_numbered_response(response: str, expected_count: int) -> List[str]:
        """Try to parse response with numbered format [1], [2], etc."""
        translations = []
        lines = response.strip().split('\n')
        current_translation = ""
        
        for line in lines:
            line = line.strip()
            if re.match(r'^\[\d+\]', line):
                # Save previous translation
                if current_translation:
                    translations.append(current_translation.strip())
                # Start new translation (remove the number part)
                current_translation = re.sub(r'^\[\d+\]\s*', '', line)
            else:
                # Continue current translation
                if current_translation:
                    current_translation += " " + line
        
        # Add the last translation
        if current_translation:
            translations.append(current_translation.strip())
        
        return translations
    
    @staticmethod
    def _parse_line_response(response: str, expected_count: int) -> List[str]:
        """Try to parse response by splitting on double line breaks"""
        parts = re.split(r'\n\s*\n', response.strip())
        cleaned_parts = []
        
        for part in parts:
            cleaned = TextProcessor.clean_translation_part(part)
            if cleaned:  # Only add non-empty parts
                cleaned_parts.append(cleaned)
        
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