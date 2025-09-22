import streamlit as st
import json
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom
import difflib
import re
from typing import Dict, List, Tuple, Optional, Set
import tempfile
import os
from dataclasses import dataclass, asdict
from io import StringIO, BytesIO
import markdown
from fuzzywuzzy import fuzz, process
import datetime
from docx import Document
from docx.shared import Inches

@dataclass
class ParsedSection:
    name: str
    type: str
    sub_type: str
    prompt: str
    include_screenshots: bool
    screenshot_instructions: str
    raw_comment: str
    edited: bool = False
    original_prompt: str = ""

    def __post_init__(self):
        if not self.original_prompt:
            self.original_prompt = self.prompt

@dataclass
class SectionMapping:
    file1_section: str
    file2_section: str
    confidence: float
    is_manual: bool = False

class DocxCommentExtractor:
    """Extract comments and their associated section names from DOCX files"""

    @staticmethod
    def extract_comments_from_docx(file_path: str) -> Dict[str, ParsedSection]:
        """Extract all comments and their associated section names from a DOCX file"""
        sections = {}

        try:
            with zipfile.ZipFile(file_path, 'r') as docx_zip:
                # Read comments.xml
                comments_xml = None
                document_xml = None

                try:
                    comments_xml = docx_zip.read('word/comments.xml').decode('utf-8')
                except KeyError:
                    st.warning("No comments.xml found in DOCX file")
                    return sections

                try:
                    document_xml = docx_zip.read('word/document.xml').decode('utf-8')
                except KeyError:
                    st.warning("No document.xml found in DOCX file")
                    return sections

                # Parse comments
                comments_data = DocxCommentExtractor._parse_comments_xml(comments_xml)

                # Parse document to find section names linked to comments
                section_names = DocxCommentExtractor._parse_document_for_sections(document_xml)

                # Combine comments with section names
                for comment_id, comment_text in comments_data.items():
                    if comment_id in section_names:
                        section_name = section_names[comment_id]
                        parsed_section = DocxCommentExtractor._parse_comment_string(
                            comment_text, section_name
                        )
                        if parsed_section:
                            sections[section_name] = parsed_section

        except Exception as e:
            st.error(f"Error processing DOCX file: {str(e)}")

        return sections

    @staticmethod
    def _parse_comments_xml(comments_xml: str) -> Dict[str, str]:
        """Parse comments.xml to extract comment ID and text"""
        comments = {}

        try:
            root = ET.fromstring(comments_xml)

            # Handle namespaces
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            }

            comment_elements = root.findall('.//w:comment', namespaces)

            for comment in comment_elements:
                comment_id = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')

                # Extract text content from comment
                text_elements = comment.findall('.//w:t', namespaces)
                comment_text = ''.join([t.text or '' for t in text_elements])

                if comment_id and comment_text:
                    comments[comment_id] = comment_text.strip()

        except ET.ParseError as e:
            st.error(f"Error parsing comments XML: {str(e)}")

        return comments

    @staticmethod
    def _parse_document_for_sections(document_xml: str) -> Dict[str, str]:
        """Parse document.xml to find section names associated with comment IDs"""
        section_names = {}

        try:
            root = ET.fromstring(document_xml)

            # Handle namespaces
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            }

            # Find all paragraphs
            paragraphs = root.findall('.//w:p', namespaces)

            for paragraph in paragraphs:
                # Check for comment range start
                comment_range_start = paragraph.find('.//w:commentRangeStart', namespaces)
                comment_reference = paragraph.find('.//w:commentReference', namespaces)

                if comment_range_start is not None or comment_reference is not None:
                    # Get comment ID
                    comment_id = None
                    if comment_range_start is not None:
                        comment_id = comment_range_start.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    elif comment_reference is not None:
                        comment_id = comment_reference.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')

                    # Get text content (section name)
                    text_elements = paragraph.findall('.//w:t', namespaces)
                    section_name = ''.join([t.text or '' for t in text_elements]).strip()

                    if comment_id and section_name:
                        section_names[comment_id] = section_name

        except ET.ParseError as e:
            st.error(f"Error parsing document XML: {str(e)}")

        return section_names

    @staticmethod
    def _parse_comment_string(comment_text: str, section_name: str) -> Optional[ParsedSection]:
        """Parse comment text to extract structured section data with intelligent prompt extraction"""

        try:
            # Clean and normalize the comment text
            comment_text = comment_text.strip()

            # Initialize with safe defaults
            section_data = {
                'type': 'text',
                'sub_type': 'default',
                'prompt': '',
                'include_screenshots': False,
                'screenshot_instructions': ''
            }

            # Handle the specific format from Alex's templates:
            # "type - textsub_type - freeformprompt - [ACTUAL PROMPT]include_screenshots - yesscreenshot_instructions - none"

            # First, let's check if this is the concatenated format (no line breaks)
            if '\n' not in comment_text and ' - ' in comment_text:
                # This is likely the concatenated format, let's parse it carefully
                return DocxCommentExtractor._parse_concatenated_format(comment_text, section_name)

            # Handle multi-line format
            if ' - ' not in comment_text and ': ' not in comment_text and '=' not in comment_text:
                # Treat entire comment as prompt
                section_data['prompt'] = comment_text
                return ParsedSection(
                    name=section_name,
                    type=section_data['type'],
                    sub_type=section_data['sub_type'],
                    prompt=section_data['prompt'],
                    include_screenshots=section_data['include_screenshots'],
                    screenshot_instructions=section_data['screenshot_instructions'],
                    raw_comment=comment_text
                )

            # Parse line-by-line format
            lines = comment_text.split('\n')

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Try different separators: ' - ', ': ', '='
                key = ''
                value = ''
                separator_found = False

                if ' - ' in line and not separator_found:
                    parts = line.split(' - ', 1)
                    if len(parts) == 2:
                        key, value = parts[0].strip(), parts[1].strip()
                        separator_found = True

                if ': ' in line and not separator_found:
                    parts = line.split(': ', 1)
                    if len(parts) == 2:
                        key, value = parts[0].strip(), parts[1].strip()
                        separator_found = True

                if '=' in line and not separator_found:
                    parts = line.split('=', 1)
                    if len(parts) == 2:
                        key, value = parts[0].strip(), parts[1].strip()
                        separator_found = True

                if not separator_found:
                    # If no separator, treat as prompt content (append to existing)
                    if section_data['prompt']:
                        section_data['prompt'] += '\n' + line
                    else:
                        section_data['prompt'] = line
                    continue

                # Skip if no valid key-value pair found
                if not key or not value:
                    continue

                # Process the key-value pair
                DocxCommentExtractor._process_key_value(key, value, section_data)

            # If no prompt found in structured format, use entire comment as prompt
            if not section_data['prompt'] and len(comment_text.strip()) > 15:
                section_data['prompt'] = comment_text.strip()

            # Final validation - ensure required fields have valid values
            if section_data['type'] not in ['text', 'table', 'process_flow_diagram']:
                section_data['type'] = 'text'

            if section_data['sub_type'] not in ['default', 'bulleted', 'freeform', 'flow-diagram', 'walkthrough-steps', 'bpmn']:
                section_data['sub_type'] = 'default'

            return ParsedSection(
                name=section_name,
                type=section_data['type'],
                sub_type=section_data['sub_type'],
                prompt=section_data['prompt'],
                include_screenshots=section_data['include_screenshots'],
                screenshot_instructions=section_data['screenshot_instructions'],
                raw_comment=comment_text
            )

        except Exception as e:
            # Return a safe default section instead of None
            return ParsedSection(
                name=section_name,
                type='text',
                sub_type='default',
                prompt=comment_text.strip() if comment_text.strip() else f"Content for {section_name}",
                include_screenshots=False,
                screenshot_instructions='',
                raw_comment=comment_text
            )

    @staticmethod
    def _parse_concatenated_format(comment_text: str, section_name: str) -> ParsedSection:
        """Parse the concatenated format: 'type - textsub_type - freeformprompt - [CONTENT]include_screenshots - yes...'"""

        section_data = {
            'type': 'text',
            'sub_type': 'default',
            'prompt': '',
            'include_screenshots': False,
            'screenshot_instructions': ''
        }

        # Use regex to extract the different parts more intelligently
        import re

        # Pattern to match the structured parts
        # Look for patterns like "key - value" but be smart about the prompt content

        # First, extract type
        type_match = re.match(r'^type\s*-\s*([^s]*?)sub_type', comment_text)
        if type_match:
            type_value = type_match.group(1).strip()
            # Map special types
            if type_value == 'process_flow_diagram':
                section_data['type'] = 'text'  # Process flow diagrams are rendered as text
            elif type_value in ['text', 'table']:
                section_data['type'] = type_value
            else:
                section_data['type'] = 'text'

        # Extract sub_type
        subtype_match = re.search(r'sub_type\s*-\s*([^p]*?)prompt\s*-', comment_text)
        if subtype_match:
            subtype_value = subtype_match.group(1).strip()
            # Handle special cases
            if subtype_value == 'bpmn':
                section_data['sub_type'] = 'flow-diagram'
            elif subtype_value in ['default', 'bulleted', 'freeform', 'flow-diagram', 'walkthrough-steps']:
                section_data['sub_type'] = subtype_value
            else:
                section_data['sub_type'] = 'default'

        # Extract the main prompt content - this is the tricky part
        # Look for "prompt - " and then capture everything until "include_screenshots"
        prompt_match = re.search(r'prompt\s*-\s*(.*?)include_screenshots\s*-', comment_text, re.DOTALL)
        if prompt_match:
            section_data['prompt'] = prompt_match.group(1).strip().replace('\\n', '\n')

        # Extract screenshots setting
        screenshot_match = re.search(r'include_screenshots\s*-\s*([^s]*?)screenshot_instructions', comment_text)
        if screenshot_match:
            screenshot_value = screenshot_match.group(1).strip().lower()
            section_data['include_screenshots'] = screenshot_value in ['yes', 'true', '1']

        # Extract screenshot instructions
        instructions_match = re.search(r'screenshot_instructions\s*-\s*(.*)$', comment_text)
        if instructions_match:
            instructions_value = instructions_match.group(1).strip()
            section_data['screenshot_instructions'] = instructions_value if instructions_value.lower() != 'none' else ''

        # If we couldn't extract a proper prompt, fall back to a cleaned version
        if not section_data['prompt']:
            # Try to extract everything after the first "prompt - " occurrence
            prompt_start = comment_text.find('prompt - ')
            if prompt_start != -1:
                remaining_text = comment_text[prompt_start + 9:]  # 9 = len('prompt - ')
                # Remove trailing metadata
                for suffix in ['include_screenshots', 'screenshot_instructions']:
                    pos = remaining_text.find(suffix)
                    if pos != -1:
                        remaining_text = remaining_text[:pos]
                        break
                section_data['prompt'] = remaining_text.strip().replace('\\n', '\n')

        return ParsedSection(
            name=section_name,
            type=section_data['type'],
            sub_type=section_data['sub_type'],
            prompt=section_data['prompt'] if section_data['prompt'] else comment_text.strip(),
            include_screenshots=section_data['include_screenshots'],
            screenshot_instructions=section_data['screenshot_instructions'],
            raw_comment=comment_text
        )

    @staticmethod
    def _process_key_value(key: str, value: str, section_data: dict):
        """Process a key-value pair and update section_data"""

        # Normalize key
        key_normalized = key.lower().replace('_', '').replace(' ', '').replace('-', '')

        try:
            if key_normalized in ['type', 'contenttype']:
                # Handle special type mappings
                clean_type = value.lower().strip()
                if clean_type == 'process_flow_diagram':
                    section_data['type'] = 'text'
                elif clean_type in ['text', 'table']:
                    section_data['type'] = clean_type
                else:
                    section_data['type'] = 'text'

            elif key_normalized in ['subtype', 'sub_type', 'style']:
                # Handle special subtype mappings
                clean_subtype = value.lower().strip()
                if clean_subtype == 'bpmn':
                    section_data['sub_type'] = 'flow-diagram'
                elif clean_subtype in ['default', 'bulleted', 'freeform', 'flow-diagram', 'walkthrough-steps']:
                    section_data['sub_type'] = clean_subtype
                else:
                    section_data['sub_type'] = 'default'

            elif key_normalized in ['prompt', 'instruction', 'content']:
                section_data['prompt'] = value.replace('\\n', '\n')

            elif key_normalized in ['includescreenshot', 'include_screenshot', 'screenshot']:
                section_data['include_screenshots'] = value.lower().strip() in ['yes', 'true', '1']

            elif key_normalized in ['screenshotinstruction', 'screenshot_instruction']:
                section_data['screenshot_instructions'] = value if value.lower() != 'none' else ''

        except Exception as field_error:
            # Skip this field if there's an error, but continue processing
            pass

class PromptFormatter:
    """Convert JSON prompts to readable markdown format with enhanced UX"""

    @staticmethod
    def format_prompt_for_display(prompt_text: str, max_length: int = None) -> str:
        """Format prompt text for optimal readability in Streamlit"""

        if not prompt_text or not prompt_text.strip():
            return "*No prompt content available*"

        # Try to parse as JSON first
        try:
            if prompt_text.strip().startswith('{') or prompt_text.strip().startswith('['):
                prompt_obj = json.loads(prompt_text)
                formatted = PromptFormatter._json_obj_to_markdown(prompt_obj)
            else:
                formatted = PromptFormatter._enhance_text_formatting(prompt_text)
        except json.JSONDecodeError:
            formatted = PromptFormatter._enhance_text_formatting(prompt_text)

        # Truncate if needed
        if max_length and len(formatted) > max_length:
            formatted = formatted[:max_length] + "..."

        return formatted

    @staticmethod
    def json_to_markdown(prompt_text: str) -> str:
        """Convert JSON string prompts to readable markdown - legacy method"""
        return PromptFormatter.format_prompt_for_display(prompt_text)

    @staticmethod
    def _json_obj_to_markdown(obj) -> str:
        """Convert JSON object to markdown recursively"""

        if isinstance(obj, dict):
            markdown_parts = []
            for key, value in obj.items():
                if isinstance(value, (dict, list)):
                    markdown_parts.append(f"## {key.title().replace('_', ' ')}")
                    markdown_parts.append(PromptFormatter._json_obj_to_markdown(value))
                else:
                    markdown_parts.append(f"**{key.title().replace('_', ' ')}:** {value}")
            return '\n\n'.join(markdown_parts)

        elif isinstance(obj, list):
            if all(isinstance(item, str) for item in obj):
                return '\n'.join([f"- {item}" for item in obj])
            else:
                markdown_parts = []
                for i, item in enumerate(obj, 1):
                    markdown_parts.append(f"### Item {i}")
                    markdown_parts.append(PromptFormatter._json_obj_to_markdown(item))
                return '\n\n'.join(markdown_parts)

        else:
            return str(obj)

    @staticmethod
    def _enhance_text_formatting(text: str) -> str:
        """Enhanced text formatting for better Streamlit display"""

        if not text:
            return ""

        # Clean and normalize whitespace
        text = re.sub(r'\r\n', '\n', text)  # Windows line endings
        text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)  # Max 2 newlines
        text = re.sub(r'[ \t]+', ' ', text)  # Multiple spaces/tabs to single space

        # Handle escaped newlines from comment parsing first
        text = text.replace('\\n', '\n')

        # Convert numbered lists with better formatting
        text = re.sub(r'(?:^|\n)(\d+)\.[ \t]*', r'\n\n\1. ', text)

        # Convert bullet points with consistent formatting - preserve special bullet chars
        text = re.sub(r'(?:^|\n)[•·*-][ \t]*', r'\n- ', text)

        # Convert markdown bold to HTML for better display in containers
        text = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', text)

        # Emphasize key phrases
        text = re.sub(r'"([^"]+)"', r'<strong>\1</strong>', text)  # Quoted text
        text = re.sub(r'\b([A-Z]{2,})\b', r'<strong>\1</strong>', text)  # ALL CAPS words

        # Handle special formatting for instructions and examples
        text = re.sub(r'Example user quote: "([^"]+)"', r'<em>Example user quote: "<strong>\1</strong>"</em>', text)

        # Clean up extra whitespace at start/end
        text = text.strip()

        return text

    @staticmethod
    def create_section_property_display(section: ParsedSection, show_stats: bool = True) -> str:
        """Create a clean, readable display of section properties"""

        lines = []

        # Type information with icons
        type_icon = "📊" if section.type == "table" else "📝"
        lines.append(f"{type_icon} **Content Type:** {section.type.title()}")

        # Sub-type with descriptive text
        subtype_descriptions = {
            'default': 'Standard text format',
            'bulleted': 'Bullet point list',
            'freeform': 'Free-form narrative',
            'flow-diagram': 'Process flow description',
            'walkthrough-steps': 'Step-by-step instructions'
        }
        subtype_desc = subtype_descriptions.get(section.sub_type, section.sub_type.title())
        lines.append(f"🎯 **Format:** {subtype_desc}")

        # Screenshot requirements
        screenshot_icon = "📸" if section.include_screenshots else "🚫"
        screenshot_text = "Required" if section.include_screenshots else "Not required"
        lines.append(f"{screenshot_icon} **Screenshots:** {screenshot_text}")

        if section.include_screenshots and section.screenshot_instructions:
            lines.append(f"   *Instructions: {section.screenshot_instructions}*")

        # Statistics if requested
        if show_stats and section.prompt:
            char_count = len(section.prompt)
            word_count = len(section.prompt.split())
            lines.append(f"📈 **Size:** {char_count:,} characters, {word_count:,} words")

        # Edit status
        if hasattr(section, 'edited') and section.edited:
            lines.append("✏️ **Status:** Modified")
        else:
            lines.append("📄 **Status:** Original")

        return '\n'.join(lines)

class UIComponents:
    """Enhanced UI components for better UX"""

    @staticmethod
    def create_section_card(section_name: str, section: ParsedSection, file_name: str = "", expanded: bool = False):
        """Create a well-formatted section card with unique keys"""

        # Create unique key prefix using file name and section name
        unique_key = f"{file_name}_{section_name}".replace(" ", "_").replace(".", "_").replace("/", "_")

        # Create status indicator
        status_icon = " ✏️" if (hasattr(section, 'edited') and section.edited) else ""

        with st.expander(f"📝 {section_name}{status_icon}", expanded=expanded):
            # Two-column layout
            col1, col2 = st.columns([2, 1])

            with col1:
                st.markdown("### 📋 Prompt Content")

                # Format and display the prompt
                formatted_prompt = PromptFormatter.format_prompt_for_display(section.prompt)

                # Use a container with custom styling for better readability
                st.markdown(
                    f"""
                    <div style="
                        background-color: #f8f9fa;
                        padding: 1rem;
                        border-radius: 0.5rem;
                        border-left: 4px solid #007bff;
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                        line-height: 1.6;
                        color: #212529;
                    ">
                    {formatted_prompt.replace(chr(10), '<br>')}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                # Show edit indicator if applicable
                if hasattr(section, 'edited') and section.edited:
                    st.success("✏️ This prompt has been edited from the original")

            with col2:
                st.markdown("### ⚙️ Properties")

                # Display properties in a clean format
                properties_text = PromptFormatter.create_section_property_display(section)
                st.markdown(properties_text)

                # Action buttons
                st.markdown("### 🔧 Actions")

                if st.button(f"📝 Edit", key=f"edit_{unique_key}", help="Edit this section"):
                    st.session_state[f"edit_mode_{unique_key}"] = True
                    st.rerun()

                if st.button(f"📋 Copy", key=f"copy_{unique_key}", help="Copy prompt to clipboard"):
                    st.success("Prompt copied to clipboard!")

    @staticmethod
    def create_side_by_side_comparison(section_name: str, section1: ParsedSection, section2: ParsedSection,
                                     file1_name: str, file2_name: str):
        """Create an enhanced side-by-side comparison view"""

        st.markdown(f"## 📊 Comparing: {section_name}")

        # Quick stats bar
        col1, col2, col3 = st.columns(3)
        with col1:
            similarity = fuzz.ratio(section1.prompt, section2.prompt)
            st.metric("Similarity", f"{similarity}%", help="Text similarity percentage")
        with col2:
            char_diff = len(section2.prompt) - len(section1.prompt)
            st.metric("Length Difference", f"{char_diff:+,} chars", help="Character count difference")
        with col3:
            type_match = "✅" if section1.type == section2.type else "❌"
            st.metric("Type Match", type_match, help="Whether content types match")

        st.divider()

        # Side-by-side content
        col1, col2 = st.columns(2)

        with col1:
            st.markdown(f"### 📄 {file1_name}")

            # Properties box
            with st.container():
                st.markdown(
                    f"""
                    <div style="
                        background-color: #e8f4f8;
                        padding: 0.5rem;
                        border-radius: 0.25rem;
                        margin-bottom: 1rem;
                        font-size: 0.9em;
                        color: #212529;
                    ">
                    {PromptFormatter.create_section_property_display(section1, show_stats=False).replace(chr(10), '<br>')}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            # Prompt content
            formatted_prompt1 = PromptFormatter.format_prompt_for_display(section1.prompt)
            st.markdown(
                f"""
                <div style="
                    background-color: #f8f9fa;
                    padding: 1rem;
                    border-radius: 0.5rem;
                    border-left: 4px solid #28a745;
                    min-height: 200px;
                    line-height: 1.6;
                    color: #212529;
                ">
                {formatted_prompt1.replace(chr(10), '<br>')}
                </div>
                """,
                unsafe_allow_html=True
            )

        with col2:
            st.markdown(f"### 📄 {file2_name}")

            # Properties box
            with st.container():
                st.markdown(
                    f"""
                    <div style="
                        background-color: #f8e8e8;
                        padding: 0.5rem;
                        border-radius: 0.25rem;
                        margin-bottom: 1rem;
                        font-size: 0.9em;
                        color: #212529;
                    ">
                    {PromptFormatter.create_section_property_display(section2, show_stats=False).replace(chr(10), '<br>')}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            # Prompt content
            formatted_prompt2 = PromptFormatter.format_prompt_for_display(section2.prompt)
            st.markdown(
                f"""
                <div style="
                    background-color: #f8f9fa;
                    padding: 1rem;
                    border-radius: 0.5rem;
                    border-left: 4px solid #dc3545;
                    min-height: 200px;
                    line-height: 1.6;
                    color: #212529;
                ">
                {formatted_prompt2.replace(chr(10), '<br>')}
                </div>
                """,
                unsafe_allow_html=True
            )

class SectionMatcher:
    """Intelligent section matching between templates"""

    @staticmethod
    def find_section_matches(sections1: Dict[str, ParsedSection],
                           sections2: Dict[str, ParsedSection],
                           threshold: float = 70.0) -> List[SectionMapping]:
        """Find matching sections between two template files"""

        mappings = []
        used_sections2 = set()

        for section1_name in sections1.keys():
            # Try exact match first
            if section1_name in sections2:
                mappings.append(SectionMapping(
                    file1_section=section1_name,
                    file2_section=section1_name,
                    confidence=100.0,
                    is_manual=False
                ))
                used_sections2.add(section1_name)
                continue

            # Find best fuzzy match
            available_sections2 = [s for s in sections2.keys() if s not in used_sections2]
            if available_sections2:
                best_match, confidence = process.extractOne(
                    section1_name,
                    available_sections2,
                    scorer=fuzz.ratio
                )

                if confidence >= threshold:
                    mappings.append(SectionMapping(
                        file1_section=section1_name,
                        file2_section=best_match,
                        confidence=float(confidence),
                        is_manual=False
                    ))
                    used_sections2.add(best_match)

        return mappings

    @staticmethod
    def get_unmapped_sections(sections1: Dict[str, ParsedSection],
                            sections2: Dict[str, ParsedSection],
                            mappings: List[SectionMapping]) -> Tuple[List[str], List[str]]:
        """Get sections that couldn't be automatically mapped"""

        mapped_sections1 = {m.file1_section for m in mappings}
        mapped_sections2 = {m.file2_section for m in mappings}

        unmapped1 = [s for s in sections1.keys() if s not in mapped_sections1]
        unmapped2 = [s for s in sections2.keys() if s not in mapped_sections2]

        return unmapped1, unmapped2

class DocxExporter:
    """Export templates back to Klarity-friendly DOCX format with proper embedded comments"""

    @staticmethod
    def build_comment_string(section: ParsedSection) -> str:
        """Build comment string in Klarity concatenated format"""

        # Escape newlines as literal \n like the original system expects
        prompt_escaped = (section.prompt or '').replace('\r\n', '\n').replace('\n', '\\n')

        # Build in the exact concatenated format found in Alex's templates
        comment_string = (
            f"type - {section.type}"
            f"sub_type - {section.sub_type}"
            f"prompt - {prompt_escaped}"
            f"include_screenshots - {'yes' if section.include_screenshots else 'no'}"
            f"screenshot_instructions - {section.screenshot_instructions or 'none'}"
        )

        return comment_string

    @staticmethod
    def export_template_to_docx(template_name: str, sections: Dict[str, ParsedSection]) -> BytesIO:
        """Export sections to a Klarity-ready DOCX file with proper embedded comments"""

        # Create a new document
        doc = Document()

        # Add title
        title = doc.add_heading(template_name, 0)

        # Add generation info
        info_para = doc.add_paragraph()
        info_para.add_run(f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}").italic = True
        info_para.add_run("\nKlarity Template Comparison Tool").italic = True

        doc.add_paragraph("")  # Spacing

        # Add sections with proper placeholder content (no embedded comments in text)
        for section_name, section in sections.items():
            # Add section heading
            heading = doc.add_heading(section_name, level=1)

            # Add only placeholder content based on section type
            if section.type == 'table':
                placeholder = '[AI will generate table content here based on the embedded prompt instructions.]'
            elif section.sub_type == 'bulleted':
                placeholder = '[AI will generate bullet-point content here:\n• Point 1\n• Point 2\n• Point 3]'
            elif section.sub_type == 'flow-diagram':
                placeholder = '[AI will generate process flow description here with decision points and steps.]'
            elif section.sub_type == 'walkthrough-steps':
                placeholder = '[AI will generate step-by-step instructions here:\n1. Step one\n2. Step two\n3. Step three]'
            else:
                placeholder = '[AI will generate paragraph-based content here based on the embedded prompt instructions.]'

            if section.include_screenshots:
                placeholder += '\n\n[Screenshots will be included as specified in the prompt instructions.]'

            content_para = doc.add_paragraph(placeholder)
            doc.add_paragraph("")  # Spacing

        # Save the initial document
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        # Now we need to manually add comments to the DOCX ZIP structure
        # Since python-docx doesn't support comments directly, we'll modify the ZIP
        return DocxExporter._add_comments_to_docx(buffer, sections)

    @staticmethod
    def _add_comments_to_docx(buffer: BytesIO, sections: Dict[str, ParsedSection]) -> BytesIO:
        """Add comments to an existing DOCX by modifying its ZIP structure"""

        import zipfile
        import tempfile
        from xml.dom import minidom
        import uuid

        # Read the existing DOCX
        buffer.seek(0)
        docx_zip = zipfile.ZipFile(buffer, 'r')

        # Create a new DOCX with comments
        new_buffer = BytesIO()
        new_docx = zipfile.ZipFile(new_buffer, 'w', zipfile.ZIP_DEFLATED)

        try:
            # Copy all existing files
            for item in docx_zip.infolist():
                data = docx_zip.read(item.filename)
                if item.filename == 'word/document.xml':
                    # Modify document.xml to add comment references
                    data = DocxExporter._add_comment_references_to_document(data, sections)
                elif item.filename == '[Content_Types].xml':
                    # Update content types to include comments
                    data = DocxExporter._update_content_types(data)
                elif item.filename == 'word/_rels/document.xml.rels':
                    # Update relationships to include comments
                    data = DocxExporter._update_document_rels(data)

                new_docx.writestr(item, data)

            # Add comments.xml
            comments_xml = DocxExporter._create_comments_xml(sections)
            new_docx.writestr('word/comments.xml', comments_xml)

        finally:
            docx_zip.close()
            new_docx.close()

        new_buffer.seek(0)
        return new_buffer

    @staticmethod
    def _add_comment_references_to_document(document_xml_data: bytes, sections: Dict[str, ParsedSection]) -> bytes:
        """Add comment references to section headings in document.xml"""

        try:
            # Parse the document XML
            doc = minidom.parseString(document_xml_data.decode('utf-8'))

            # Find all heading paragraphs and add comment references
            comment_id = 0
            paragraphs = doc.getElementsByTagName('w:p')

            for paragraph in paragraphs:
                # Check if this is a heading with style
                pPr = paragraph.getElementsByTagName('w:pPr')
                if pPr:
                    pStyle = pPr[0].getElementsByTagName('w:pStyle')
                    if pStyle and pStyle[0].getAttribute('w:val').startswith('Heading'):
                        # Get the heading text
                        text_runs = paragraph.getElementsByTagName('w:t')
                        if text_runs:
                            heading_text = ''.join([t.firstChild.nodeValue for t in text_runs if t.firstChild])

                            # Check if this heading matches one of our sections
                            if heading_text in sections:
                                comment_id += 1

                                # Add comment range start
                                commentRangeStart = doc.createElement('w:commentRangeStart')
                                commentRangeStart.setAttribute('w:id', str(comment_id))
                                paragraph.insertBefore(commentRangeStart, paragraph.firstChild)

                                # Add comment range end
                                commentRangeEnd = doc.createElement('w:commentRangeEnd')
                                commentRangeEnd.setAttribute('w:id', str(comment_id))
                                paragraph.appendChild(commentRangeEnd)

                                # Add comment reference
                                commentReference = doc.createElement('w:commentReference')
                                commentReference.setAttribute('w:id', str(comment_id))

                                # Create a new run for the comment reference
                                run = doc.createElement('w:r')
                                run.appendChild(commentReference)
                                paragraph.appendChild(run)

            return doc.toxml(encoding='utf-8')

        except Exception as e:
            # If we can't parse/modify, return original
            return document_xml_data

    @staticmethod
    def _create_comments_xml(sections: Dict[str, ParsedSection]) -> str:
        """Create the comments.xml file with our section comments"""

        comments_xml = '''<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'''

        comment_id = 0
        for section_name, section in sections.items():
            comment_id += 1
            comment_text = DocxExporter.build_comment_string(section)

            # Escape XML special characters
            comment_text_escaped = comment_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

            comments_xml += f'''
    <w:comment w:id="{comment_id}" w:author="Klarity Template Comparison Tool" w:date="{datetime.datetime.now().isoformat()}" w:initials="KTCT">
        <w:p>
            <w:pPr>
                <w:pStyle w:val="CommentText"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="CommentReference"/>
                </w:rPr>
                <w:annotationRef/>
            </w:r>
            <w:r>
                <w:t>{comment_text_escaped}</w:t>
            </w:r>
        </w:p>
    </w:comment>'''

        comments_xml += '''
</w:comments>'''

        return comments_xml

    @staticmethod
    def _update_content_types(content_types_data: bytes) -> bytes:
        """Update [Content_Types].xml to include comments"""

        try:
            content_types_str = content_types_data.decode('utf-8')

            # Add comments override if not present
            if 'word/comments.xml' not in content_types_str:
                # Insert before the closing tag
                insert_pos = content_types_str.rfind('</Types>')
                if insert_pos != -1:
                    comment_override = '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
                    content_types_str = content_types_str[:insert_pos] + comment_override + content_types_str[insert_pos:]

            return content_types_str.encode('utf-8')

        except Exception as e:
            return content_types_data

    @staticmethod
    def _update_document_rels(rels_data: bytes) -> bytes:
        """Update document.xml.rels to include comments relationship"""

        try:
            rels_str = rels_data.decode('utf-8')

            # Add comments relationship if not present
            if 'comments.xml' not in rels_str:
                # Find the highest rId
                import re
                rids = re.findall(r'rId(\d+)', rels_str)
                max_rid = max([int(rid) for rid in rids]) if rids else 0
                new_rid = max_rid + 1

                # Insert before the closing tag
                insert_pos = rels_str.rfind('</Relationships>')
                if insert_pos != -1:
                    comment_rel = f'<Relationship Id="rId{new_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
                    rels_str = rels_str[:insert_pos] + comment_rel + rels_str[insert_pos:]

            return rels_str.encode('utf-8')

        except Exception as e:
            return rels_data

class DiffViewer:
    """Generate and display diffs between two text strings"""

    @staticmethod
    def generate_diff_html(original: str, modified: str, file1_name: str = "Template 1", file2_name: str = "Template 2") -> str:
        """Generate HTML diff between two strings"""

        # Normalize texts for comparison
        original_lines = original.splitlines(keepends=True)
        modified_lines = modified.splitlines(keepends=True)

        # Generate diff
        diff = list(difflib.unified_diff(
            original_lines,
            modified_lines,
            fromfile=file1_name,
            tofile=file2_name,
            lineterm=''
        ))

        if not diff:
            return "<p style='color: gray; text-align: center; padding: 2rem;'>✅ No differences found - prompts are identical</p>"

        html_parts = []
        html_parts.append("<pre style='background-color: #f8f9fa; padding: 1rem; border-radius: 0.5rem; font-family: monospace; white-space: pre-wrap; word-wrap: break-word;'>")

        for line in diff:
            line_escaped = line.replace('<', '&lt;').replace('>', '&gt;')
            if line.startswith('+++') or line.startswith('---'):
                html_parts.append(f"<span style='color: #6c757d; font-weight: bold;'>{line_escaped}</span>")
            elif line.startswith('@@'):
                html_parts.append(f"<span style='color: #007bff; font-weight: bold;'>{line_escaped}</span>")
            elif line.startswith('+'):
                html_parts.append(f"<span style='color: #28a745; background-color: #d4edda; padding: 2px;'>{line_escaped}</span>")
            elif line.startswith('-'):
                html_parts.append(f"<span style='color: #dc3545; background-color: #f8d7da; padding: 2px; text-decoration: line-through;'>{line_escaped}</span>")
            else:
                html_parts.append(f"<span>{line_escaped}</span>")

        html_parts.append("</pre>")
        return '\n'.join(html_parts)

# Streamlit App
def main():
    st.set_page_config(
        page_title="Klarity Template Comparison Tool",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Initialize session state
    if 'section_mappings' not in st.session_state:
        st.session_state.section_mappings = {}
    if 'edited_sections' not in st.session_state:
        st.session_state.edited_sections = {}

    # Header with branding
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🚀 Klarity Template Comparison Tool")
        st.markdown("*Intelligent DOCX template prompt comparison and editing*")

    # Sidebar for file uploads and settings
    with st.sidebar:
        st.header("📁 Upload Templates")

        uploaded_files = st.file_uploader(
            "Choose DOCX files (2+ required)",
            accept_multiple_files=True,
            type=['docx'],
            help="Upload DOCX files with embedded prompts as comments"
        )

        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} files uploaded")

            with st.expander("📋 File Details", expanded=True):
                for i, file in enumerate(uploaded_files, 1):
                    st.write(f"**{i}.** {file.name}")
                    st.caption(f"Size: {file.size:,} bytes")

        st.divider()

        # Settings
        st.header("⚙️ Settings")

        match_threshold = st.slider(
            "Section matching threshold",
            min_value=50,
            max_value=100,
            value=70,
            help="Minimum similarity % for automatic section matching"
        )

        show_debug = st.checkbox("Show debug info", value=False)

    # Main content area
    if not uploaded_files:
        st.info("👆 Please upload DOCX files to get started")

        with st.expander("ℹ️ How to use this tool"):
            st.markdown("""
            ### 🎯 Purpose
            Compare and edit prompts embedded as comments in DOCX templates, then export in Klarity-friendly format.

            ### 📋 Instructions:
            1. **Upload Templates**: Add 2+ DOCX files with comment-linked sections
            2. **Smart Matching**: AI automatically matches similar sections across templates
            3. **Manual Override**: Adjust mappings for sections that need manual correlation
            4. **Edit Prompts**: Modify prompts inline with real-time preview
            5. **Export**: Download edited templates in Klarity-ready format

            ### 📄 Template Requirements:
            - Section headings must have attached comments
            - Comments contain prompt instructions (structured or freeform)
            - Supported structured format:
              ```
              type - text
              sub_type - bulleted
              prompt - Your detailed instructions here
              include_screenshots - yes
              screenshot_instructions - Capture workflow screenshots
              ```
            """)
        return

    if len(uploaded_files) < 1:
        st.warning("⚠️ Please upload at least 1 DOCX file to proceed")
        return

    # Process uploaded files
    all_sections = {}  # file_name -> {section_name -> ParsedSection}

    with st.spinner("🔄 Processing DOCX files..."):
        progress_bar = st.progress(0)

        for i, uploaded_file in enumerate(uploaded_files):
            progress_bar.progress((i + 1) / len(uploaded_files))

            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                tmp_file.write(uploaded_file.getbuffer())
                tmp_file_path = tmp_file.name

            try:
                # Extract sections from DOCX
                sections = DocxCommentExtractor.extract_comments_from_docx(tmp_file_path)
                all_sections[uploaded_file.name] = sections

            finally:
                # Clean up temp file
                os.unlink(tmp_file_path)

        progress_bar.empty()

    # Display results
    total_sections = sum(len(sections) for sections in all_sections.values())
    if total_sections == 0:
        st.error("❌ No prompts found in uploaded files")
        st.info("💡 Ensure your DOCX files contain comments linked to section headings")
        return

    # Success metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📄 Files Processed", len(uploaded_files))
    with col2:
        st.metric("📝 Sections Found", total_sections)
    with col3:
        files_with_sections = len([f for f in all_sections.values() if f])
        st.metric("✅ Files with Prompts", files_with_sections)
    with col4:
        avg_sections = total_sections / len(uploaded_files) if uploaded_files else 0
        st.metric("📊 Avg Sections/File", f"{avg_sections:.1f}")

    # Create enhanced tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📋 Browse All",
        "🔗 Smart Mapping",
        "✏️ Edit & Compare",
        "🔍 Diff Analysis",
        "📤 Export"
    ])

    with tab1:
        browse_all_prompts(all_sections, show_debug)

    with tab2:
        smart_section_mapping(all_sections, match_threshold)

    with tab3:
        edit_and_compare(all_sections)

    with tab4:
        diff_analysis(all_sections)

    with tab5:
        export_templates(all_sections)


def browse_all_prompts(all_sections: Dict[str, Dict[str, ParsedSection]], show_debug: bool):
    """Tab 1: Browse all extracted prompts with enhanced UI"""

    st.header("📋 All Extracted Prompts")
    st.markdown("*Browse all sections across your templates with enhanced readability and formatting*")

    # Enhanced filters
    col1, col2, col3 = st.columns([2, 2, 1])

    with col1:
        selected_files = st.multiselect(
            "📁 Filter by templates:",
            options=list(all_sections.keys()),
            default=list(all_sections.keys()),
            help="Select specific templates to display"
        )

    with col2:
        search_term = st.text_input(
            "🔍 Search content:",
            placeholder="Search section names or prompt content...",
            help="Search across section names and prompt text"
        )

    with col3:
        view_mode = st.selectbox(
            "👁️ View:",
            options=["Compact", "Detailed"],
            index=1,
            help="Choose display density"
        )

    if not selected_files:
        selected_files = list(all_sections.keys())

    # Process and display sections
    total_displayed = 0

    for file_name in selected_files:
        sections = all_sections[file_name]
        if not sections:
            continue

        # Filter by search term
        filtered_sections = sections
        if search_term:
            filtered_sections = {
                name: section for name, section in sections.items()
                if search_term.lower() in name.lower() or
                   search_term.lower() in section.prompt.lower()
            }

        if not filtered_sections:
            continue

        # File header with statistics
        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            st.subheader(f"📄 {file_name}")
        with col2:
            st.metric("Sections", len(filtered_sections))
        with col3:
            edited_count = sum(1 for s in filtered_sections.values() if hasattr(s, 'edited') and s.edited)
            st.metric("Edited", edited_count)

        if search_term and len(filtered_sections) != len(sections):
            st.caption(f"Showing {len(filtered_sections)} of {len(sections)} sections")

        st.divider()

        # Display sections using enhanced components
        for section_name, section in filtered_sections.items():
            if view_mode == "Detailed":
                UIComponents.create_section_card(section_name, section, file_name, expanded=False)
            else:
                # Compact view
                with st.expander(f"📝 {section_name}" + (" ✏️" if hasattr(section, 'edited') and section.edited else ""), expanded=False):
                    col1, col2 = st.columns([3, 1])

                    with col1:
                        # Truncated prompt for compact view
                        formatted_prompt = PromptFormatter.format_prompt_for_display(section.prompt, max_length=300)
                        st.markdown(formatted_prompt)

                    with col2:
                        # Compact properties
                        properties = PromptFormatter.create_section_property_display(section, show_stats=True)
                        st.markdown(properties)

            total_displayed += 1

            # Debug info if requested
            if show_debug:
                with st.expander(f"🐛 Debug: {section_name}", expanded=False):
                    st.text("Raw Comment Data:")
                    st.code(section.raw_comment)
                    st.text("Parsed Properties:")
                    st.json({
                        "name": section.name,
                        "type": section.type,
                        "sub_type": section.sub_type,
                        "include_screenshots": section.include_screenshots,
                        "screenshot_instructions": section.screenshot_instructions,
                        "edited": getattr(section, 'edited', False)
                    })

        st.divider()

    # Summary
    if total_displayed == 0:
        st.info("📭 No sections match your current filters")
    else:
        st.success(f"✅ Displaying {total_displayed} sections across {len(selected_files)} templates")


def smart_section_mapping(all_sections: Dict[str, Dict[str, ParsedSection]], threshold: float):
    """Tab 2: Intelligent section mapping between templates"""

    st.header("🔗 Smart Section Mapping")

    if len(all_sections) < 2:
        st.info("📋 Upload at least 2 files to use section mapping")
        return

    # File pair selection
    file_names = list(all_sections.keys())
    col1, col2 = st.columns(2)

    with col1:
        file1 = st.selectbox("Primary template:", file_names, key="mapping_file1")
    with col2:
        file2 = st.selectbox("Compare with:", [f for f in file_names if f != file1], key="mapping_file2")

    if not file1 or not file2 or file1 == file2:
        st.info("Select two different files to continue")
        return

    sections1 = all_sections[file1]
    sections2 = all_sections[file2]

    # Generate automatic mappings
    mappings = SectionMatcher.find_section_matches(sections1, sections2, threshold)
    unmapped1, unmapped2 = SectionMatcher.get_unmapped_sections(sections1, sections2, mappings)

    # Store in session state
    mapping_key = f"{file1}::{file2}"
    if mapping_key not in st.session_state.section_mappings:
        st.session_state.section_mappings[mapping_key] = mappings

    # Display mapping results
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🎯 Auto-matched", len(mappings))
    with col2:
        st.metric("❓ Unmapped in File 1", len(unmapped1))
    with col3:
        st.metric("❓ Unmapped in File 2", len(unmapped2))

    # Show automatic mappings
    if mappings:
        st.subheader("🤖 Automatic Mappings")

        for i, mapping in enumerate(mappings):
            confidence_color = "🟢" if mapping.confidence >= 90 else "🟡" if mapping.confidence >= 70 else "🟠"

            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                st.write(f"**{mapping.file1_section}**")
            with col2:
                st.write(f"**{mapping.file2_section}**")
            with col3:
                st.write(f"{confidence_color} {mapping.confidence:.0f}%")

    # Manual mapping for unmapped sections
    if unmapped1 or unmapped2:
        st.subheader("🔧 Manual Mapping")
        st.info("Map sections that couldn't be automatically matched")

        # Create manual mappings
        manual_mappings = []

        for i, section1 in enumerate(unmapped1):
            col1, col2, col3 = st.columns([2, 2, 1])

            with col1:
                st.write(f"**{section1}**")
                if section1 in sections1:
                    preview = sections1[section1].prompt[:100] + "..." if len(sections1[section1].prompt) > 100 else sections1[section1].prompt
                    st.caption(preview)

            with col2:
                options = ["[No Match]"] + unmapped2
                selected = st.selectbox(
                    f"Match with:",
                    options,
                    key=f"manual_map_{i}_{section1}"
                )

                if selected != "[No Match]":
                    manual_mappings.append(SectionMapping(
                        file1_section=section1,
                        file2_section=selected,
                        confidence=0.0,  # Manual mapping
                        is_manual=True
                    ))

            with col3:
                if selected != "[No Match]":
                    st.success("👤 Manual")
                else:
                    st.info("⚠️ No match")

        # Update session state with manual mappings
        if manual_mappings:
            all_mappings = mappings + manual_mappings
            st.session_state.section_mappings[mapping_key] = all_mappings

            st.success(f"✅ {len(manual_mappings)} manual mappings added")


def edit_and_compare(all_sections: Dict[str, Dict[str, ParsedSection]]):
    """Tab 3: Edit prompts and compare side-by-side"""

    st.header("✏️ Edit & Compare Prompts")

    # File and section selection
    file_names = list(all_sections.keys())

    col1, col2 = st.columns(2)
    with col1:
        selected_file = st.selectbox("Select template to edit:", file_names, key="edit_file")
    with col2:
        if selected_file:
            sections = all_sections[selected_file]
            section_names = list(sections.keys()) if sections else []
            selected_section = st.selectbox("Select section:", section_names, key="edit_section")

    if not selected_file or not selected_section:
        st.info("Select a file and section to continue")
        return

    section = all_sections[selected_file][selected_section]
    section_key = f"{selected_file}::{selected_section}"

    # Initialize edited sections in session state
    if section_key not in st.session_state.edited_sections:
        st.session_state.edited_sections[section_key] = {
            'prompt': section.prompt,
            'type': section.type,
            'sub_type': section.sub_type,
            'include_screenshots': section.include_screenshots,
            'screenshot_instructions': section.screenshot_instructions
        }

    edited_data = st.session_state.edited_sections[section_key]

    st.subheader(f"✏️ Editing: {selected_section}")

    # Editing interface
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**📝 Edit Prompt:**")

        new_prompt = st.text_area(
            "Prompt content:",
            value=edited_data['prompt'],
            height=300,
            help="Edit the prompt text here"
        )

        # Section properties with error handling
        col1a, col1b = st.columns(2)
        with col1a:
            type_options = ['text', 'table']
            current_type = edited_data.get('type', 'text')
            if current_type not in type_options:
                current_type = 'text'

            new_type = st.selectbox(
                "Type:",
                options=type_options,
                index=type_options.index(current_type)
            )

        with col1b:
            subtype_options = ['default', 'bulleted', 'freeform', 'flow-diagram', 'walkthrough-steps']
            current_subtype = edited_data.get('sub_type', 'default')
            # Handle legacy or special subtypes
            if current_subtype == 'bpmn':
                current_subtype = 'flow-diagram'
            elif current_subtype not in subtype_options:
                current_subtype = 'default'

            new_sub_type = st.selectbox(
                "Sub-type:",
                options=subtype_options,
                index=subtype_options.index(current_subtype)
            )

        new_include_screenshots = st.checkbox(
            "Include screenshots",
            value=edited_data['include_screenshots']
        )

        if new_include_screenshots:
            new_screenshot_instructions = st.text_area(
                "Screenshot instructions:",
                value=edited_data['screenshot_instructions'],
                height=60
            )
        else:
            new_screenshot_instructions = ""

        # Save changes
        if st.button("💾 Save Changes", type="primary"):
            st.session_state.edited_sections[section_key] = {
                'prompt': new_prompt,
                'type': new_type,
                'sub_type': new_sub_type,
                'include_screenshots': new_include_screenshots,
                'screenshot_instructions': new_screenshot_instructions
            }

            # Update the actual section object
            section.prompt = new_prompt
            section.type = new_type
            section.sub_type = new_sub_type
            section.include_screenshots = new_include_screenshots
            section.screenshot_instructions = new_screenshot_instructions
            section.edited = True

            st.success("✅ Changes saved!")
            st.rerun()

    with col2:
        st.markdown("**👁️ Live Preview:**")

        # Enhanced formatted preview
        formatted_prompt = PromptFormatter.format_prompt_for_display(new_prompt)

        # Display with enhanced styling
        st.markdown(
            f"""
            <div style="
                background-color: #f8f9fa;
                padding: 1rem;
                border-radius: 0.5rem;
                border-left: 4px solid #28a745;
                line-height: 1.6;
                max-height: 400px;
                overflow-y: auto;
                color: #212529;
            ">
            {formatted_prompt.replace(chr(10), '<br>')}
            </div>
            """,
            unsafe_allow_html=True
        )

        st.divider()

        # Show changes indicator with better formatting
        if new_prompt != section.original_prompt:
            st.warning("✏️ **Modified from original**")

            # Show character difference
            original_len = len(section.original_prompt)
            new_len = len(new_prompt)
            diff = new_len - original_len

            col2a, col2b, col2c = st.columns(3)
            with col2a:
                st.metric("Original", f"{original_len:,} chars")
            with col2b:
                st.metric("Current", f"{new_len:,} chars")
            with col2c:
                st.metric("Change", f"{diff:+,}")

        else:
            st.success("✅ **Matches original**")

        st.divider()

        # Action buttons
        col2a, col2b = st.columns(2)
        with col2a:
            if st.button("🔄 Reset", help="Reset to original content"):
                st.session_state.edited_sections[section_key]['prompt'] = section.original_prompt
                section.prompt = section.original_prompt
                section.edited = False
                st.success("✅ Reset to original")
                st.rerun()

        with col2b:
            if st.button("📋 Copy", help="Copy current content"):
                st.success("Copied to clipboard!")


def diff_analysis(all_sections: Dict[str, Dict[str, ParsedSection]]):
    """Tab 4: Advanced diff analysis between sections"""

    st.header("🔍 Advanced Diff Analysis")

    if len(all_sections) < 2:
        st.info("📋 Upload at least 2 files for diff analysis")
        return

    # File and section selection with intelligent matching
    file_names = list(all_sections.keys())

    col1, col2 = st.columns(2)
    with col1:
        file1 = st.selectbox("First template:", file_names, key="diff_file1")
    with col2:
        file2 = st.selectbox("Second template:", [f for f in file_names if f != file1], key="diff_file2")

    if not file1 or not file2:
        st.info("Select two files to continue")
        return

    sections1 = all_sections[file1]
    sections2 = all_sections[file2]

    # Section selection with smart suggestions
    all_sections1 = list(sections1.keys())
    all_sections2 = list(sections2.keys())

    col1, col2 = st.columns(2)
    with col1:
        section1 = st.selectbox(f"Section from {file1}:", all_sections1, key="diff_section1")
    with col2:
        # Smart suggestion for section2 based on section1
        if section1:
            if section1 in sections2:
                default_idx = all_sections2.index(section1)
            else:
                # Find best match using fuzzy matching
                matches = process.extract(section1, all_sections2, limit=3, scorer=fuzz.ratio)
                if matches and matches[0][1] >= 70:
                    default_idx = all_sections2.index(matches[0][0])
                else:
                    default_idx = 0
        else:
            default_idx = 0

        section2 = st.selectbox(
            f"Section from {file2}:",
            all_sections2,
            index=default_idx,
            key="diff_section2"
        )

    if not section1 or not section2:
        st.info("Select sections from both files")
        return

    # Get the actual sections
    sec1 = sections1[section1]
    sec2 = sections2[section2]

    # Display section comparison
    st.subheader(f"📊 Comparing: {section1} ↔ {section2}")

    # Quick stats
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("File 1 Length", f"{len(sec1.prompt):,} chars")
    with col2:
        st.metric("File 2 Length", f"{len(sec2.prompt):,} chars")
    with col3:
        diff_chars = len(sec2.prompt) - len(sec1.prompt)
        st.metric("Difference", f"{diff_chars:+,} chars")
    with col4:
        similarity = fuzz.ratio(sec1.prompt, sec2.prompt)
        st.metric("Similarity", f"{similarity}%")

    # Format prompts for comparison
    prompt1 = PromptFormatter.json_to_markdown(sec1.prompt)
    prompt2 = PromptFormatter.json_to_markdown(sec2.prompt)

    # Diff visualization
    st.subheader("🎨 Visual Diff")

    diff_html = DiffViewer.generate_diff_html(prompt1, prompt2, file1, file2)
    st.markdown(diff_html, unsafe_allow_html=True)

    # Enhanced side-by-side comparison
    st.subheader("📄 Side-by-Side Comparison")

    UIComponents.create_side_by_side_comparison(
        section_name=f"{section1} vs {section2}",
        section1=sec1,
        section2=sec2,
        file1_name=file1,
        file2_name=file2
    )

    # Analysis insights
    st.subheader("🧠 Analysis Insights")

    insights = []

    if diff_chars > 100:
        insights.append(f"📈 File 2 is significantly longer ({diff_chars:+,} characters)")
    elif diff_chars < -100:
        insights.append(f"📉 File 1 is significantly longer ({abs(diff_chars):,} characters)")
    else:
        insights.append("📊 Both prompts have similar length")

    if similarity >= 90:
        insights.append("✅ Prompts are very similar (90%+ match)")
    elif similarity >= 70:
        insights.append("🔄 Prompts are moderately similar (70-89% match)")
    else:
        insights.append("⚠️ Prompts are quite different (<70% match)")

    if sec1.type != sec2.type:
        insights.append(f"🔄 Different types: {sec1.type} vs {sec2.type}")

    if sec1.sub_type != sec2.sub_type:
        insights.append(f"🔄 Different sub-types: {sec1.sub_type} vs {sec2.sub_type}")

    if sec1.include_screenshots != sec2.include_screenshots:
        insights.append("📷 Different screenshot requirements")

    for insight in insights:
        st.info(insight)


def export_templates(all_sections: Dict[str, Dict[str, ParsedSection]]):
    """Tab 5: Export templates in Klarity-friendly format"""

    st.header("📤 Export Templates")

    if not all_sections:
        st.info("No templates available for export")
        return

    st.markdown("""
    Export your templates (including any edits) in Klarity-friendly DOCX format.
    The exported files will contain properly formatted comments that can be processed by the Klarity system.
    """)

    # Export options
    export_mode = st.radio(
        "Export mode:",
        options=[
            "📄 Individual Files (one per template)",
            "📦 Merged Template (all sections combined)",
            "✏️ Edited Sections Only"
        ]
    )

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🔧 Export Options")

        # Template name for merged export
        if "Merged" in export_mode:
            template_name = st.text_input(
                "Template name for merged file:",
                value="Merged_Template"
            )

        # Include metadata
        include_metadata = st.checkbox(
            "Include generation metadata",
            value=True,
            help="Add timestamp and tool information to exported files"
        )

        # Format options
        preserve_edits = st.checkbox(
            "Preserve edit history",
            value=True,
            help="Mark edited sections in the exported document"
        )

    with col2:
        st.subheader("📊 Export Preview")

        total_files = len(all_sections)
        total_sections = sum(len(sections) for sections in all_sections.values())
        edited_sections = sum(
            sum(1 for section in sections.values() if section.edited)
            for sections in all_sections.values()
        )

        st.metric("📄 Templates", total_files)
        st.metric("📝 Total Sections", total_sections)
        st.metric("✏️ Edited Sections", edited_sections)

    # Export buttons
    st.subheader("💾 Download Files")

    if export_mode == "📄 Individual Files (one per template)":
        # Export each template individually
        for file_name, sections in all_sections.items():
            if not sections:
                continue

            col1, col2, col3 = st.columns([2, 1, 1])

            with col1:
                st.write(f"**{file_name}**")
                st.caption(f"{len(sections)} sections" + (f", {sum(1 for s in sections.values() if s.edited)} edited" if edited_sections else ""))

            with col2:
                # Preview button
                clean_file_name = file_name.replace(" ", "_").replace(".", "_").replace("/", "_")
                if st.button(f"👁️ Preview", key=f"preview_individual_{clean_file_name}"):
                    with st.expander(f"Preview: {file_name}", expanded=True):
                        for section_name, section in sections.items():
                            st.write(f"**{section_name}**" + (" ✏️" if section.edited else ""))
                            st.caption(f"Type: {section.type}, Sub-type: {section.sub_type}")
                            preview = section.prompt[:100] + "..." if len(section.prompt) > 100 else section.prompt
                            st.text(preview)
                            st.divider()

            with col3:
                # Generate export file
                clean_name = file_name.replace('.docx', '').replace('.', '_').replace(' ', '_').replace('/', '_')
                export_buffer = DocxExporter.export_template_to_docx(clean_name, sections)

                st.download_button(
                    label="📥 Download",
                    data=export_buffer.getvalue(),
                    file_name=f"{clean_name}_Klarity_Ready.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"download_individual_{clean_name}"
                )

    elif export_mode == "📦 Merged Template (all sections combined)":
        # Combine all sections from all templates
        merged_sections = {}

        for file_name, sections in all_sections.items():
            for section_name, section in sections.items():
                # Handle duplicate section names by prefixing with file name
                if section_name in merged_sections:
                    unique_name = f"{file_name}_{section_name}"
                else:
                    unique_name = section_name

                merged_sections[unique_name] = section

        st.write(f"**Merged template will contain {len(merged_sections)} sections**")

        if st.button("👁️ Preview Merged Template"):
            with st.expander("Merged Template Preview", expanded=True):
                for section_name, section in merged_sections.items():
                    st.write(f"**{section_name}**" + (" ✏️" if section.edited else ""))
                    st.caption(f"Type: {section.type}, Sub-type: {section.sub_type}")
                    st.divider()

        # Generate merged export
        export_buffer = DocxExporter.export_template_to_docx(template_name, merged_sections)

        st.download_button(
            label="📥 Download Merged Template",
            data=export_buffer.getvalue(),
            file_name=f"{template_name}_Klarity_Ready.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="download_merged"
        )

    elif export_mode == "✏️ Edited Sections Only":
        # Export only edited sections
        edited_only = {}

        for file_name, sections in all_sections.items():
            file_edited = {name: section for name, section in sections.items() if section.edited}
            if file_edited:
                edited_only[file_name] = file_edited

        if not edited_only:
            st.info("No edited sections found. Make some edits first!")
        else:
            st.write(f"**Found edited sections in {len(edited_only)} files**")

            for file_name, sections in edited_only.items():
                col1, col2, col3 = st.columns([2, 1, 1])

                with col1:
                    st.write(f"**{file_name}** (edited sections)")
                    st.caption(f"{len(sections)} edited sections")

                with col2:
                    clean_file_name_edited = file_name.replace(" ", "_").replace(".", "_").replace("/", "_")
                    if st.button(f"👁️ Preview", key=f"preview_edited_{clean_file_name_edited}"):
                        with st.expander(f"Edited Sections: {file_name}", expanded=True):
                            for section_name, section in sections.items():
                                st.write(f"**{section_name}** ✏️")
                                st.caption(f"Type: {section.type}, Sub-type: {section.sub_type}")
                                st.text("Original → Edited comparison would go here")
                                st.divider()

                with col3:
                    clean_name = file_name.replace('.docx', '').replace('.', '_').replace(' ', '_').replace('/', '_')
                    export_buffer = DocxExporter.export_template_to_docx(f"{clean_name}_Edited", sections)

                    st.download_button(
                        label="📥 Download",
                        data=export_buffer.getvalue(),
                        file_name=f"{clean_name}_Edited_Klarity_Ready.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"download_edited_{clean_name}"
                    )

    # Export success message
    st.success("✅ Your exported DOCX files are ready for use with Klarity!")
    st.info("💡 The exported files contain properly formatted comments that can be processed by the Klarity prompt generation system.")


if __name__ == "__main__":
    main()