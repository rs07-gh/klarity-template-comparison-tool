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
        """Parse comment text to extract structured section data"""

        try:
            lines = comment_text.split('\n')
            section_data = {
                'type': 'text',
                'sub_type': 'default',
                'prompt': '',
                'include_screenshots': False,
                'screenshot_instructions': ''
            }

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Try different separators: ' - ', ': ', '='
                key = ''
                value = ''

                if ' - ' in line:
                    parts = line.split(' - ', 1)
                    key, value = parts[0].strip(), parts[1].strip()
                elif ': ' in line:
                    parts = line.split(': ', 1)
                    key, value = parts[0].strip(), parts[1].strip()
                elif '=' in line:
                    parts = line.split('=', 1)
                    key, value = parts[0].strip(), parts[1].strip()
                else:
                    # If no separator, treat as prompt content
                    if not section_data['prompt']:
                        section_data['prompt'] = line
                    continue

                # Normalize key
                key_normalized = key.lower().replace('_', '').replace(' ', '')

                if key_normalized in ['type', 'contenttype']:
                    section_data['type'] = value.lower()
                elif key_normalized in ['subtype', 'sub_type', 'style']:
                    section_data['sub_type'] = value.lower()
                elif key_normalized in ['prompt', 'instruction', 'content']:
                    section_data['prompt'] = value.replace('\\n', '\n')
                elif key_normalized in ['includescreenshot', 'include_screenshot', 'screenshot']:
                    section_data['include_screenshots'] = value.lower() in ['yes', 'true', '1']
                elif key_normalized in ['screenshotinstruction', 'screenshot_instruction']:
                    section_data['screenshot_instructions'] = value if value != 'none' else ''

            # If no prompt found in structured format, use entire comment as prompt
            if not section_data['prompt'] and len(comment_text.strip()) > 15:
                section_data['prompt'] = comment_text.strip()

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
            st.error(f"Error parsing comment for section '{section_name}': {str(e)}")
            return None

class PromptFormatter:
    """Convert JSON prompts to readable markdown format"""

    @staticmethod
    def json_to_markdown(prompt_text: str) -> str:
        """Convert JSON string prompts to readable markdown"""

        # Try to parse as JSON first
        try:
            if prompt_text.strip().startswith('{') or prompt_text.strip().startswith('['):
                prompt_obj = json.loads(prompt_text)
                return PromptFormatter._json_obj_to_markdown(prompt_obj)
        except json.JSONDecodeError:
            pass

        # If not JSON, enhance plain text formatting
        return PromptFormatter._enhance_text_formatting(prompt_text)

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
        """Enhance plain text with better markdown formatting"""

        # Add proper line spacing
        text = re.sub(r'\n\s*\n\s*\n+', '\n\n', text)  # Max 2 newlines

        # Convert numbered lists
        text = re.sub(r'(\d+)\.\s*', r'\n\1. ', text)

        # Convert bullet points
        text = re.sub(r'[•·*-]\s*', '- ', text)

        # Emphasize key phrases (words in quotes or caps)
        text = re.sub(r'"([^"]+)"', r'**\1**', text)
        text = re.sub(r'\b([A-Z]{2,})\b', r'**\1**', text)

        return text.strip()

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
    """Export templates back to Klarity-friendly DOCX format"""

    @staticmethod
    def build_comment_string(section: ParsedSection) -> str:
        """Build comment string in Klarity format"""

        # Escape newlines as literal \n like the original system expects
        prompt_escaped = (section.prompt or '').replace('\r\n', '\n').replace('\n', '\\n')

        # Klarity-friendly format matching the original system
        comment_params = [
            f"type - {section.type}",
            f"sub_type - {section.sub_type}",
            f"prompt - {prompt_escaped}",
            f"include_screenshots - {'yes' if section.include_screenshots else 'no'}",
            f"screenshot_instructions - {section.screenshot_instructions or 'none'}"
        ]

        return '\n'.join(comment_params)

    @staticmethod
    def export_template_to_docx(template_name: str, sections: Dict[str, ParsedSection]) -> BytesIO:
        """Export sections to a Klarity-ready DOCX file"""

        # Create a new document
        doc = Document()

        # Add title
        title = doc.add_heading(template_name, 0)

        # Add generation info
        info_para = doc.add_paragraph()
        info_para.add_run(f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}").italic = True
        info_para.add_run("\nKlarity Prompt Generator - Template Comparer").italic = True

        doc.add_paragraph("")  # Spacing

        # Add sections
        for section_name, section in sections.items():
            # Add section heading
            heading = doc.add_heading(section_name, level=1)

            # Add comment to heading (this is a simplified approach)
            # Note: docx library doesn't support comments directly, but we can add
            # the comment data as hidden text or in a structured way
            comment_text = DocxExporter.build_comment_string(section)

            # Add comment as a hidden paragraph (for manual processing)
            comment_para = doc.add_paragraph()
            comment_run = comment_para.add_run(f"[COMMENT: {comment_text}]")
            comment_run.font.color.rgb = None  # This will be processed by Klarity

            # Add placeholder content
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

        # Save to BytesIO
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        return buffer

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
        page_title="Klarity Template Comparer",
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
        st.title("🚀 Klarity Template Comparer")
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
    """Tab 1: Browse all extracted prompts"""

    st.header("📋 All Extracted Prompts")

    # File selector
    selected_files = st.multiselect(
        "Filter by files (empty = show all):",
        list(all_sections.keys()),
        default=list(all_sections.keys())
    )

    if not selected_files:
        selected_files = list(all_sections.keys())

    # Search functionality
    search_term = st.text_input("🔍 Search sections and prompts:")

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

        st.subheader(f"📄 {file_name}")
        st.caption(f"{len(filtered_sections)} sections" + (f" (filtered from {len(sections)})" if search_term else ""))

        for section_name, section in filtered_sections.items():
            with st.expander(f"📝 {section_name}" + (" ✏️" if section.edited else ""), expanded=False):
                col1, col2 = st.columns([3, 1])

                with col1:
                    st.markdown("**Formatted Prompt:**")
                    formatted_prompt = PromptFormatter.json_to_markdown(section.prompt)
                    st.markdown(formatted_prompt)

                    if section.edited:
                        st.info("✏️ This prompt has been edited")

                with col2:
                    st.markdown("**Section Details:**")

                    # Status badges
                    st.write("📊 **Status:**")
                    if section.edited:
                        st.success("✏️ Edited")
                    else:
                        st.info("📄 Original")

                    st.write("🏷️ **Properties:**")
                    st.code(f"""Type: {section.type}
Sub-type: {section.sub_type}
Screenshots: {'Yes' if section.include_screenshots else 'No'}""")

                    if section.screenshot_instructions:
                        st.write("📸 **Screenshot Instructions:**")
                        st.text(section.screenshot_instructions)

                    # Character count
                    char_count = len(section.prompt)
                    word_count = len(section.prompt.split())
                    st.write("📈 **Stats:**")
                    st.code(f"{char_count:,} chars\n{word_count:,} words")

                    if show_debug:
                        with st.expander("🐛 Debug Info", expanded=False):
                            st.text("Raw Comment:")
                            st.code(section.raw_comment)


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

        # Section properties
        col1a, col1b = st.columns(2)
        with col1a:
            new_type = st.selectbox(
                "Type:",
                options=['text', 'table'],
                index=['text', 'table'].index(edited_data['type'])
            )

        with col1b:
            new_sub_type = st.selectbox(
                "Sub-type:",
                options=['default', 'bulleted', 'freeform', 'flow-diagram', 'walkthrough-steps'],
                index=['default', 'bulleted', 'freeform', 'flow-diagram', 'walkthrough-steps'].index(edited_data['sub_type'])
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

        # Show formatted preview
        formatted_prompt = PromptFormatter.json_to_markdown(new_prompt)
        st.markdown(formatted_prompt)

        # Show changes indicator
        if new_prompt != section.original_prompt:
            st.info("✏️ This prompt has been modified")

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
                st.metric("Difference", f"{diff:+,}")

        # Reset button
        if st.button("🔄 Reset to Original"):
            st.session_state.edited_sections[section_key]['prompt'] = section.original_prompt
            section.prompt = section.original_prompt
            section.edited = False
            st.success("✅ Reset to original content")
            st.rerun()


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

    # Side-by-side comparison
    st.subheader("📄 Side-by-Side View")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**📄 {file1}: {section1}**")
        if sec1.edited:
            st.info("✏️ This version has been edited")
        st.markdown(prompt1)

        st.markdown("**Properties:**")
        st.code(f"""Type: {sec1.type}
Sub-type: {sec1.sub_type}
Screenshots: {'Yes' if sec1.include_screenshots else 'No'}""")

    with col2:
        st.markdown(f"**📄 {file2}: {section2}**")
        if sec2.edited:
            st.info("✏️ This version has been edited")
        st.markdown(prompt2)

        st.markdown("**Properties:**")
        st.code(f"""Type: {sec2.type}
Sub-type: {sec2.sub_type}
Screenshots: {'Yes' if sec2.include_screenshots else 'No'}""")

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
                if st.button(f"👁️ Preview", key=f"preview_{file_name}"):
                    with st.expander(f"Preview: {file_name}", expanded=True):
                        for section_name, section in sections.items():
                            st.write(f"**{section_name}**" + (" ✏️" if section.edited else ""))
                            st.caption(f"Type: {section.type}, Sub-type: {section.sub_type}")
                            preview = section.prompt[:100] + "..." if len(section.prompt) > 100 else section.prompt
                            st.text(preview)
                            st.divider()

            with col3:
                # Generate export file
                clean_name = file_name.replace('.docx', '').replace('.', '_')
                export_buffer = DocxExporter.export_template_to_docx(clean_name, sections)

                st.download_button(
                    label="📥 Download",
                    data=export_buffer.getvalue(),
                    file_name=f"{clean_name}_Klarity_Ready.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"download_{file_name}"
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
                    if st.button(f"👁️ Preview", key=f"preview_edited_{file_name}"):
                        with st.expander(f"Edited Sections: {file_name}", expanded=True):
                            for section_name, section in sections.items():
                                st.write(f"**{section_name}** ✏️")
                                st.caption(f"Type: {section.type}, Sub-type: {section.sub_type}")
                                st.text("Original → Edited comparison would go here")
                                st.divider()

                with col3:
                    clean_name = file_name.replace('.docx', '').replace('.', '_')
                    export_buffer = DocxExporter.export_template_to_docx(f"{clean_name}_Edited", sections)

                    st.download_button(
                        label="📥 Download",
                        data=export_buffer.getvalue(),
                        file_name=f"{clean_name}_Edited_Klarity_Ready.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"download_edited_{file_name}"
                    )

    # Export success message
    st.success("✅ Your exported DOCX files are ready for use with Klarity!")
    st.info("💡 The exported files contain properly formatted comments that can be processed by the Klarity prompt generation system.")


if __name__ == "__main__":
    main()