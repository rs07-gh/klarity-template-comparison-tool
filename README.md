# Klarity Template Comparer

A comprehensive Streamlit application for comparing, editing, and managing prompts embedded as comments in DOCX templates. This enterprise-ready tool provides intelligent section matching, inline editing, and Klarity-friendly export capabilities.

## 🚀 Key Features

### Core Functionality
- 📄 **Smart DOCX Processing**: Automatically extracts prompts from comments in DOCX files
- 🧠 **Intelligent Section Matching**: AI-powered fuzzy matching to correlate sections across templates
- ✏️ **Inline Prompt Editing**: Edit prompts directly in the browser with live preview
- 🔍 **Advanced Diff Analysis**: Visual comparison with detailed change highlighting
- 📤 **Klarity Export**: Generate DOCX files ready for Klarity prompt system

### User Experience
- 🎨 **Responsive Design**: Professional UI optimized for organizational use
- 🔍 **Smart Search**: Find sections and prompts across multiple templates
- 📊 **Rich Analytics**: Character counts, similarity scores, and edit tracking
- 🔗 **Section Mapping**: Manual override for sections that need custom correlation
- 📋 **Multi-Mode Export**: Individual files, merged templates, or edited-only versions

## Quick Start

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the App**:
   ```bash
   streamlit run app.py
   ```

3. **Upload DOCX Files**: Upload 2 or more DOCX files with embedded prompts as comments

4. **Compare**: Use the different tabs to view and compare your prompts

## How It Works

### DOCX Template Format
The tool expects DOCX files with:
- Section headings that have comments attached
- Comments containing prompt instructions

### Supported Comment Formats

**Simple Format** (entire comment is the prompt):
```
Write detailed instructions for this section focusing on clarity and completeness.
```

**Structured Format** (key-value pairs):
```
type - text
sub_type - bulleted
prompt - Create a comprehensive bullet list covering all key points
include_screenshots - yes
screenshot_instructions - Include workflow screenshots
```

### Supported Keys in Structured Format
- `type` / `contentType`: text, table
- `sub_type` / `style`: bulleted, freeform, default, flow-diagram, walkthrough-steps
- `prompt` / `instruction` / `content`: The main prompt text
- `include_screenshots` / `screenshot`: yes/no or true/false
- `screenshot_instructions`: Instructions for screenshots

## 🎯 Application Tabs

### 📋 Browse All
- **Smart Search**: Find sections and prompts across all templates
- **File Filtering**: Focus on specific templates
- **Status Tracking**: See which prompts have been edited
- **Rich Details**: View formatted prompts with metadata and statistics

### 🔗 Smart Mapping
- **AI Matching**: Automatic fuzzy matching of similar sections
- **Confidence Scores**: See how confident the AI is about matches
- **Manual Override**: Map sections that couldn't be auto-matched
- **Visual Feedback**: Color-coded confidence indicators

### ✏️ Edit & Compare
- **Live Editing**: Modify prompts with real-time preview
- **Property Management**: Edit section types, sub-types, and screenshot settings
- **Change Tracking**: See character differences and edit history
- **Reset Options**: Easily revert to original content

### 🔍 Diff Analysis
- **Smart Suggestions**: AI suggests the best matching section to compare
- **Visual Diff**: Highlighted additions, removals, and changes
- **Similarity Scoring**: Quantitative similarity analysis
- **Insights Engine**: Automated analysis of differences and patterns

### 📤 Export
- **Multiple Modes**: Individual files, merged templates, or edited-only
- **Klarity-Ready Format**: Properly formatted comments for Klarity processing
- **Metadata Options**: Include generation timestamps and tool info
- **Preview Mode**: See what will be exported before download

## Technical Details

### Core Components

1. **DocxCommentExtractor**: Extracts comments from DOCX files using XML parsing
2. **PromptFormatter**: Converts JSON prompts to readable markdown
3. **DiffViewer**: Generates HTML diffs between text strings

### File Processing
- Uses Python's `zipfile` to read DOCX files
- Parses `word/comments.xml` and `word/document.xml`
- Handles XML namespaces correctly
- Extracts comment text and links to section headings

### Prompt Enhancement
- Automatically converts JSON strings to formatted markdown
- Enhances plain text with better formatting
- Handles bullet points, numbered lists, and emphasis

## Requirements

- Python 3.7+
- Streamlit 1.28.0+
- python-docx 0.8.11+
- lxml 4.9.0+
- markdown 3.5.0+
- fuzzywuzzy 0.18.0+ (for intelligent section matching)
- python-Levenshtein 0.20.0+ (for fast fuzzy matching)

## Usage Examples

### Basic Usage
1. Upload your DOCX template files
2. Navigate to "All Prompts" to see extracted content
3. Use "Side-by-Side Compare" to compare specific sections
4. Use "Diff View" for detailed change analysis

### Team Collaboration
- Share the folder with colleagues
- They can run locally: `streamlit run app.py`
- Compare templates before/after modifications
- Review changes across different template versions

## Troubleshooting

### No Prompts Found
- Ensure DOCX files have comments linked to headings
- Check that comments contain text (not just formatting)
- Verify comment structure if using structured format

### Processing Errors
- Make sure DOCX files are valid (not corrupted)
- Check file permissions
- Try with simpler DOCX files first

### Display Issues
- Refresh the browser if visualizations don't load
- Clear browser cache for styling issues
- Check console for JavaScript errors

## Hosting on Streamlit Cloud

To share with remote team members:

1. **Create GitHub Repository**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin YOUR_GITHUB_REPO_URL
   git push -u origin main
   ```

2. **Deploy to Streamlit Cloud**:
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Connect your GitHub repository
   - Set main file path: `app.py`
   - Deploy!

3. **Share the URL**: Team members can access via the provided URL

## License

This is a standalone tool created for internal team use. Feel free to modify as needed.