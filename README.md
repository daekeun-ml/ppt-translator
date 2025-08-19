# PowerPoint Translator using Amazon Bedrock

A powerful PowerPoint translation tool that leverages Amazon Bedrock models for high-quality translation. This service can be used both as a standalone command-line tool and as a FastMCP (Fast Model Context Protocol) service for integration with AI assistants like Amazon Q Developer. It translates PowerPoint presentations while preserving formatting and structure.

## Features

- **PowerPoint Translation**: Translate text content in PowerPoint presentations
- **Amazon Bedrock Integration**: Uses Amazon Bedrock models for high-quality translation
- **Format Preservation**: Maintains original formatting, layouts, and styles
- **Language-Specific Fonts**: Automatically applies appropriate fonts for target languages
- **Color & Style Preservation**: Preserves original text colors and formatting even for untranslated content
- **Standalone & MCP Support**: Use as a command-line tool or integrate with AI assistants via FastMCP
- **Multiple Languages**: Supports translation between various languages
- **Batch Processing**: Can handle multiple slides and text elements efficiently
- **Selective Translation**: Translate entire presentations or specific slides

## Examples

### Translation

The PowerPoint Translator maintains the original formatting while accurately translating content:

<table>
<tr>
<td><img src="imgs/original-en-complex.png" alt="English" width="450"/></td>
<td><img src="imgs/translated-ko-complex.png" alt="Korean Example" width="450"/></td>
</tr>
<tr>
<td align="center"><em>Original presentation slide in English <br> with complex layout</em></td>
<td align="center"><em>Same presentation translated to Korean <br> with preserved formatting and layout</em></td>
</tr>
</table>

### Kiro MCP Examples

![kiro1](imgs/kiro-example1.png)

![kiro2](imgs/kiro-example2.png)

![kiro3](imgs/kiro-example3.png)

### Usage Examples

**Translate entire presentation:**
```bash
uv run python server.py --translate --input-file presentation.pptx --target-language ko
```

![standalone](imgs/standalone.png)

**Translate specific slides:**
```bash
uv run python server.py --translate-slides "1,3,5" --input-file presentation.pptx --target-language ja
```

**Get slide information:**
```bash
uv run python server.py --slide-info --input-file presentation.pptx
```

![get-slideinfo](imgs/get-slideinfo.png)

## Prerequisites

- Python 3.11 or higher
- AWS Account with Bedrock access
- AWS CLI configured with appropriate credentials
- Access to Amazon Bedrock models (e.g., Claude, Nova, etc.)

### AWS Credentials Setup

Before using this service, ensure your AWS credentials are properly configured. You have several options:

1. **AWS CLI Configuration (Recommended)**:
   ```bash
   aws configure
   ```
   This will prompt you for:
   - AWS Access Key ID
   - AWS Secret Access Key
   - Default region name
   - Default output format

2. **AWS Profile Configuration**:
   ```bash
   aws configure --profile your-profile-name
   ```

3. **Environment Variables** (if needed):
   ```bash
   export AWS_ACCESS_KEY_ID=your_access_key
   export AWS_SECRET_ACCESS_KEY=your_secret_key
   export AWS_DEFAULT_REGION=us-east-1
   ```

4. **IAM Roles** (when running on EC2 instances)

The service will automatically use your configured AWS credentials. You can specify which profile to use in the `.env` file.

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/daekeun-ml/ppt-translator
   cd ppt-translator
   ```

2. **Install dependencies using uv (recommended)**:
   ```bash
   uv sync
   ```
   
   Or using pip:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up environment variables**:
   ```bash
   cp .env.example .env
   ```
   
   Edit `.env` file with your configuration:
   ```bash
   # AWS Configuration
   AWS_REGION=us-east-1
   AWS_PROFILE=default
   
   # Translation Configuration
   DEFAULT_TARGET_LANGUAGE=ko
   BEDROCK_MODEL_ID=us.anthropic.claude-3-7-sonnet-20250219-v1:0
   
   # Translation Settings
   MAX_TOKENS=4000
   TEMPERATURE=0.1
   ENABLE_POLISHING=true
   BATCH_SIZE=20
   CONTEXT_THRESHOLD=5
   
   # Font Settings by Language
   FONT_KOREAN=맑은 고딕
   FONT_JAPANESE=Yu Gothic UI
   FONT_ENGLISH=Amazon Ember
   FONT_CHINESE=Microsoft YaHei
   FONT_DEFAULT=Arial
   
   # Debug Settings
   DEBUG=false

   # Post-processing Settings
   ENABLE_TEXT_AUTOFIT=true
   TEXT_LENGTH_THRESHOLD=10
   ```

   **Note**: AWS credentials (Access Key ID and Secret Access Key) are not needed in the `.env` file if you have already configured them using `aws configure`. The service will automatically use your AWS CLI credentials.

## Usage

### Standalone Command-Line Usage

The PowerPoint Translator can be used directly from the command line:

```bash
# Translate entire presentation to Korean
uv run python server.py --translate --input-file orig.pptx --target-language ko

# Translate specific slides (individual slides)
uv run python server.py --translate-slides "1,3,5" --input-file orig.pptx --target-language ko

# Translate slide range
uv run python server.py --translate-slides "2-4" --input-file orig.pptx --target-language ko

# Translate mixed (individual + range)
uv run python server.py --translate-slides "1,3-5,8" --input-file orig.pptx --target-language ko

# Get slide information and previews
uv run python server.py --slide-info --input-file orig.pptx

# Using python directly
python server.py --translate --input-file orig.pptx --target-language ko
python server.py --translate-slides "1,3" --input-file orig.pptx --target-language ko
```

### FastMCP Server Mode (for AI Assistant Integration)

Start the FastMCP server for integration with AI assistants like Amazon Q Developer:

```bash
# Using uv (recommended)
uv run python server.py --mcp

# Using python directly
python server.py --mcp
```

## FastMCP Setup (Amazon Q Developer and Kiro)

If you haven't already installed Amazon Q Developer or Kiro, please refer to this:

- Amazon Q Developer CLI: https://docs.aws.amazon.com/amazonq/latest/qdeveloper-ug/command-line-installing.html
- Kiro: https://kiro.dev

### 2. Configure FastMCP Server

Create or update your Q Developer FastMCP configuration file:

#### Kiro
**User Level** `~/.kiro/settings/mcp.json`

#### Amazon Q Developer
**On macOS/Linux**: `~/.aws/amazonq/mcp.json`
**On Windows**: `%APPDATA%\amazonq\mcp.json`

Add the PowerPoint Translator FastMCP server configuration:

```json
{
  "mcpServers": {
    "ppt-translator": {
      "command": "python",
      "args": ["/path/to/ppt-translator/fastmcp_server.py"],
      "env": {
        "AWS_REGION": "us-east-1",
        "AWS_PROFILE": "default",
        "BEDROCK_MODEL_ID": "us.anthropic.claude-3-7-sonnet-20250219-v1:0"
      },
      "disabled": false,
      "autoApprove": [
        "translate_powerpoint",
        "get_slide_info",
        "get_slide_preview",
        "translate_specific_slides"
      ]    
    }
  }
}
```

**Important**: Replace `/path/to/ppt-translator/` with the actual path to your cloned repository.

### 3. Verify FastMCP Server

Test that the FastMCP server is working:

```bash
# Navigate to your project directory
cd /path/to/ppt-translator

# Test the FastMCP server using uv
uv run python server.py --mcp --test

# Or test using python directly
python server.py --mcp --test
```


### 4. Use PowerPoint Translation

Once connected, you can use commands like (User input does not have to be in English):

```
Translate original.pptx to Korean
```

```
Please translate slides 10 to 13 of original.pptx into Japanese.
```

## Available FastMCP Tools

The FastMCP server provides the following tools:

- **`translate_powerpoint`**: Translate an entire PowerPoint presentation
  - Parameters:
    - `input_file`: Path to the input PowerPoint file (.pptx)
    - `target_language`: Target language code (default: 'ko')
    - `output_file`: Path for the translated output file (optional, auto-generated)
    - `model_id`: Amazon Bedrock model ID (default: Claude 3.7 Sonnet)
    - `enable_polishing`: Enable natural language polishing (default: true)

- **`translate_specific_slides`**: Translate only specific slides in a PowerPoint presentation
  - Parameters:
    - `input_file`: Path to the input PowerPoint file (.pptx)
    - `slide_numbers`: Comma-separated slide numbers to translate (e.g., "1,3,5" or "2-4,7")
    - `target_language`: Target language code (default: 'ko')
    - `output_file`: Path for the translated output file (optional, auto-generated)
    - `model_id`: Amazon Bedrock model ID (default: Claude 3.7 Sonnet)
    - `enable_polishing`: Enable natural language polishing (default: true)

- **`get_slide_info`**: Get information about slides in a PowerPoint presentation
  - Parameters:
    - `input_file`: Path to the PowerPoint file (.pptx)
  - Returns: Overview with slide count and preview of each slide's content

- **`get_slide_preview`**: Get detailed preview of a specific slide's content
  - Parameters:
    - `input_file`: Path to the PowerPoint file (.pptx)
    - `slide_number`: Slide number to preview (1-based indexing)

- **`list_supported_languages`**: List all supported target languages for translation

- **`list_supported_models`**: List all supported Amazon Bedrock models

- **`get_translation_help`**: Get help information about using the translator

## Configuration

### Environment Variables

- `AWS_REGION`: AWS region for Bedrock service (default: us-east-1)
- `AWS_PROFILE`: AWS profile to use (default: default)
- `DEFAULT_TARGET_LANGUAGE`: Default target language for translation (default: ko)
- `BEDROCK_MODEL_ID`: Bedrock model ID for translation (default: us.anthropic.claude-3-7-sonnet-20250219-v1:0)
- `MAX_TOKENS`: Maximum tokens for translation requests (default: 4000)
- `TEMPERATURE`: Temperature setting for AI model (default: 0.1)
- `ENABLE_POLISHING`: Enable translation polishing (default: true)
- `BATCH_SIZE`: Number of texts to process in a batch (default: 20)
- `CONTEXT_THRESHOLD`: Number of texts to trigger context-aware translation (default: 5)
- `DEBUG`: Enable debug logging (default: false)

### Supported Languages

The service supports translation between major languages including:
- English (en)
- Korean (ko)
- Japanese (ja)
- Chinese Simplified (zh)
- Chinese Traditional (zh-tw)
- Spanish (es)
- French (fr)
- German (de)
- Italian (it)
- Portuguese (pt)
- Russian (ru)
- Arabic (ar)
- Hindi (hi)
- And many more...

## Troubleshooting

### Common Issues

1. **AWS Credentials Not Found**:
   - Ensure AWS credentials are properly configured
   - Check AWS CLI configuration: `aws configure list`

2. **Bedrock Access Denied**:
   - Verify your AWS account has access to Bedrock
   - Check if the specified model is available in your region

3. **FastMCP Connection Issues**:
   - Verify the path in mcp.json is correct
   - Check that Python and dependencies are properly installed
   - Review logs in Q Developer for error messages
   - Test the server: `uv run python server.py --mcp --test`

4. **PowerPoint File Issues**:
   - Ensure the input file is a valid PowerPoint (.pptx) file
   - Check file permissions for both input and output paths

5. **Module Import Errors**:
   - Use `uv run` to ensure proper virtual environment activation
   - Install dependencies: `uv sync`

## Development

### Project Structure

```
ppt-translator/
├── server.py              # Main server entry point (standalone & MCP)
├── fastmcp_server.py      # FastMCP server implementation
├── ppt_handler.py         # PowerPoint processing logic
├── translation_engine.py  # Translation service
├── bedrock_client.py      # Amazon Bedrock client
├── post_processing.py     # Post-processing utilities for translations
├── config.py             # Configuration management
├── dependencies.py       # Dependency management
├── text_utils.py         # Text processing utilities
├── prompts.py            # Translation prompts
├── requirements.txt      # Python dependencies
├── pyproject.toml        # Project configuration (uv)
└── imgs/                 # Example images and screenshots
```

### Using uv for Development

This project uses `uv` for dependency management:

```bash
# Install dependencies
uv sync

# Run the server
uv run python server.py --mcp

# Run tests
uv run python server.py --mcp --test

# Add new dependencies
uv add package-name

# Update dependencies
uv sync --upgrade
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
