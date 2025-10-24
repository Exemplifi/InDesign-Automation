# Exemplifi SmartDoc InDesign Import

Automated Word/RTF import script for InDesign with intelligent formatting, style mapping, and overflow handling.

## Features

### 🎯 **Smart Document Import**
- Imports Word (.docx) and RTF files with preserved formatting
- Automatically handles text overflow by creating new pages
- Creates text frames inside page margins for proper layout

### 🎨 **Intelligent Style Mapping**
- Maps Word styles to InDesign paragraph styles:
  - `Normal` → `Body`
  - `Heading 1-6` → `Heading 1-6`
  - Bullet points → `Bullets` and `Sub-Bullets`
- Preserves bold/italic formatting from RTF files
- Applies table styles automatically

### 🧹 **Content Optimization**
- Collapses multiple paragraph returns (^p^p → ^p)
- Trims leading/trailing spaces
- Removes empty paragraphs
- Clears style overrides for clean formatting

### 📄 **Advanced Layout Handling**
- Detects and handles text overflow
- Creates new pages automatically when needed
- Maintains proper margins and text frame positioning
- Supports bullet points (●) and sub-bullets (○)

## Usage

### Prerequisites
1. Open an InDesign document with the following paragraph styles:
   - `Body`
   - `Heading 1`, `Heading 2`, `Heading 3`, `Heading 4`, `Heading 5`, `Heading 6`
   - `Bullets`
   - `Sub-Bullets`
   - `Table Style 1` (for tables)

### Running the Script
1. Open InDesign and create/open a document
2. Go to **Window ▸ Utilities ▸ Scripts ▸ Scripts Panel**
3. Double-click `Exemplifi SmartDoc InDesign Import.jsx`
4. Select your Word (.docx) or RTF file
5. The script will automatically:
   - Import the content
   - Apply appropriate styles
   - Handle overflow with new pages
   - Clean up formatting

## Installation

1. Clone this repository to your InDesign Scripts Panel directory:
   ```bash
   git clone https://github.com/Exemplifi/InDesign-Automation.git
   ```

2. The script will appear in your Scripts Panel automatically

## Technical Details

### Style Mapping Logic
- **Word "Normal" style** → InDesign "Body" style
- **Word "Heading 1-6"** → InDesign "Heading 1-6" styles
- **Bullet points (●)** → "Bullets" style
- **Sub-bullets (○)** → "Sub-Bullets" style
- **Tables** → "Table Style 1"

### Content Processing
- Removes RTF bullet characters and tabs
- Preserves inline formatting (bold/italic)
- Optimizes paragraph spacing
- Handles empty content gracefully

### Error Handling
- Graceful fallback for missing styles
- Preserves original interaction settings
- Comprehensive error reporting

## Requirements

- Adobe InDesign (tested with version 20.0+)
- Word (.docx) or RTF files with standard formatting
- InDesign document with required paragraph styles

## Support

For issues or feature requests, please contact the Exemplifi development team.
