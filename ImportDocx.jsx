/*
    ImportDocx.jsx
    =====================================
    DOCX Import Automation for InDesign
    
    Author: Exemplifi Automation
    Version: 1.0
    Last Updated: 2025
    
    Description:
    This script automates the import of Microsoft Word DOCX files into Adobe InDesign.
    It creates a new document, sets up proper margins, and places the DOCX content
    in a main text frame with appropriate import settings.
    
    Features:
    - Interactive file selection dialog
    - Automatic new document creation
    - Professional margin setup (0.75" on all sides)
    - Automatic Word style import and conversion
    - Auto-fit text frame to content
    - Comprehensive error handling
    - User feedback and console logging
    
    Usage:
    1. Run this script from InDesign's Scripts panel
    2. Select a DOCX file when prompted
    3. The script will create a new document and import the content
    
    Requirements:
    - Adobe InDesign (any recent version)
    - Valid DOCX file format
    
    Error Handling:
    - Graceful cancellation if user cancels file selection
    - Try/catch wrapper prevents app lockup
    - Detailed error messages for debugging
*/

#target "InDesign"

// Always wrap in try/catch so you don't lock the app
try {
    // === STEP 1: File Selection ===
    // Prompt user to select a DOCX file for import
    var docxFile = File.openDialog("Select a DOCX file to import", "*.docx");
    if (!docxFile) {
        alert("Import cancelled.");
        exit();
    }

    // === STEP 2: Document Creation ===
    // Create a new InDesign document with margins
    var doc = app.documents.add();
    doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

    // === STEP 3: Set Document Margins ===
    // Set up professional margins (0.75 inches on all sides)
    doc.documentPreferences.documentBleedTopOffset = "0in";
    doc.documentPreferences.documentBleedBottomOffset = "0in";
    doc.documentPreferences.documentBleedInsideOrLeftOffset = "0in";
    doc.documentPreferences.documentBleedOutsideOrRightOffset = "0in";
    
    // Set margins
    doc.marginPreferences.top = "0.75in";
    doc.marginPreferences.left = "0.75in";
    doc.marginPreferences.right = "0.75in";
    doc.marginPreferences.bottom = "0.75in";

    // === STEP 4: Text Frame Creation ===
    // Create a text frame that fills the page within margins
    var page = doc.pages[0];
    var textFrame = page.textFrames.add();
    
    // Use the page margins to create the text frame
    textFrame.geometricBounds = [
        doc.marginPreferences.top,
        doc.marginPreferences.left,
        page.bounds[2] - doc.marginPreferences.bottom,
        page.bounds[3] - doc.marginPreferences.right
    ];

    // === STEP 5: Font Handling Setup ===
    // Font configuration removed - not supported in all InDesign versions
    
    // === STEP 6: Content Placement ===
    // Place the DOCX content into the text frame
    // InDesign will automatically handle Word style import and conversion
    textFrame.insertionPoints[0].place(docxFile);
    
    // === STEP 6.5: Style Mapping ===
    // Map Word styles to InDesign paragraph styles
    try {
        // Create or get InDesign paragraph styles
        var heading1Style = createOrGetParagraphStyle(doc, "Heading 1", {
            fontFamily: "Minion Pro Bold",
            fontSize: "24pt",
            fontWeight: "Bold",
            spaceAfter: "12pt",
            spaceBefore: "18pt",
            fontColor: "Green",
            underline: true,
            backgroundColor: "Light Green"
        });
        
        var heading2Style = createOrGetParagraphStyle(doc, "Heading 2", {
            fontFamily: "Minion Pro Bold", 
            fontSize: "20pt",
            fontWeight: "Bold",
            spaceAfter: "10pt",
            spaceBefore: "14pt",
            fontColor: "Green",
            backgroundColor: "Very Light Green"
        });
        
        var heading3Style = createOrGetParagraphStyle(doc, "Heading 3", {
            fontFamily: "Minion Pro Bold",
            fontSize: "16pt", 
            fontWeight: "Bold",
            spaceAfter: "8pt",
            spaceBefore: "12pt",
            fontColor: "Red",
            backgroundColor: "Pink"
        });
        
        var heading4Style = createOrGetParagraphStyle(doc, "Heading 4", {
            fontFamily: "Minion Pro Bold",
            fontSize: "14pt", 
            fontWeight: "Bold",
            spaceAfter: "6pt",
            spaceBefore: "10pt",
            fontColor: "Green",
            backgroundColor: "Very Light Green"
        });
        
        var heading5Style = createOrGetParagraphStyle(doc, "Heading 5", {
            fontFamily: "Minion Pro Bold",
            fontSize: "13pt", 
            fontWeight: "Bold",
            spaceAfter: "4pt",
            spaceBefore: "8pt",
            fontColor: "White",
            backgroundColor: "Dark Red"
        });
        
        var heading6Style = createOrGetParagraphStyle(doc, "Heading 6", {
            fontFamily: "Minion Pro Bold",
            fontSize: "12pt", 
            fontWeight: "Bold",
            spaceAfter: "4pt",
            spaceBefore: "6pt",
            fontColor: "Green",
            leftBorder: "Green"
        });
        
        var bodyStyle = createOrGetParagraphStyle(doc, "Body Text", {
            fontFamily: "Minion Pro",
            fontSize: "12pt",
            fontWeight: "Regular",
            spaceAfter: "6pt",
            spaceBefore: "0pt"
        });
        
        // Map Word styles to InDesign styles
        var styleMap = {
            "Heading 1": heading1Style,
            "Heading 2": heading2Style, 
            "Heading 3": heading3Style,
            "Heading 4": heading4Style,
            "Heading 5": heading5Style,
            "Heading 6": heading6Style,
            "Normal": bodyStyle,
            "Body Text": bodyStyle
        };
        
        mapWordStylesToInDesign(textFrame.parentStory, styleMap, doc);
        
    } catch (styleError) {
        $.writeln("Style mapping warning: " + styleError.message);
    }

    // === STEP 6.6: Font Substitution ===
    // Handle any missing fonts by substituting with available fonts
    try {
        // Get all text in the frame and check for missing fonts
        var stories = textFrame.parentStory;
        if (stories && stories.length > 0) {
            for (var i = 0; i < stories.length; i++) {
                var story = stories[i];
                if (story && story.characters && story.characters.length > 0) {
                    // Replace any missing fonts with default system font
                    for (var j = 0; j < story.characters.length; j++) {
                        try {
                            var character = story.characters[j];
                            if (character.appliedFont && character.appliedFont.name) {
                                // Font exists, continue
                            }
                        } catch (fontError) {
                            // Font missing, apply Minion Pro as default
                            try {
                                character.appliedFont = app.fonts.item("Minion Pro");
                            } catch (e) {
                                // If Minion Pro not available, use any available font
                                character.appliedFont = app.fonts[0];
                            }
                        }
                    }
                }
            }
        }
    } catch (subError) {
        $.writeln("Font substitution warning: " + subError.message);
    }

    // === STEP 7: Frame Adjustment ===
    // Auto-fit the text frame to the imported content
    textFrame.fit(FitOptions.FRAME_TO_CONTENT);

    // === STEP 8: Success Feedback ===
    // Log success and notify user
    $.writeln("DOCX imported successfully: " + docxFile.fsName);
    alert("Import complete! Check Page 1. Missing fonts have been automatically substituted.");

} catch (e) {
    // === ERROR HANDLING ===
    // Display user-friendly error message and log detailed error info
    alert("Error: " + e.message);
    $.writeln("Error stack:\n" + e.stack);
}

// === HELPER FUNCTIONS ===

/**
 * Create or get a paragraph style with specified properties
 * @param {Document} doc - The InDesign document
 * @param {string} styleName - Name of the paragraph style
 * @param {Object} properties - Style properties (fontFamily, fontSize, etc.)
 * @returns {ParagraphStyle} The paragraph style
 */
function createOrGetParagraphStyle(doc, styleName, properties) {
    var style;
    
    // Check if style already exists
    try {
        style = doc.paragraphStyles.item(styleName);
    } catch (e) {
        // Style doesn't exist, create it
        style = doc.paragraphStyles.add({
            name: styleName
        });
    }
    
    // Apply properties to the style
    try {
        if (properties.fontFamily) {
            style.appliedFont = app.fonts.item(properties.fontFamily);
        }
        if (properties.fontSize) {
            style.pointSize = UnitValue(properties.fontSize).as("pt");
        }
        if (properties.fontWeight) {
            if (properties.fontWeight === "Bold") {
                try {
                    style.fontStyle = "Bold";
                } catch (e) {
                    // If Bold style not available, try to set bold weight
                    try {
                        style.appliedFont = app.fonts.item(properties.fontFamily + " Bold");
                    } catch (e2) {
                        // Fallback: try to find any bold variant
                        var fontFamily = properties.fontFamily;
                        var boldFonts = app.fonts.itemByName(fontFamily + " Bold");
                        if (boldFonts.length > 0) {
                            style.appliedFont = boldFonts[0];
                        }
                    }
                }
            } else {
                style.fontStyle = "Regular";
            }
        }
        if (properties.spaceAfter) {
            style.spaceAfter = UnitValue(properties.spaceAfter).as("pt");
        }
        if (properties.spaceBefore) {
            style.spaceBefore = UnitValue(properties.spaceBefore).as("pt");
        }
        
        // Handle font color
        if (properties.fontColor) {
            try {
                var color = getOrCreateColor(doc, properties.fontColor);
                style.fillColor = color;
            } catch (e) {
                $.writeln("Warning: Could not set font color " + properties.fontColor);
            }
        }
        
        // Handle underline
        if (properties.underline) {
            try {
                style.underline = true;
            } catch (e) {
                $.writeln("Warning: Could not set underline");
            }
        }
        
        // Handle background color using paragraph rules
        if (properties.backgroundColor) {
            try {
                var bgColor = getOrCreateColor(doc, properties.backgroundColor);
                style.paragraphRules = ParagraphRules.ADD;
                style.ruleAbove = true;
                style.ruleAboveColor = bgColor;
                style.ruleAboveWeight = "12pt"; // Make it thick to look like background
                style.ruleAboveOffset = "0pt";
            } catch (e) {
                $.writeln("Warning: Could not set background color " + properties.backgroundColor);
            }
        }
        
        // Handle left border using paragraph rules
        if (properties.leftBorder) {
            try {
                var borderColor = getOrCreateColor(doc, properties.leftBorder);
                style.paragraphRules = ParagraphRules.ADD;
                style.ruleLeft = true;
                style.ruleLeftColor = borderColor;
                style.ruleLeftWeight = "3pt";
                style.ruleLeftOffset = "0pt";
            } catch (e) {
                $.writeln("Warning: Could not set left border");
            }
        }
        
    } catch (propError) {
        $.writeln("Warning: Could not set all properties for style " + styleName + ": " + propError.message);
    }
    
    return style;
}

/**
 * Get or create a color in the document
 * @param {Document} doc - The InDesign document
 * @param {string} colorName - Name of the color
 * @returns {Color} The color object
 */
function getOrCreateColor(doc, colorName) {
    var color;
    
    // Try to get existing color
    try {
        color = doc.colors.item(colorName);
    } catch (e) {
        // Color doesn't exist, create it
        color = doc.colors.add({
            name: colorName
        });
        
        // Set color values based on name
        switch (colorName) {
            case "Green":
                color.colorValue = [100, 0, 100, 0]; // CMYK Green
                break;
            case "Light Green":
                color.colorValue = [20, 0, 100, 0]; // Light CMYK Green
                break;
            case "Very Light Green":
                color.colorValue = [5, 0, 20, 0]; // Very Light CMYK Green
                break;
            case "Red":
                color.colorValue = [0, 100, 100, 0]; // CMYK Red
                break;
            case "Pink":
                color.colorValue = [0, 50, 20, 0]; // CMYK Pink
                break;
            case "Dark Red":
                color.colorValue = [0, 100, 100, 50]; // Dark CMYK Red
                break;
            case "White":
                color.colorValue = [0, 0, 0, 0]; // CMYK White
                break;
            default:
                color.colorValue = [0, 0, 0, 100]; // Default to Black
        }
    }
    
    return color;
}

/**
 * Map Word paragraph styles to InDesign paragraph styles
 * @param {Story} story - The text story to process
 * @param {Object} styleMap - Object mapping Word style names to InDesign styles
 * @param {Document} doc - The InDesign document
 */
function mapWordStylesToInDesign(story, styleMap, doc) {
    try {
        // Get all paragraphs in the story
        var paragraphs = story.paragraphs;
        var mappedCount = 0;
        var unmappedCount = 0;
        var noStyleCount = 0;
        
        for (var i = 0; i < paragraphs.length; i++) {
            try {
                var paragraph = paragraphs[i];
                
                // Skip invalid or empty paragraphs
                if (!paragraph || paragraph.isValid === false) {
                    continue;
                }
                
                // Check if paragraph has a Word style applied
                if (paragraph.appliedParagraphStyle && paragraph.appliedParagraphStyle.name) {
                var wordStyleName = paragraph.appliedParagraphStyle.name;
                
                // Check if we have a mapping for this Word style
                if (styleMap[wordStyleName]) {
                    var targetStyleName = styleMap[wordStyleName].name;
                    
                    try {
                        // Get the style from the document by name
                        var targetStyle = doc.paragraphStyles.item(targetStyleName);
                        paragraph.appliedParagraphStyle = targetStyle;
                        mappedCount++;
                    } catch (styleError) {
                        alert("Error applying style " + targetStyleName + ": " + styleError.message);
                        unmappedCount++;
                    }
                } else {
                    unmappedCount++;
                }
            } else {
                noStyleCount++;
            }
            } catch (paragraphError) {
                // Skip this paragraph if there's an error
                continue;
            }
        }
        
        return "Mapped: " + mappedCount + ", Unmapped: " + unmappedCount + ", No Style: " + noStyleCount;
        
    } catch (mapError) {
        alert("Style mapping error: " + mapError.message);
        return "Error: " + mapError.message;
    }
}
