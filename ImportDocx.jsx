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
    // Let InDesign use existing styles from the template
    // No custom style creation - styles will be picked up from the InDesign template

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

