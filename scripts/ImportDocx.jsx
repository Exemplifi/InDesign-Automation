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
    - Word style import and mapping
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
    // Create a new InDesign document with default settings
    var doc = app.documents.add();
    doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

    // === STEP 3: Layout Configuration ===
    // Set up professional margins (0.75 inches on all sides)
    // These can be customized based on your document requirements
    var marginTop = "0.75in";
    var marginLeft = "0.75in";
    var marginRight = "0.75in";
    var marginBottom = "0.75in";

    // === STEP 4: Text Frame Creation ===
    // Create a main text frame on the first page with proper margins
    var page = doc.pages[0];
    var bounds = page.bounds; // [y1, x1, y2, x2] - page boundaries
    
    // Calculate text frame bounds respecting margins
    var textFrameBounds = [
        bounds[0] + UnitValue(marginTop).as("pt"),    // Top margin
        bounds[1] + UnitValue(marginLeft).as("pt"),   // Left margin
        bounds[2] - UnitValue(marginBottom).as("pt"), // Bottom margin
        bounds[3] - UnitValue(marginRight).as("pt")   // Right margin
    ];

    // Add the text frame to the page
    var textFrame = page.textFrames.add({
        geometricBounds: textFrameBounds
    });

    // === STEP 5: Import Configuration ===
    // Configure Word import preferences for optimal results
    var importOptions = app.wordImportPreferences;
    importOptions.preserveLocalOverrides = false;  // Don't preserve Word formatting overrides
    importOptions.importUnusedStyles = false;      // Don't import unused Word styles
    importOptions.importWordStyles = true;         // Import Word paragraph styles
    importOptions.convertPageBreaks = true;        // Convert Word page breaks to InDesign

    // === STEP 6: Content Placement ===
    // Place the DOCX content into the text frame
    textFrame.place(docxFile);

    // === STEP 7: Frame Adjustment ===
    // Auto-fit the text frame to the imported content
    textFrame.fit(FitOptions.FRAME_TO_CONTENT);

    // === STEP 8: Success Feedback ===
    // Log success and notify user
    $.writeln("DOCX imported successfully: " + docxFile.fsName);
    alert("Import complete! Check Page 1.");

} catch (e) {
    // === ERROR HANDLING ===
    // Display user-friendly error message and log detailed error info
    alert("Error: " + e.message);
    $.writeln("Error stack:\n" + e.stack);
}
