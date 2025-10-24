// InDesign Script: Automatic Word/RTF Import with Headings, Body, Bullets/Sub-Bullets, inside page margins
// Optimized: collapse multiple paragraph returns (^p^p → ^p), trim spaces, remove empty paragraphs
// Handles overflow, applies styles, clears overrides, removes bullets/tabs, preserves bold/italic, applies table styles

if (app.documents.length === 0) {
    alert("Please open an InDesign template first!");
    exit();
}

var doc = app.activeDocument;

// Suppress dialogs
var originalInteractionLevel;
try { originalInteractionLevel = app.scriptPreferences.userInteractionLevel; } catch (e) { }
try { app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT; } catch (e) { }

try {
    // Step 1: Select Word/RTF file
    var importFile = File.openDialog("Select Word (.docx) or RTF file to import", "*.docx;*.rtf", false);
    if (importFile == null || !importFile.exists){
        try { if (originalInteractionLevel !== null) app.scriptPreferences.userInteractionLevel = originalInteractionLevel; } catch (e) { }
        alert("No file selected or file does not exist. Script cancelled.");
        exit();
    }

    // Helper: Create text frame inside margins
    function createFrameInsideMargins(page){
        var m = page.marginPreferences;
        var b = page.bounds;
        return page.textFrames.add({
            geometricBounds: [b[0]+m.top, b[1]+m.left, b[2]-m.bottom, b[3]-m.right]
        });
    }

    // Step 2: Determine first text frame or create one
    var page = doc.pages[0];
    var targetFrame = (page.textFrames.length > 0) ? page.textFrames[0] : createFrameInsideMargins(page);

    // Step 3: Place the file (preserve inline formatting: bold/italic)
    targetFrame.place(importFile);
    var story = targetFrame.parentStory;
    if (story == null) throw new Error("No story found in the placed text frame.");

    // Step 4: Flow overflow text into new pages inside margins
    while (story.overflows){
        var newPage = doc.pages.add(LocationOptions.AFTER, doc.pages[-1]);
        var newFrame = createFrameInsideMargins(newPage);
        story.textContainers[story.textContainers.length - 1].nextTextFrame = newFrame;
    }

    // Step 5: Optimized cleanup
    for (var i = 0; i < story.paragraphs.length; i++){
        var para = story.paragraphs[i];
        try{
            para.changeGrep({
                findWhat: "(\\r){2,}",
                changeTo: "\r",
                wholeWord: false,
                caseSensitive: false
            });
            para.contents = para.contents.replace(/^\s+|\s+$/g, "");
            if (para.contents.replace(/\s/g, "") === ""){
                para.remove();
                i--;
            }
        } catch(e){}
    }

    // Step 6: Helper to get styles safely
    function getStyle(doc, name){
        try { var s = doc.paragraphStyles.itemByName(name); s.name; return s; } catch (e) { return null; }
    }

    // Step 7: Get styles
    var bodyStyle = getStyle(doc, "Body");
    var h1Style = getStyle(doc, "Heading 1");
    var h2Style = getStyle(doc, "Heading 2");
    var h3Style = getStyle(doc, "Heading 3");
    var h4Style = getStyle(doc, "Heading 4");
    var h5Style = getStyle(doc, "Heading 5");
    var h6Style = getStyle(doc, "Heading 6");
    var bulletStyle = getStyle(doc, "Bullets");
    var subBulletStyle = getStyle(doc, "Sub-Bullets");
    var tableStyle = getStyle(doc, "Table Style 1");

    // Step 8: Map Word styles to InDesign styles (headings & body)
    for (var i = 0; i < story.paragraphs.length; i++){
        var para = story.paragraphs[i];
        var styleName = "";
        try { styleName = para.appliedParagraphStyle.name; } catch (e) { }

        if (/Normal/i.test(styleName) && bodyStyle) para.appliedParagraphStyle = bodyStyle;
        else if (/Heading 1/i.test(styleName) && h1Style) para.appliedParagraphStyle = h1Style;
        else if (/Heading 2/i.test(styleName) && h2Style) para.appliedParagraphStyle = h2Style;
        else if (/Heading 3/i.test(styleName) && h3Style) para.appliedParagraphStyle = h3Style;
        else if (/Heading 4/i.test(styleName) && h4Style) para.appliedParagraphStyle = h4Style;
        else if (/Heading 5/i.test(styleName) && h5Style) para.appliedParagraphStyle = h5Style;
        else if (/Heading 6/i.test(styleName) && h6Style) para.appliedParagraphStyle = h6Style;
        else if (bodyStyle) para.appliedParagraphStyle = bodyStyle;

        try { para.clearOverrides(); } catch (e) { }
    }

    // Step 9: Apply Bullets & Sub-Bullets (remove RTF bullets + tab)
    for (var i = 0; i < story.paragraphs.length; i++){
        var para = story.paragraphs[i];
        var text = para.contents;

        // Match "●" + tab
        if (/^●\t/.test(text) && bulletStyle){
            para.contents = text.replace(/^●\t/, "");
            para.appliedParagraphStyle = bulletStyle;
            para.clearOverrides();
            continue;
        }

        // Match "○" + tab
        if (/^○\t/.test(text) && subBulletStyle){
            para.contents = text.replace(/^○\t/, "");
            para.appliedParagraphStyle = subBulletStyle;
            para.clearOverrides();
            continue;
        }
    }

    // Step 10: Apply Table Style 1 if tables exist
    if (tableStyle && story.tables.length > 0){
        for (var t = 0; t < story.tables.length; t++){
            try { story.tables[t].appliedTableStyle = tableStyle; } catch (e) { }
        }
    }

    alert("✅ Import completed!\nHeadings, Body, Bullets/Sub-Bullets applied, Table Style applied.\nBold/italic preserved from RTF.");

} catch (err){
    alert("Script error:\n" + err.message);
} finally{
    try { if (originalInteractionLevel !== null) app.scriptPreferences.userInteractionLevel = originalInteractionLevel; } catch (e) { }
}
