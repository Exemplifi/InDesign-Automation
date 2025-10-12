#target "InDesign"

var doc = app.activeDocument; // or app.documents.add()
var page = doc.pages[0];

// Example: Add a text frame
var frame = page.textFrames.add({
    geometricBounds: ["1in", "1in", "3in", "5in"], // [y1,x1,y2,x2]
    contents: "Hello world"
});

// Example: Apply a style (if it exists)
try {
    frame.texts[0].appliedParagraphStyle = doc.paragraphStyles.itemByName("Body");
} catch(e) {
    $.writeln("Style not found");
}
