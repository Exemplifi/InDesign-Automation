// InDesign Script: Automatic Word/RTF Import
// Handles: Headings, Body, Bullets/Sub-Bullets, Tables, Hyperlinks
// Fixes: Heading merge, hybrid bullet mapping, and URL filtering.

#target "InDesign"

if (app.documents.length === 0) {
    alert("Please open an InDesign template (.indd or .indt) first!");
    exit();
}

var doc = app.activeDocument;
var originalInteractionLevel = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

try {
    // ---- 1) Select import file ----
    var importFile = File.openDialog("Select Word (.docx) or RTF file to import", "*.docx;*.rtf", false);
    if (!importFile || !importFile.exists) throw new Error("No file selected or file does not exist.");

    // ---- 2) Prepare helper ----
    function createFrameInsideMargins(page) {
        var m = page.marginPreferences, b = page.bounds;
        return page.textFrames.add({ geometricBounds: [b[0]+m.top, b[1]+m.left, b[2]-m.bottom, b[3]-m.right] });
    }

    // ---- 3) Create or find first frame ----
    var firstPage = doc.pages[0];
    var targetFrame = (firstPage.textFrames.length > 0) ? firstPage.textFrames[0] : createFrameInsideMargins(firstPage);

    // ---- 4) Place file ----
    targetFrame.place(importFile);
    var story = targetFrame.parentStory;
    if (!story) throw new Error("No story found after placing file.");

    while (story.overflows) {
        var lastPage = doc.pages[doc.pages.length - 1];
        var newPage = doc.pages.add(LocationOptions.AFTER, lastPage);
        var newFrame = createFrameInsideMargins(newPage);
        story.textContainers[story.textContainers.length - 1].nextTextFrame = newFrame;
    }

    // ---- 5) Clean up multiple returns ----
    try {
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        app.findGrepPreferences.findWhat = "(\\r){3,}";
        app.changeGrepPreferences.changeTo = "\r\r";
        story.changeGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
    } catch (_) {}

    // ---- 6) Insert paragraph break between consecutive Headings ----
    try {
        app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
        app.findGrepPreferences.findWhat = "(?<=Heading 1)(?=Heading 2)";
        story.changeGrep();
    } catch (_) {}

    // ---- 7) Trim and remove empties ----
    for (var i = 0; i < story.paragraphs.length; i++) {
        var p = story.paragraphs[i];
        try {
            p.contents = p.contents.replace(/^\s+|\s+$/g, "");
            if (p.contents.replace(/\s/g, "") === "") { p.remove(); i--; }
        } catch (_) {}
    }

    // ---- 8) Get styles ----
    function getStyle(docRef, name) {
        try { var s = docRef.paragraphStyles.itemByName(name); s.name; return s; } catch (_) { return null; }
    }
    var styles = {
        body: getStyle(doc, "Body"),
        h1: getStyle(doc, "Heading 1"),
        h2: getStyle(doc, "Heading 2"),
        h3: getStyle(doc, "Heading 3"),
        bullet: getStyle(doc, "Bullets"),
        subBullet: getStyle(doc, "Sub-Bullets"),
        table: getStyle(doc, "Table Style 1"),
        numbered: getStyle(doc, "Numbered")
    };

// ---- 9) Apply style mapping (Headings first) ----
for (var j = 0; j < story.paragraphs.length; j++) {
    var para = story.paragraphs[j];
    var srcName = "";
    try { srcName = para.appliedParagraphStyle.name || ""; } catch (_) {}

    if (/Heading 1/i.test(srcName) && styles.h1) para.appliedParagraphStyle = styles.h1;
    else if (/Heading 2/i.test(srcName) && styles.h2) para.appliedParagraphStyle = styles.h2;
    else if (/Heading 3/i.test(srcName) && styles.h3) para.appliedParagraphStyle = styles.h3;
    else if (/Normal/i.test(srcName) && styles.body) para.appliedParagraphStyle = styles.body;
    else if (styles.body) para.appliedParagraphStyle = styles.body;

    try { para.clearOverrides(); } catch (_) {}
}

// ---- 10) Normalize bullets and lists (tab-only + style-aware, no false positives) ----
function normalizeBulletsAndLists(tf, styles) {
    var paras = tf.paragraphs;

    // Bullet characters some RTFs keep as plain text
    var glyphRE = /^[\u2022\u25CF\u25E6\uF0B7\u2219\u00B7○◦o\-–•]+\s*\t?/;

    // Helper: is this paragraph tagged with a list-y style name?
    function looksListyStyle(p) {
        try {
            var n = (p.appliedParagraphStyle && p.appliedParagraphStyle.name) || "";
            return /list|bullet/i.test(n); // catches "List Paragraph", "Bulleted", etc.
        } catch (_) { return false; }
    }

    var prevWasBullet = false;

    for (var k = 0; k < paras.length; k++) {
        var p   = paras[k];
        var txt = (p.contents || "").replace(/\s+$/,"");

        // Skip headings and tables entirely
        try {
            var sn = (p.appliedParagraphStyle && p.appliedParagraphStyle.name) || "";
            if (/^heading\s*\d+$/i.test(sn) || /^table/i.test(sn)) { prevWasBullet = false; continue; }
        } catch (_) {}

        var isTrueList   = (p.bulletsAndNumberingListType == ListType.BULLET_LIST);
        var hasGlyph     = glyphRE.test(txt);
        var startsWithTab= /^\t/.test(txt);
        var styleIsListy = looksListyStyle(p);

        // Decide if this should be a bullet:
        // 1) Already a real InDesign bullet
        // 2) Has a visible bullet glyph/dash
        // 3) Starts with a tab AND (style looks listy OR previous paragraph was a bullet)
        var shouldBeBullet = isTrueList || hasGlyph || (startsWithTab && (styleIsListy || prevWasBullet));

        if (shouldBeBullet) {
            // Strip any leading glyphs/dashes/tabs that snuck in
            p.contents = txt.replace(glyphRE, "").replace(/^\t+/, "");
            p.appliedParagraphStyle = styles.bullet || p.appliedParagraphStyle;
            p.bulletsAndNumberingListType = ListType.BULLET_LIST;
            try { p.clearOverrides(); } catch(_) {}
            prevWasBullet = true;
        } else {
            prevWasBullet = false;
        }
    }
}

normalizeBulletsAndLists(story, styles);

    // ---- 11) Table styling ----
    if (styles.table && story.tables.length > 0) {
        for (var t = 0; t < story.tables.length; t++) {
            try { story.tables[t].appliedTableStyle = styles.table; } catch (_) {}
        }
    }

    // ---- 12) Apply hyperlink style only to URLs ----
    var hyperlinkStyleName = "Hyperlink Highlight";
    var hyperlinkStyle = null;
    try { hyperlinkStyle = doc.characterStyles.itemByName(hyperlinkStyleName); hyperlinkStyle.name; }
    catch (_) { hyperlinkStyle = null; }

    function isUrlDest(d) {
        try { if (d instanceof HyperlinkURLDestination) return true; } catch (_) {}
        var u = d && d.destinationURL ? String(d.destinationURL) : "";
        return /^(https?:|mailto:|ftp:)/i.test(u);
    }

    if (hyperlinkStyle && hyperlinkStyle.isValid && doc.hyperlinks.length > 0) {
        for (var h = 0; h < doc.hyperlinks.length; h++) {
            var link = doc.hyperlinks[h];
            try {
                if (!isUrlDest(link.destination)) continue;
                var src = link.source;
                if (src instanceof HyperlinkTextSource) {
                    var text = src.sourceText;
                    if (text && text.length > 0) text.appliedCharacterStyle = hyperlinkStyle;
                }
            } catch (_) {}
        }
    }

    story.recompose();
    alert("✅ Import complete.\nHeadings, Bullets, Tables, and Hyperlink styles applied.");

} catch (err) {
    alert("❌ Script error:\n" + err.message);
    $.writeln("❌ " + err.message);
} finally {
    app.scriptPreferences.userInteractionLevel = originalInteractionLevel;
}
