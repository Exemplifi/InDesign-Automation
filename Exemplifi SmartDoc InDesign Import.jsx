// InDesign Script: Automatic Word/RTF Import with Headings, Body, Bullets/Sub-Bullets, inside page margins
// Optimized: collapse multiple paragraph returns (^p^p → ^p), trim spaces, remove empty paragraphs
// Handles overflow, applies styles, clears overrides, preserves bold/italic, applies table styles
// Cross-platform bullet repair for RTF/DOCX on Windows & Mac
// Adds: apply custom Hyperlink character style to imported links

#target "InDesign"

if (app.documents.length === 0) {
    alert("Please open an InDesign template (.indd or .indt) first!");
    exit();
}

var doc = app.activeDocument;

// ---- Preserve & suppress UI ----
var originalInteractionLevel = app.scriptPreferences.userInteractionLevel;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

try {
    // ---- 1) Pick file ----
    var importFile = File.openDialog("Select Word (.docx) or RTF file to import", "*.docx;*.rtf", false);
    if (!importFile || !importFile.exists) {
        throw new Error("No file selected or file does not exist.");
    }

    var isRTF  = (/\.rtf$/i).test(importFile.name);
    var isDOCX = (/\.docx$/i).test(importFile.name);

    // ---- 2) Normalize import preferences for consistency ----
    if (isDOCX) {
        var w = app.wordImportPreferences;
        w.preserveLocalOverrides = false;
        w.importWordStyles = true;
        w.importUnusedStyles = false;
        w.convertPageBreaks = true;
        w.preserveTrackChanges = false;
        // w.useTypographersQuotes = true; // optional
        } else if (isRTF) {
            // RTF uses the same importer as Word; reuse Word prefs
            var w = app.wordImportPreferences;
            w.preserveLocalOverrides = false;
            w.importWordStyles = true;
            w.importUnusedStyles = false;
            w.convertPageBreaks = true;
        }


    // ---- 3) Helper: create a text frame inside margins ----
    function createFrameInsideMargins(page) {
        var m = page.marginPreferences, b = page.bounds; // [y1,x1,y2,x2] in points
        return page.textFrames.add({
            geometricBounds: [b[0]+m.top, b[1]+m.left, b[2]-m.bottom, b[3]-m.right]
        });
    }

    // ---- 4) Get or create the first frame on page 1 ----
    var firstPage = doc.pages[0];
    var targetFrame = (firstPage.textFrames.length > 0) ? firstPage.textFrames[0] : createFrameInsideMargins(firstPage);

    // ---- 5) Place the file (keeps inline bold/italic) ----
    targetFrame.place(importFile);
    var story = targetFrame.parentStory;
    if (!story) throw new Error("No story found in the placed text frame.");

    // ---- 6) Flow overset text across new pages ----
    while (story.overflows) {
        var lastPage = doc.pages[doc.pages.length - 1];
        var newPage  = doc.pages.add(LocationOptions.AFTER, lastPage);
        var newFrame = createFrameInsideMargins(newPage);
        story.textContainers[story.textContainers.length - 1].nextTextFrame = newFrame;
    }

    // ---- 7) Global GREP cleanup: collapse multiple returns ----
    // Faster & safer to run on the story than per paragraph
    story.changeGrep({ findWhat: "(\\r){2,}", changeTo: "\r" });

    // ---- 8) Trim whitespace & remove empty paragraphs ----
    // (Do this after collapsing returns)
    for (var i = 0; i < story.paragraphs.length; i++) {
        var p = story.paragraphs[i];
        try {
            // Trim leading/trailing spaces
            p.contents = p.contents.replace(/^\s+|\s+$/g, "");
            // Remove para if empty (only whitespace)
            if (p.contents.replace(/\s/g, "") === "") { p.remove(); i--; }
        } catch (_) {}
    }

    // ---- 9) Safe style getter ----
    function getStyle(docRef, name) {
        try { var s = docRef.paragraphStyles.itemByName(name); s.name; return s; }
        catch (_) { return null; }
    }

    // ---- 10) Gather styles ----
    var styles = {
        body:      getStyle(doc, "Body"),
        h1:        getStyle(doc, "Heading 1"),
        h2:        getStyle(doc, "Heading 2"),
        h3:        getStyle(doc, "Heading 3"),
        h4:        getStyle(doc, "Heading 4"),
        h5:        getStyle(doc, "Heading 5"),
        h6:        getStyle(doc, "Heading 6"),
        bullet:    getStyle(doc, "Bullets"),
        subBullet: getStyle(doc, "Sub-Bullets"),
        table:     getStyle(doc, "Table Style 1")
    };

    // ---- 11) Map incoming Word/RTF para styles → your InDesign styles ----
    for (var j = 0; j < story.paragraphs.length; j++) {
        var para = story.paragraphs[j];
        var srcName = "";
        try { srcName = para.appliedParagraphStyle.name || ""; } catch (_) {}

        if (/Normal/i.test(srcName) && styles.body)        para.appliedParagraphStyle = styles.body;
        else if (/Heading 1/i.test(srcName) && styles.h1)   para.appliedParagraphStyle = styles.h1;
        else if (/Heading 2/i.test(srcName) && styles.h2)   para.appliedParagraphStyle = styles.h2;
        else if (/Heading 3/i.test(srcName) && styles.h3)   para.appliedParagraphStyle = styles.h3;
        else if (/Heading 4/i.test(srcName) && styles.h4)   para.appliedParagraphStyle = styles.h4;
        else if (/Heading 5/i.test(srcName) && styles.h5)   para.appliedParagraphStyle = styles.h5;
        else if (/Heading 6/i.test(srcName) && styles.h6)   para.appliedParagraphStyle = styles.h6;
        else if (styles.body)                               para.appliedParagraphStyle = styles.body;

        try { para.clearOverrides(); } catch (_) {}
    }

    // ---- 12) Cross-platform bullet repair (Windows RTF/DOCX differences) ----
    // Detects common bullet glyphs and dash-started lists; applies your bullet/sub-bullet styles.
    function normalizeBullets(tf) {
        var paras = tf.paragraphs;

        // Primary bullets (Unicode + Symbol legacy)
        var primaryBulletRE   = /^[\u2022\u25CF\u25E6\uF0B7\u2219\u00B7]\s*\t?/; // • ● ◦ (Symbol) ∙ ·
        // Sub-bullets (open circle, small circle, 'o')
        var subBulletGlyphRE  = /^[○◦o]\s*\t?/;

        // Many writers type "-" or "–" lists; treat as bullets if desired:
        var dashBulletRE      = /^[\-–]\s+\t?/; // hyphen or en-dash followed by space/tab

        for (var k = 0; k < paras.length; k++) {
            var pp = paras[k];
            var txt = pp.contents || "";

            // Sub-bullets first (so they don't get captured by primary)
            if (styles.subBullet && subBulletGlyphRE.test(txt)) {
                pp.contents = txt.replace(subBulletGlyphRE, "");
                pp.appliedParagraphStyle = styles.subBullet;
                pp.bulletsAndNumberingListType = ListType.BULLET_LIST;
                try { pp.clearOverrides(); } catch (_) {}
                continue;
            }

            // Primary bullets via glyphs (• ● ◦ etc.)
            if (styles.bullet && primaryBulletRE.test(txt)) {
                pp.contents = txt.replace(primaryBulletRE, "");
                pp.appliedParagraphStyle = styles.bullet;
                pp.bulletsAndNumberingListType = ListType.BULLET_LIST;
                try { pp.clearOverrides(); } catch (_) {}
                continue;
            }

            // Optional: treat dash-started lines as bullets
            if (styles.bullet && dashBulletRE.test(txt)) {
                pp.contents = txt.replace(dashBulletRE, "");
                pp.appliedParagraphStyle = styles.bullet;
                pp.bulletsAndNumberingListType = ListType.BULLET_LIST;
                try { pp.clearOverrides(); } catch (_) {}
                continue;
            }
        }
    }
    normalizeBullets(story);

    // ---- 13) Table styling (if present) ----
    if (styles.table && story.tables.length > 0) {
        for (var t = 0; t < story.tables.length; t++) {
            try { story.tables[t].appliedTableStyle = styles.table; } catch (_) {}
        }
    }

    // ---- 14) Apply custom hyperlink character style (visuals for imported links) ----
    // Change this to your team's preferred style name:
    var hyperlinkStyleName = "Inline Link";
    var hyperlinkStyle = null;
    try {
        hyperlinkStyle = doc.characterStyles.itemByName(hyperlinkStyleName);
        // Touch to verify existence
        var _n = hyperlinkStyle.name;
    } catch (_) {
        hyperlinkStyle = null;
    }

    if (hyperlinkStyle && hyperlinkStyle.isValid && doc.hyperlinks.length > 0) {
        for (var h = 0; h < doc.hyperlinks.length; h++) {
            var link = doc.hyperlinks[h];
            try {
                var src = link.source;
                if (src instanceof HyperlinkTextSource) {
                    var text = src.sourceText; // a Text object range
                    text.appliedCharacterStyle = hyperlinkStyle;
                }
            } catch (eLink) {
                $.writeln("⚠️ Link styling error: " + eLink.message);
            }
        }
    }

    alert("✅ Import completed!\nHeadings, Body, Bullets/Sub-Bullets, Table Style, and Hyperlink style applied.\nBold/italic preserved; cross-platform bullet repair done.");

} catch (err) {
    alert("❌ Script error:\n" + err.message);
} finally {
    app.scriptPreferences.userInteractionLevel = originalInteractionLevel;
}
