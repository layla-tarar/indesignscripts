// TitleCaseHeadings.jsx
// Step 1: Applies Title Case to all paragraphs with specified heading styles.
// Step 2: Lowercases common articles, prepositions, and conjunctions
//         (unless they are the first word of the heading).
// Step 3: Fixes specific terms that Title Case breaks.
// Run from InDesign: File > Scripts > Scripts Panel, then double-click.

var doc = app.activeDocument;

var stylesToFix = [
    "Head_Section",
    "Head_SubsectionUnnumbered",
    "Head_SubSubSectionUnnumbered",
    "Head_SubsectionNumbered",
    "Head_CropName"
];

// Words to lowercase (when not the first word).
// APA 7th ed: lowercase articles, coordinating conjunctions, and prepositions
// under 4 letters. "From" (4), "Into" (4), "With" (4) are capitalized per APA.
var lowercaseWords = [
    "A", "An", "And", "As", "At", "But", "By", "For",
    "In", "Nor", "Of", "On", "Or", "Per", "So", "The", "To",
    "Up", "Via", "Vs", "Yet"
];

// Specific whole-word replacements
// Format: [find (what Title Case produces), replace (correct form)]
var specificFixes = [
    // Abbreviations
    ["Ge", "GE"],
    ["Dna", "DNA"],
    ["Esps", "ESPS"],
    ["Epsps", "EPSPS"],

    // Bt (Title Case lowercases the t)
    ["bt", "Bt"],

    // Protein names — Title Case lowercases internal capitals
    ["Cry1ab", "Cry1Ab"],
    ["Cry1ac", "Cry1Ac"],
    ["Cry1bb", "Cry1Bb"],
    ["Cry1f", "Cry1F"],
    ["Cry1fa", "Cry1Fa"],
    ["Cry2ab", "Cry2Ab"],
    ["Cry2aa", "Cry2Aa"],
    ["Cry3bb1", "Cry3Bb1"],
    ["Vip3aa", "Vip3Aa"],

    // Scientific names — Title Case capitalizes the species epithet
    // Full names
    ["Bacillus Thuringiensis", "Bacillus thuringiensis"],
    ["Zea Mays", "Zea mays"],
    ["Gossypium Hirsutum", "Gossypium hirsutum"],
    ["Oryza Sativa", "Oryza sativa"],
    ["Manduca Sexta", "Manduca sexta"],
    ["Vigna Unguiculata", "Vigna unguiculata"],
    ["Arabidopsis Thaliana", "Arabidopsis thaliana"],
    // Abbreviated names
    ["B. Thuringiensis", "B. thuringiensis"],
    ["Z. Mays", "Z. mays"],
    ["G. Hirsutum", "G. hirsutum"],
    ["O. Sativa", "O. sativa"],
    ["M. Sexta", "M. sexta"],
    ["V. Unguiculata", "V. unguiculata"],
    ["A. Thaliana", "A. thaliana"]
];

var titleCaseCount = 0;
var lowercaseCount = 0;
var specificFixCount = 0;

// --- STEP 1: Apply Title Case ---
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

for (var s = 0; s < stylesToFix.length; s++) {
    try {
        var style = doc.paragraphStyles.itemByName(stylesToFix[s]);
        app.findGrepPreferences = NothingEnum.nothing;
        app.findGrepPreferences.appliedParagraphStyle = style;
        var found = doc.findGrep();
        for (var i = 0; i < found.length; i++) {
            found[i].changecase(ChangecaseMode.TITLECASE);
            titleCaseCount++;
        }
    } catch (e) {
        // Style not found, skip
    }
}

// --- STEP 2: Lowercase articles/prepositions (not first word) ---

for (var s = 0; s < stylesToFix.length; s++) {
    try {
        var style = doc.paragraphStyles.itemByName(stylesToFix[s]);
        for (var w = 0; w < lowercaseWords.length; w++) {
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            app.findGrepPreferences.appliedParagraphStyle = style;
            app.findGrepPreferences.findWhat = "(?<=\\s)" + lowercaseWords[w] + "(?=\\s|$)";
            app.changeGrepPreferences.changeTo = lowercaseWords[w].toLowerCase();
            var results = doc.changeGrep();
            lowercaseCount += results.length;
        }
    } catch (e) {
        // Style not found, skip
    }
}

// --- STEP 3: Fix specific terms ---
// Uses GREP find/change scoped to heading styles.
// Multi-word terms (scientific names) use literal find text, not \b boundaries.
// Single-word terms use \b word boundaries for whole-word matching.

for (var s = 0; s < stylesToFix.length; s++) {
    try {
        var style = doc.paragraphStyles.itemByName(stylesToFix[s]);
        for (var f = 0; f < specificFixes.length; f++) {
            app.findGrepPreferences = NothingEnum.nothing;
            app.changeGrepPreferences = NothingEnum.nothing;
            app.findGrepPreferences.appliedParagraphStyle = style;

            var findTerm = specificFixes[f][0];
            var replaceTerm = specificFixes[f][1];

            // If the find term contains a space, it's multi-word (scientific name)
            // — match it literally. Otherwise use word boundaries.
            if (findTerm.indexOf(" ") >= 0) {
                // Multi-word: escape the period in abbreviated names
                var escaped = findTerm.replace(/\./g, "\\.");
                app.findGrepPreferences.findWhat = escaped;
            } else {
                app.findGrepPreferences.findWhat = "\\b" + findTerm + "\\b";
            }
            app.changeGrepPreferences.changeTo = replaceTerm;
            var results = doc.changeGrep();
            specificFixCount += results.length;
        }
    } catch (e) {
        // Style not found, skip
    }
}

// Clean up
app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

alert(
    "Done!\n\n" +
    "Title Case applied to: " + titleCaseCount + " headings\n" +
    "Words lowercased (articles/prepositions): " + lowercaseCount + "\n" +
    "Specific terms fixed: " + specificFixCount + "\n\n" +
    "Quick review recommended for any edge cases the script may have missed."
);
