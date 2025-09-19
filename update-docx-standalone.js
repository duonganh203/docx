import fs from "fs";
import JSZip from "jszip";

// Load translations from JSON file
const loadTranslations = () => {
    const jsonContent = fs.readFileSync("example_texts_complex.json", "utf8");
    const data = JSON.parse(jsonContent);

    // Create Map for fast lookup: original text -> translated text
    const translationMap = new Map();

    Object.values(data).forEach((item) => {
        const originalText = item.originalText.trim();
        const translatedText = item.translatedText;
        translationMap.set(originalText, translatedText);
    });

    return translationMap;
};

// Find and replace text in XML
const replaceTextInXml = (xmlContent, translateFunction) => {
    // Find all <w:t>...</w:t> tags and replace content inside
    return xmlContent.replace(
        /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g,
        (fullMatch, attributes, textContent) => {
            const translatedText = translateFunction(textContent || "");
            return `<w:t${attributes || ""}>${translatedText}</w:t>`;
        }
    );
};

// Translate a text segment
const translateText = (text, translationMap) => {
    // If text is empty, return original
    if (!text || !text.trim()) {
        return text;
    }

    const cleanText = text.trim();

    // Check exact match first
    if (translationMap.has(cleanText)) {
        return translationMap.get(cleanText);
    }

    // If no exact match, find partial match
    for (const [originalText, translatedText] of translationMap.entries()) {
        if (cleanText.toLowerCase().includes(originalText.toLowerCase())) {
            return translatedText;
        }
    }

    // No translation found, return original text
    return text;
};

// Process DOCX file
const translateDocxFile = async (inputFile, outputFile, translationMap) => {
    // Read DOCX file as ZIP
    const fileBuffer = fs.readFileSync(inputFile);
    const zip = await new JSZip().loadAsync(fileBuffer);

    // Create translate function
    const translateFunction = (text) => translateText(text, translationMap);

    // Process main document.xml file
    const mainDocument = zip.file("word/document.xml");
    if (!mainDocument) {
        throw new Error("document.xml not found - invalid DOCX file");
    }

    const mainXmlContent = await mainDocument.async("string");
    const translatedXmlContent = replaceTextInXml(
        mainXmlContent,
        translateFunction
    );
    zip.file("word/document.xml", translatedXmlContent);

    // Process other XML files (header, footer, footnotes...)
    const xmlFiles = Object.keys(zip.files).filter((fileName) =>
        /^word\/(header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(
            fileName
        )
    );

    for (const fileName of xmlFiles) {
        const xmlFile = zip.file(fileName);
        if (!xmlFile) continue;

        const xmlContent = await xmlFile.async("string");

        // Only process if file contains text to translate
        if (xmlContent.includes("<w:t")) {
            const translatedContent = replaceTextInXml(
                xmlContent,
                translateFunction
            );
            zip.file(fileName, translatedContent);
        }
    }

    // Save translated file
    const outputBuffer = await zip.generateAsync({ type: "nodebuffer" });
    fs.writeFileSync(outputFile, outputBuffer);

    console.log(`Translation completed: ${inputFile} -> ${outputFile}`);
};

// Main program
const main = async () => {
    try {
        // Load translations from JSON file
        const translationMap = loadTranslations();

        // Translate DOCX file
        await translateDocxFile("input.docx", "output.docx", translationMap);

        console.log("Done!");
    } catch (error) {
        console.error("Error:", error.message);
    }
};

// Run program
main();
