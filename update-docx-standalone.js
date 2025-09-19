import fs from "fs";
import path from "path";
import JSZip from "jszip";

const loadTranslationMapping = () => {
    const jsonData = fs.readFileSync("example_texts_complex.json", "utf8");
    const translations = JSON.parse(jsonData);
    const mapping = new Map();
    Object.values(translations).forEach(({ originalText, translatedText }) => {
        mapping.set(originalText.trim(), translatedText);
    });
    return mapping;
};

const updateXmlTextNodes = (xml, transform) =>
    xml.replace(
        /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g,
        (_, attrs, inner) =>
            `<w:t${attrs || ""}>${transform(inner || "")}</w:t>`
    );

const processDocx = async (inputPath, outputPath, transform) => {
    const buf = fs.readFileSync(inputPath);
    const zip = await new JSZip().loadAsync(buf);

    // Process main document
    const docEntry = zip.file("word/document.xml");
    if (!docEntry)
        throw new Error("document.xml not found - not a valid DOCX file");

    const docXml = await docEntry.async("string");
    zip.file("word/document.xml", updateXmlTextNodes(docXml, transform));

    // Process additional XML files
    const candidates = Object.keys(zip.files).filter((name) =>
        /^word\/(header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(
            name
        )
    );

    for (const name of candidates) {
        const entry = zip.file(name);
        if (!entry) continue;

        const xml = await entry.async("string");
        if (xml.includes("<w:t")) {
            zip.file(name, updateXmlTextNodes(xml, transform));
        }
    }

    const out = await zip.generateAsync({ type: "nodebuffer" });
    fs.writeFileSync(outputPath, out);
};

const createTransform = (translationMapping) => (text) => {
    if (!text?.trim()) return text;

    const trimmedText = text.trim();

    // Exact match
    if (translationMapping.has(trimmedText)) {
        return translationMapping.get(trimmedText);
    }

    // Partial match
    for (const [original, translated] of translationMapping.entries()) {
        if (trimmedText.toLowerCase().includes(original.toLowerCase())) {
            return translated;
        }
    }

    return text;
};

const main = async () => {
    const translationMapping = loadTranslationMapping();
    const transform = createTransform(translationMapping);
    await processDocx("input.docx", "output.docx", transform);
};

main();
