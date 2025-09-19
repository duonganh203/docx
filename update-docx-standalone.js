import fs from "fs";
import path from "path";
import JSZip from "jszip";

const LOREM_WORDS = [
    "Lorem",
    "ipsum",
    "dolor",
    "sit",
    "amet",
    "consectetur",
    "adipiscing",
    "elit",
    "sed",
    "do",
    "eiusmod",
    "tempor",
    "incididunt",
    "ut",
    "labore",
    "et",
    "dolore",
    "magna",
    "aliqua",
    "Ut",
    "enim",
    "ad",
    "minim",
    "veniam",
    "quis",
    "nostrud",
    "exercitation",
    "ullamco",
    "laboris",
    "nisi",
    "aliquip",
    "ex",
    "ea",
    "commodo",
    "consequat",
    "Duis",
    "aute",
    "irure",
    "in",
    "reprehenderit",
    "voluptate",
    "velit",
    "esse",
    "cillum",
    "fugiat",
    "nulla",
    "pariatur",
    "Excepteur",
    "sint",
];

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

const transform = (text) => {
    if (!text?.trim()) return text;

    const wordCount = Math.max(1, text.trim().split(/\s+/).length);
    return Array.from(
        { length: wordCount },
        (_, i) => LOREM_WORDS[i % LOREM_WORDS.length]
    ).join(" ");
};

const main = async () => {
    const [, , inArg, outArg] = process.argv;

    if (!inArg) {
        console.error(
            "Usage: node update-docx-standalone.js <input.docx> [output.docx]"
        );
        process.exit(1);
    }

    const inputPath = path.resolve(inArg);
    const outputPath = outArg
        ? path.resolve(outArg)
        : path.join(
              path.dirname(inputPath),
              path.basename(inputPath, path.extname(inputPath)) + ".vi.docx"
          );

    if (!fs.existsSync(inputPath)) {
        console.error(`Error: Input file does not exist: ${inputPath}`);
        process.exit(1);
    }

    try {
        await processDocx(inputPath, outputPath, transform);
        console.log(`✅ Successfully processed: ${outputPath}`);
    } catch (error) {
        console.error(`❌ Error: ${error.message}`);
        process.exit(1);
    }
};

main().catch((err) => {
    console.error(`❌ Unexpected error: ${err.message}`);
    process.exit(1);
});
