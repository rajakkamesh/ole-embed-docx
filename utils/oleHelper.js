const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const { readFile, writeFile } = require("./fileHelper");
const { v4: uuidv4 } = require("uuid");

/**
 * Determines the appropriate icon for the file type.
 * @param {string} filePath - Path to the file (e.g., .xls, .pdf).
 * @param {string} customIconPath - Optional custom icon path.
 * @returns {string} - Path to the icon file.
 */
function getIconForFile(filePath, customIconPath) {
  if (customIconPath && fs.existsSync(customIconPath)) {
    return customIconPath; // Use user-provided icon
  }

  const ext = path.extname(filePath).toLowerCase();
  const defaultIcons = {
    ".xls": "assets/excel_icon.png",
    ".xlsx": "assets/excel_icon.png",
    ".pdf": "assets/pdf_icon.png",
    ".doc": "assets/word_icon.png",
    ".docx": "assets/word_icon.png",
    ".ppt": "assets/ppt_icon.png",
    ".pptx": "assets/ppt_icon.png",
  };

  return defaultIcons[ext] || "assets/default_icon.png"; // Use default or fallback icon
}

/**
 * Embeds multiple floating OLE objects into a DOCX file.
 * @param {Array} oleFiles - List of { filePath, customIconPath, width, height, x, y } objects.
 * @param {string} outputDocxPath - Path to save the DOCX.
 */
async function embedFloatingOleObjects(oleFiles, outputDocxPath) {
  let doc = new Document();

  // Load the DOCX structure
  const docBuffer = await Packer.toBuffer(doc);
  const zip = await JSZip.loadAsync(docBuffer);

  const relsPath = "word/_rels/document.xml.rels";
  let relsXML = await zip.file(relsPath).async("text");

  const documentXMLPath = "word/document.xml";
  let documentXML = await zip.file(documentXMLPath).async("text");

  for (let i = 0; i < oleFiles.length; i++) {
    const {
      filePath,
      customIconPath,
      width = 100,
      height = 100,
      x = 0,
      y = 0,
    } = oleFiles[i];
    const oleFileName = path.basename(filePath);
    const oleFileBuffer = readFile(filePath);
    const oleObjectId = `rId${uuidv4().split("-")[0]}`;

    // Load icon
    const iconPath = getIconForFile(filePath, customIconPath);
    const iconFileBuffer = readFile(iconPath);
    const iconFileName = path.basename(iconPath);
    const iconId = `rId${uuidv4().split("-")[0]}`;

    // Add OLE file
    const olePath = `word/embeddings/${oleFileName}`;
    zip.file(olePath, oleFileBuffer);

    // Add icon
    const iconPathInDocx = `word/media/${iconFileName}`;
    zip.file(iconPathInDocx, iconFileBuffer);

    // Update relationships
    relsXML = relsXML.replace(
      "</Relationships>",
      `<Relationship Id="${oleObjectId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/${oleFileName}"/>\n
             <Relationship Id="${iconId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${iconFileName}"/>\n</Relationships>`
    );

    // Add floating OLE object and icon with custom positioning to document.xml
    documentXML = documentXML.replace(
      "</w:body>",
      `<w:p>
                <w:r>
                    <w:drawing>
                        <wp:anchor behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                            <wp:simplePos x="${x}" y="${y}"/>
                            <wp:extent cx="${width * 9525}" cy="${
        height * 9525
      }"/>
                            <wp:docPr id="1" name="OLEObject${i}"/>
                            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:blipFill>
                                            <a:blip r:embed="${iconId}"/>
                                        </pic:blipFill>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:anchor>
                    </w:drawing>
                </w:r>
            </w:p>
            <w:object w:anchorId="${oleObjectId}">
                <v:shape id="${oleObjectId}" style="width:${width}px;height:${height}px">
                    <o:OLEObject Type="Embed" ShapeID="${oleObjectId}" DrawAspect="Icon" ObjectID="${oleObjectId}" r:id="${oleObjectId}"/>
                </v:shape>
            </w:object>\n</w:body>`
    );
  }

  // Update files in ZIP
  zip.file(relsPath, relsXML);
  zip.file(documentXMLPath, documentXML);

  // Save final DOCX
  const finalBuffer = await zip.generateAsync({ type: "nodebuffer" });
  writeFile(outputDocxPath, finalBuffer);
}

module.exports = { embedFloatingOleObjects };
