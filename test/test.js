const { embedFloatingOleObjects } = require("../index");

const oleFiles = [
    { filePath: "test/sample.xls", customIconPath: "test/custom_excel_icon.png", width: 120, height: 120, x: 100, y: 200 },
    { filePath: "test/report.pdf", width: 80, height: 80, x: 300, y: 150 }, // Uses default PDF icon
];

const outputDocx = "test/output.docx";

embedFloatingOleObjects(oleFiles, outputDocx)
    .then(() => console.log("Floating OLE objects embedded successfully!"))
    .catch((error) => console.error("Error:", error));
