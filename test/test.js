const { embedOleObject } = require("../index");

// Test the embedding function
const oleFile = "test/sample.xls"; // Replace with actual file path
const outputDocx = "test/output.docx";

embedOleObject(oleFile, outputDocx)
  .then(() => console.log("OLE object embedded successfully!"))
  .catch((error) => console.error("Error:", error));
