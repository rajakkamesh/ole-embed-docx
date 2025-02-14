const fs = require("fs");

/**
 * Reads a file and returns its buffer.
 * @param {string} filePath - Path to the file.
 * @returns {Buffer}
 */
function readFile(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }
  return fs.readFileSync(filePath);
}

/**
 * Writes a buffer to a file.
 * @param {string} filePath - Path where the file should be saved.
 * @param {Buffer} buffer - Data to write.
 */
function writeFile(filePath, buffer) {
  fs.writeFileSync(filePath, buffer);
  console.log(`File saved at: ${filePath}`);
}

module.exports = { readFile, writeFile };
