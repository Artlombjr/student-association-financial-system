// Locations of the Apps Script sources, relative to this folder.
const path = require("path");
const ROOT = path.join(__dirname, "..", "apps-script");
const WEB = path.join(ROOT, "claim-web-app");
const TOOLS = path.join(ROOT, "sheet-tools");
const WHERE = {
  "Code.gs": WEB, "Form.html": WEB, "appsscript.json": WEB,
};
const file = name => path.join(WHERE[name] || TOOLS, name);   // every other file is a sheet-bound script
module.exports = { file, ROOT, WEB, TOOLS };
