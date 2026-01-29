const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const previewPath = path.join(root, "template-preview.html");
const templatePath = path.join(root, "template.html");

const STYLE_START = "/* SYNC:STYLE:START */";
const STYLE_END = "/* SYNC:STYLE:END */";
const BODY_START = "<!-- SYNC:BODY:START -->";
const BODY_END = "<!-- SYNC:BODY:END -->";

function extractBetween(content, startMarker, endMarker) {
  const startIdx = content.indexOf(startMarker);
  if (startIdx === -1) throw new Error(`Missing start marker: ${startMarker}`);
  const endIdx = content.indexOf(endMarker);
  if (endIdx === -1) throw new Error(`Missing end marker: ${endMarker}`);
  const startContentIdx = startIdx + startMarker.length;
  return content.slice(startContentIdx, endIdx);
}

function replaceBetween(content, startMarker, endMarker, replacement) {
  const startIdx = content.indexOf(startMarker);
  if (startIdx === -1) throw new Error(`Missing start marker: ${startMarker}`);
  const endIdx = content.indexOf(endMarker);
  if (endIdx === -1) throw new Error(`Missing end marker: ${endMarker}`);
  const startContentIdx = startIdx + startMarker.length;
  return content.slice(0, startContentIdx) + replacement + content.slice(endIdx);
}

const previewHtml = fs.readFileSync(previewPath, "utf8");
const templateHtml = fs.readFileSync(templatePath, "utf8");

const styleBlock = extractBetween(previewHtml, STYLE_START, STYLE_END);
const bodyBlock = extractBetween(previewHtml, BODY_START, BODY_END);

let nextTemplate = templateHtml;
nextTemplate = replaceBetween(nextTemplate, STYLE_START, STYLE_END, styleBlock);
nextTemplate = replaceBetween(nextTemplate, BODY_START, BODY_END, bodyBlock);

fs.writeFileSync(templatePath, nextTemplate, "utf8");

console.log("Synced template.html from template-preview.html");
