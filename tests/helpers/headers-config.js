const fs = require("node:fs");
const path = require("node:path");

const ROOT_DIR = path.resolve(__dirname, "../..");
const HEADERS_PATH = path.join(ROOT_DIR, "_headers.example");

function parseHeadersConfig(text) {
  const rules = [];
  let currentRule = null;

  text.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) return;

    if (!/^\s/.test(line)) {
      currentRule = { pattern: trimmed, headers: {} };
      rules.push(currentRule);
      return;
    }

    if (!currentRule) return;

    const separatorIndex = trimmed.indexOf(":");
    if (separatorIndex < 0) return;

    const name = trimmed.slice(0, separatorIndex).trim();
    const value = trimmed.slice(separatorIndex + 1).trim();

    if (name) currentRule.headers[name] = value;
  });

  return rules;
}

function loadHeadersRules() {
  return parseHeadersConfig(fs.readFileSync(HEADERS_PATH, "utf8"));
}

function matchesPattern(pattern, pathname) {
  if (pattern === "/*") return true;
  if (pattern.endsWith("*")) return pathname.startsWith(pattern.slice(0, -1));
  return pathname === pattern;
}

function getHeadersForPath(pathname, rules = loadHeadersRules()) {
  return rules.reduce((headers, rule) => {
    if (!matchesPattern(rule.pattern, pathname)) return headers;
    return { ...headers, ...rule.headers };
  }, {});
}

function parseCspDirectives(cspValue) {
  return String(cspValue || "")
    .split(";")
    .map((directive) => directive.trim())
    .filter(Boolean)
    .reduce((directives, directive) => {
      const [name, ...parts] = directive.split(/\s+/);
      directives[name] = parts;
      return directives;
    }, {});
}

module.exports = {
  HEADERS_PATH,
  getHeadersForPath,
  loadHeadersRules,
  parseCspDirectives,
  parseHeadersConfig,
};