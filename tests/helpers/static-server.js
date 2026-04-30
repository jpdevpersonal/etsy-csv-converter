const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");

const { getHeadersForPath, loadHeadersRules } = require("./headers-config");

const ROOT_DIR = path.resolve(__dirname, "../..");
const HOST = process.env.HOST || "127.0.0.1";
const PORT = Number(process.env.PORT || 4173);
const HEADER_RULES = loadHeadersRules();

const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".gif": "image/gif",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".jpeg": "image/jpeg",
  ".jpg": "image/jpeg",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8",
  ".webp": "image/webp",
};

function resolveRequestPath(urlPathname) {
  const pathname = decodeURIComponent(urlPathname.split("?")[0]);
  const relativePath = pathname === "/" ? "/index.html" : pathname;
  const absolutePath = path.normalize(path.join(ROOT_DIR, relativePath));

  if (!absolutePath.startsWith(ROOT_DIR)) return null;
  return absolutePath;
}

function sendResponse(response, statusCode, body, headers = {}) {
  response.writeHead(statusCode, headers);
  response.end(body);
}

const server = http.createServer((request, response) => {
  const url = new URL(request.url, `http://${request.headers.host || `${HOST}:${PORT}`}`);
  const filePath = resolveRequestPath(url.pathname);

  if (!filePath) {
    sendResponse(response, 403, "Forbidden", { "Content-Type": "text/plain; charset=utf-8" });
    return;
  }

  let stat;
  try {
    stat = fs.statSync(filePath);
  } catch {
    sendResponse(response, 404, "Not Found", { "Content-Type": "text/plain; charset=utf-8" });
    return;
  }

  if (stat.isDirectory()) {
    sendResponse(response, 403, "Forbidden", { "Content-Type": "text/plain; charset=utf-8" });
    return;
  }

  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || "application/octet-stream";
  const headers = {
    "Content-Length": String(stat.size),
    "Content-Type": contentType,
    ...getHeadersForPath(url.pathname, HEADER_RULES),
  };

  if (request.method === "HEAD") {
    response.writeHead(200, headers);
    response.end();
    return;
  }

  const stream = fs.createReadStream(filePath);
  response.writeHead(200, headers);
  stream.pipe(response);
  stream.on("error", () => {
    if (!response.headersSent) {
      sendResponse(response, 500, "Internal Server Error", { "Content-Type": "text/plain; charset=utf-8" });
    } else {
      response.destroy();
    }
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Static test server running at http://${HOST}:${PORT}`);
});

function shutdown() {
  server.close(() => process.exit(0));
}

process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);