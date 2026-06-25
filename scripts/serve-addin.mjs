import { createServer as createHttpServer } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import { existsSync, readFileSync } from "node:fs";
import { extname, join, normalize, resolve } from "node:path";

const args = process.argv.slice(2);
const options = parseArgs(args);
const root = resolve(options.root ?? "packages/addin");
const port = Number(options.port ?? 3000);

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
};

const handler = (request, response) => {
  const url = new URL(request.url ?? "/", `http://localhost:${port}`);
  const relativePath = url.pathname === "/" ? "taskpane/index.html" : url.pathname.slice(1);
  const target = resolve(root, normalize(relativePath));
  if (!target.startsWith(root)) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }
  if (!existsSync(target)) {
    response.writeHead(404);
    response.end("Not found");
    return;
  }
  response.writeHead(200, {
    "Content-Type": mimeTypes[extname(target)] ?? "application/octet-stream",
    "Cache-Control": "no-store",
  });
  response.end(readFileSync(target));
};

const certPath = options.cert ? resolve(options.cert) : undefined;
const keyPath = options.key ? resolve(options.key) : undefined;
const pfxPath = options.pfx ? resolve(options.pfx) : undefined;
const useHttps = pfxPath || (certPath && keyPath);
const httpsOptions = pfxPath
  ? { pfx: readFileSync(pfxPath), passphrase: options.passphrase }
  : certPath && keyPath
    ? { cert: readFileSync(certPath), key: readFileSync(keyPath) }
    : undefined;
const server = useHttps
  ? createHttpsServer(httpsOptions, handler)
  : createHttpServer(handler);

server.listen(port, () => {
  const protocol = useHttps ? "https" : "http";
  console.log(`mdpr-ppt add-in assets: ${protocol}://localhost:${port}/taskpane/index.html`);
  if (!useHttps) {
    console.log("Office manifests require HTTPS SourceLocation. Pass --cert and --key for PowerPoint sideloading.");
  }
  console.log(`Serving ${join(root, "")}`);
});

function parseArgs(values) {
  const parsed = {};
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (!value.startsWith("--")) throw new Error(`Unexpected argument: ${value}`);
    const key = value.slice(2);
    const next = values[index + 1];
    if (!next || next.startsWith("--")) {
      parsed[key] = "true";
    } else {
      parsed[key] = next;
      index += 1;
    }
  }
  return parsed;
}
