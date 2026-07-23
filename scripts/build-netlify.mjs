import { mkdir, readFile, writeFile } from "node:fs/promises";

const backendUrl = (process.env.BACKEND_URL || "").replace(/\/$/, "");
if (!backendUrl) {
  throw new Error("BACKEND_URL is required for the Netlify build");
}

await mkdir("dist", { recursive: true });
const template = await readFile("netlify-site/index.html", "utf8");
const output = template.replaceAll("__BACKEND_URL__", backendUrl);
await writeFile("dist/index.html", output, "utf8");
