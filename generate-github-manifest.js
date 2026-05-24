const fs = require("fs");
const path = require("path");

const [user, repo] = process.argv.slice(2);

if (!user || !repo) {
  console.error("Uso: npm run manifest:github -- USUARIO REPO");
  process.exit(1);
}

const root = path.join(__dirname, "..");
const templatePath = path.join(root, "manifest.github-pages.template.xml");
const outputPath = path.join(root, "manifest.github-pages.xml");

const manifest = fs
  .readFileSync(templatePath, "utf8")
  .replaceAll("USUARIO", user)
  .replaceAll("REPO", repo);

fs.writeFileSync(outputPath, manifest, "utf8");
console.log(`Creado manifest.github-pages.xml para https://${user}.github.io/${repo}/`);
