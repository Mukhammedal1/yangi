// Bir marta ishga tushiriladi:  node tools/font-yasash.js
// DejaVu .ttf fayllarini base64 JS modulga aylantiradi.

const fs = require("fs");
const path = require("path");

const manba = {
  "sans.js": "DejaVuSans.ttf",
  "sans-bold.js": "DejaVuSans-Bold.ttf",
};

const chiqish = path.join(__dirname, "..", "fonts");
fs.mkdirSync(chiqish, { recursive: true });

for (const [nom, ttf] of Object.entries(manba)) {
  const yol = require.resolve("dejavu-fonts-ttf/ttf/" + ttf);
  const b64 = fs.readFileSync(yol).toString("base64");

  fs.writeFileSync(
    path.join(chiqish, nom),
    `// ${ttf} — base64. require() bilan olinadi, shuning uchun\n` +
      `// Vercel bundler uni har doim ko'radi.\n` +
      `module.exports = Buffer.from(\n  "${b64}",\n  "base64"\n);\n`
  );

  console.log(nom, "tayyor");
}