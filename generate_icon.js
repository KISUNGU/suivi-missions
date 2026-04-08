const fs = require('node:fs');
const pngToIcoModule = require('png-to-ico');

const pngToIco = pngToIcoModule.default || pngToIcoModule.imagesToIco;

async function main() {
  const buffer = await pngToIco('logo_pnda_square.png');
  fs.writeFileSync('logo_pnda.ico', buffer);
  console.log('logo_pnda.ico generated');
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});