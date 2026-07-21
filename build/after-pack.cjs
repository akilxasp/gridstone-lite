const { execFileSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");

// ponytail: electron-builder's electronLanguages flag no-ops here, so prune
// the Chromium locale.pak files by hand. Keep only these.
const KEEP_LOCALES = new Set(["en", "en_GB"]);

exports.default = async function afterPack(context) {
  if (context.electronPlatformName !== "darwin") return;

  const appPath = path.join(context.appOutDir, `${context.packager.appInfo.productFilename}.app`);

  const localesDir = path.join(
    appPath,
    "Contents/Frameworks/Electron Framework.framework/Versions/A/Resources",
  );
  for (const entry of fs.readdirSync(localesDir)) {
    if (!entry.endsWith(".lproj")) continue;
    if (KEEP_LOCALES.has(entry.slice(0, -".lproj".length))) continue;
    fs.rmSync(path.join(localesDir, entry), { recursive: true, force: true });
  }

  // Ad-hoc sign every mac build. identity:null skips signing entirely, and an
  // unsigned app won't launch on Apple Silicon. (Universal builds also need
  // this because the lipo merge invalidates any prior signature.)
  const entitlements = path.join(context.packager.projectDir, "build", "entitlements.mac.plist");
  execFileSync("/usr/bin/codesign", [
    "--force",
    "--deep",
    "--sign", "-",
    "--options", "runtime",
    "--entitlements", entitlements,
    appPath,
  ], { stdio: "inherit" });
};
