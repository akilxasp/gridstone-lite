const { execFileSync } = require("node:child_process");
const path = require("node:path");

exports.default = async function afterPack(context) {
  if (context.electronPlatformName !== "darwin" || !context.appOutDir.endsWith("mac-universal")) return;

  const appPath = path.join(context.appOutDir, `${context.packager.appInfo.productFilename}.app`);
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
