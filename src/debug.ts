function log(type: string, message: string) {
  if (type === "error") {
    console.error(`[ERROR] ${message}`);
  } else if (type === "warn") {
    console.warn(`[WARNING] ${message}`);
  } else if (type === "info") {
    console.info(`[INFO] ${message}`);
  } else {
    console.log(`[DEBUG] ${message}`);
  }
}
