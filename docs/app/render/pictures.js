/**
 * Picture placeholders as real images, drawn in the browser with a canvas.
 *
 * In Word, right-click a placeholder and choose Change Picture. In Google Docs, select it and
 * choose Replace image. The new picture takes the placeholder's size and position, the way
 * picture frames worked in Publisher.
 */
import { walkBlocks } from "../core/compose.js";
import { parseDataUrl } from "./zip.js";

function drawPlaceholder(block) {
  const scale = Math.min(2, 1400 / Math.max(block.width, block.height));
  const w = Math.max(8, Math.round(block.width * scale));
  const h = Math.max(8, Math.round(block.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const g = canvas.getContext("2d");
  g.fillStyle = block.fill;
  g.fillRect(0, 0, w, h);

  // Pictogram: a framed landscape with a sun.
  const s = Math.max(14, Math.min(w, h) * 0.26);
  const x = (w - s * 1.33) / 2;
  const y = (h - s) / 2 - (h > 90 ? s * 0.18 : 0);
  g.strokeStyle = block.ink;
  g.fillStyle = block.ink;
  g.globalAlpha = 0.8;
  g.lineWidth = Math.max(1.5, s * 0.06);
  g.strokeRect(x, y, s * 1.33, s);
  g.beginPath();
  g.arc(x + s * 0.95, y + s * 0.3, s * 0.11, 0, Math.PI * 2);
  g.fill();
  g.beginPath();
  g.moveTo(x + s * 0.08, y + s * 0.9);
  g.lineTo(x + s * 0.45, y + s * 0.45);
  g.lineTo(x + s * 0.7, y + s * 0.7);
  g.lineTo(x + s * 0.85, y + s * 0.56);
  g.lineTo(x + s * 1.25, y + s * 0.9);
  g.closePath();
  g.fill();

  if (block.label && h > 90) {
    const size = Math.max(11, Math.min(26, h / 12));
    g.globalAlpha = 0.85;
    g.font = `600 ${size}px system-ui, -apple-system, "Segoe UI", sans-serif`;
    g.textAlign = "center";
    g.textBaseline = "top";
    g.fillText(block.label, w / 2, y + s + size * 0.7);
  }
  return canvas.toDataURL("image/png");
}

/**
 * Returns a lookup function for docxBytes/flatPackage and the Google renderer:
 * image block -> { base64, mime }. Logos use the stored logo; placeholders are drawn.
 */
export function prepareImages(blocks) {
  const map = new Map();
  walkBlocks(blocks, (b) => {
    if (b.t !== "image") return;
    const url = b.src || drawPlaceholder(b);
    const parsed = parseDataUrl(url);
    if (parsed) map.set(b, parsed);
  });
  return (block) => map.get(block) || null;
}

/** Reads an image file into a data URL and its pixel size, scaled down to at most `max` pixels. */
export function readLogoFile(file, max = 800) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("The file could not be read."));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("That file is not a picture Offizin can use. Try a PNG or JPEG."));
      img.onload = () => {
        const ratio = Math.min(1, max / Math.max(img.width, img.height));
        const w = Math.max(1, Math.round(img.width * ratio));
        const h = Math.max(1, Math.round(img.height * ratio));
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        canvas.getContext("2d").drawImage(img, 0, 0, w, h);
        resolve({ src: canvas.toDataURL("image/png"), width: w, height: h });
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}
