/**
 * Client-side evidence file parsing: plain text, markdown, CSV, JSON (FileReader),
 * PDF (pdfjs-dist), DOCX (mammoth).
 */

import * as pdfjsLib from "pdfjs-dist/build/pdf.mjs";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.mjs?url";
import mammoth from "mammoth";

let pdfWorkerConfigured = false;

function ensurePdfWorker() {
  if (pdfWorkerConfigured) return;
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
  pdfWorkerConfigured = true;
}

async function extractTextFromPdf(file) {
  ensurePdfWorker();
  const buf = await file.arrayBuffer();
  const data = new Uint8Array(buf);
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;
  const pageTexts = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const line = textContent.items
      .map((item) => ("str" in item ? item.str : ""))
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
    if (line) pageTexts.push(line);
  }
  return pageTexts.join("\n\n").trim();
}

async function extractTextFromDocx(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return String(result.value || "").trim();
}

/**
 * @returns {Promise<
 *   | { kind: "text"; text: string }
 *   | { kind: "failed"; error: string }
 * >}
 */
export async function parseEvidenceFileToText(file) {
  const name = file.name || "unnamed";
  const lower = name.toLowerCase();
  const isPdf = lower.endsWith(".pdf") || file.type === "application/pdf";
  const isDocx =
    lower.endsWith(".docx") ||
    file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  if (isPdf) {
    try {
      const text = await extractTextFromPdf(file);
      return {
        kind: "text",
        text: text || "[No extractable text found in this PDF.]",
      };
    } catch (e) {
      const msg = e?.message || String(e);
      return { kind: "failed", error: msg };
    }
  }

  if (isDocx) {
    try {
      const text = await extractTextFromDocx(file);
      return {
        kind: "text",
        text: text || "[No text extracted from this DOCX.]",
      };
    } catch (e) {
      const msg = e?.message || String(e);
      return { kind: "failed", error: msg };
    }
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const raw = String(reader.result ?? "");
      if (lower.endsWith(".json")) {
        try {
          const j = JSON.parse(raw);
          resolve({
            kind: "text",
            text: typeof j === "object" ? JSON.stringify(j, null, 2) : raw,
          });
        } catch {
          resolve({ kind: "text", text: raw });
        }
      } else {
        resolve({ kind: "text", text: raw });
      }
    };
    reader.onerror = () => reject(reader.error || new Error("File read failed"));
    reader.readAsText(file);
  });
}
