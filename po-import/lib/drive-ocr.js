'use strict';

const { getAccessToken } = require('./google-auth');

/**
 * OCR an image through Google Drive.
 *
 * Drive converts an uploaded image to a Doc and runs its own OCR; exporting that Doc as
 * text gives the recognised content. Nothing is retained — the temporary Doc is deleted in
 * a finally block even when the export fails.
 *
 * The OCR engine is Google's, running inside the user's own Drive. No third-party service
 * and no AI model of ours is involved.
 */

const UPLOAD_URL =
  'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocr=true&ocrLanguage=en';
const BOUNDARY = '----ypppoPoImport';

function multipartBody(metadata, blob, mime) {
  return Buffer.concat([
    Buffer.from(`--${BOUNDARY}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`),
    Buffer.from(JSON.stringify(metadata)),
    Buffer.from(`\r\n--${BOUNDARY}\r\nContent-Type: ${mime}\r\n\r\n`),
    blob,
    Buffer.from(`\r\n--${BOUNDARY}--\r\n`),
  ]);
}

async function deleteFile(id, token) {
  try {
    await fetch(`https://www.googleapis.com/drive/v3/files/${id}`, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}` },
    });
    return true;
  } catch {
    return false;
  }
}

/** OCR one PNG buffer, returning its plain text. */
async function ocrImage(png, { name = 'po-ocr-temp' } = {}) {
  const token = await getAccessToken();

  const uploadRes = await fetch(UPLOAD_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': `multipart/related; boundary=${BOUNDARY}`,
    },
    body: multipartBody(
      { name, mimeType: 'application/vnd.google-apps.document' },
      png,
      'image/png'
    ),
  });

  if (!uploadRes.ok) {
    throw new Error(`Drive upload failed (${uploadRes.status}): ${await uploadRes.text()}`);
  }

  const { id } = await uploadRes.json();

  try {
    const exportRes = await fetch(
      `https://www.googleapis.com/drive/v3/files/${id}/export?mimeType=text/plain`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!exportRes.ok) {
      throw new Error(`Drive export failed (${exportRes.status}): ${await exportRes.text()}`);
    }
    return await exportRes.text();
  } finally {
    // Never leave scratch files in the user's Drive, even on the error path.
    const gone = await deleteFile(id, token);
    if (!gone) console.warn(`  ! could not delete temp Drive file ${id} — remove it manually`);
  }
}

/** OCR every rasterised page and concatenate, in order. */
async function ocrPages(pages, { onPage } = {}) {
  const parts = [];
  for (const p of pages) {
    const text = await ocrImage(p.png, { name: `po-ocr-p${p.page}` });
    parts.push(text);
    if (onPage) onPage(p.page, pages.length, text.length);
  }
  return parts.join('\n');
}

module.exports = { ocrImage, ocrPages };
