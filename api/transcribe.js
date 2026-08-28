import ffmpegPath from 'ffmpeg-static';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';

const execFileAsync = promisify(execFile);

export const config = {
  api: { bodyParser: { sizeLimit: '1mb' } }, // just a fileId now, not video bytes
};

// Google Drive shows a "can't scan this file for viruses" interstitial HTML
// page instead of the raw file for larger downloads — extract the confirm
// token it embeds and retry with it to get the actual bytes.
async function downloadDriveFile(fileId, destPath) {
  const base = 'https://drive.google.com/uc?export=download&id=' + fileId;
  let resp = await fetch(base, { redirect: 'follow' });
  let buf = Buffer.from(await resp.arrayBuffer());

  const head = buf.slice(0, 4000).toString('utf-8');
  if (head.includes('confirm=')) {
    const match = head.match(/confirm=([0-9A-Za-z_]+)/);
    if (match) {
      const cookie = resp.headers.get('set-cookie') || '';
      resp = await fetch(base + '&confirm=' + match[1], {
        headers: cookie ? { Cookie: cookie } : {},
        redirect: 'follow',
      });
      buf = Buffer.from(await resp.arrayBuffer());
    }
  }

  await fs.writeFile(destPath, buf);
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  const { fileId } = req.body || {};
  if (!fileId) return res.status(400).json({ ok: false, error: 'fileId required' });

  if (!process.env.GROQ_API_KEY) {
    return res.status(500).json({ ok: false, error: 'GROQ_API_KEY not configured' });
  }

  const tmp = os.tmpdir();
  const inputPath = path.join(tmp, fileId + '_in');
  const audioPath = path.join(tmp, fileId + '_audio.mp3');

  try {
    await downloadDriveFile(fileId, inputPath);

    // Downmix to mono 64kbps — a multi-minute video's audio track comes out
    // to a few MB this way, comfortably under Groq's ~25MB transcription
    // upload limit that the full video routinely blows past.
    await execFileAsync(ffmpegPath, [
      '-y', '-i', inputPath,
      '-vn', '-ac', '1', '-b:a', '64k',
      audioPath,
    ]);

    const audioBuf = await fs.readFile(audioPath);

    const fd = new FormData();
    fd.append('file', new Blob([audioBuf], { type: 'audio/mpeg' }), 'audio.mp3');
    fd.append('model', 'whisper-large-v3-turbo');

    const resp = await fetch('https://api.groq.com/openai/v1/audio/transcriptions', {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + process.env.GROQ_API_KEY },
      body: fd,
    });

    const text = await resp.text();
    if (!resp.ok) return res.status(502).json({ ok: false, error: text.slice(0, 500) });

    const data = JSON.parse(text);
    res.status(200).json({ ok: true, transcript: data.text || '' });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  } finally {
    await fs.rm(inputPath, { force: true }).catch(() => {});
    await fs.rm(audioPath, { force: true }).catch(() => {});
  }
}
