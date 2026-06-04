export const config = {
  api: { bodyParser: { sizeLimit: '5mb' } },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  try {
    const { audio, mimeType } = req.body;

    if (!process.env.OPENAI_API_KEY) {
      return res.status(200).json({ ok: true, transcript: '' });
    }

    const bytes = Buffer.from(audio, 'base64');
    const blob = new Blob([bytes], { type: mimeType || 'audio/webm' });

    const fd = new FormData();
    fd.append('file', blob, 'recording.webm');
    fd.append('model', 'whisper-1');

    const resp = await fetch('https://api.openai.com/v1/audio/transcriptions', {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + process.env.OPENAI_API_KEY },
      body: fd,
    });

    if (!resp.ok) {
      const err = await resp.text();
      return res.status(200).json({ ok: false, transcript: '', error: err });
    }

    const data = await resp.json();
    res.json({ ok: true, transcript: data.text || '' });
  } catch (err) {
    res.status(200).json({ ok: false, transcript: '', error: err.message });
  }
}
