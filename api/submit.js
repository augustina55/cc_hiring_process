const GAS_URL =
  'https://script.google.com/macros/s/AKfycbxoraWDXlBoXdXiN0A2-bIW9dBpnV3nM6HpV7cmRn9tM8ZJVV7iERIxsHInVuzEXJXs/exec';

export const config = {
  api: { bodyParser: false },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Method not allowed' });
  }

  try {
    // Read raw body so we don't impose a tiny size limit
    const raw = await new Promise((resolve, reject) => {
      const parts = [];
      req.on('data', c => parts.push(c));
      req.on('end', () => resolve(Buffer.concat(parts).toString('utf-8')));
      req.on('error', reject);
    });

    const gasResp = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: raw,
      redirect: 'follow',
    });

    const data = await gasResp.json();
    res.status(200).json(data);
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message });
  }
}
