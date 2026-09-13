export default async function handler(request, response) {
  if (request.method !== 'POST') {
    response.setHeader('Allow', 'POST');
    return response.status(405).json({ ok: false, error: 'Method tidak dibenarkan.' });
  }

  const scriptUrl = process.env.ERPH_SCRIPT_URL;
  const apiToken = process.env.ERPH_API_TOKEN;
  if (!scriptUrl || !apiToken) {
    return response.status(500).json({ ok: false, error: 'Tetapan sambungan eRPH belum lengkap di Vercel.' });
  }

  try {
    const upstream = await fetch(scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ ...request.body, apiToken })
    });
    const raw = await upstream.text();
    let body;
    try { body = JSON.parse(raw); }
    catch { body = { ok: false, error: 'Apps Script memulangkan respons yang tidak sah.' }; }
    return response.status(upstream.ok ? 200 : 502).json(body);
  } catch (error) {
    return response.status(502).json({ ok: false, error: `Sambungan ke Apps Script gagal: ${error.message}` });
  }
}
