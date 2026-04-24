export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-dre-password');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método não permitido' });

  const senha = req.headers['x-dre-password'] || '';
  if (!process.env.DRE_PASSWORD_READ || senha !== process.env.DRE_PASSWORD_READ)
    return res.status(401).json({ error: 'Senha inválida' });

  if (!process.env.ANTHROPIC_API_KEY)
    return res.status(503).json({ error: 'Análise IA não configurada no servidor' });

  const { prompt } = req.body || {};
  if (!prompt || typeof prompt !== 'string')
    return res.status(400).json({ error: 'Prompt inválido' });

  const upstream = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json',
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 3500,
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  if (!upstream.ok) {
    const err = await upstream.json().catch(() => ({}));
    return res.status(upstream.status).json({ error: err.error?.message || `HTTP ${upstream.status}` });
  }

  const data = await upstream.json();
  return res.status(200).json({ text: data.content[0].text });
}
