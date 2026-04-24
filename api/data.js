export default function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-dre-password');

  if (req.method === 'OPTIONS') return res.status(200).end();

  const senha = req.headers['x-dre-password'] || '';
  if (!process.env.DRE_PASSWORD || senha !== process.env.DRE_PASSWORD)
    return res.status(401).json({ error: 'Senha inválida' });

  const raw = process.env.DRE_DATA_JSON || '';
  if (!raw) return res.status(404).json({ error: 'Sem dados configurados no servidor' });

  try {
    res.status(200).json(JSON.parse(raw));
  } catch {
    res.status(500).json({ error: 'Dados inválidos no servidor — verifique o env DRE_DATA_JSON' });
  }
}
