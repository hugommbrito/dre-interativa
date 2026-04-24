import { kv } from '@vercel/kv';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-dre-password');

  if (req.method === 'OPTIONS') return res.status(200).end();

  const senha = req.headers['x-dre-password'] || '';
  if (!process.env.DRE_PASSWORD_READ || senha !== process.env.DRE_PASSWORD_READ)
    return res.status(401).json({ error: 'Senha inválida' });

  const raw = await kv.get('dre_data');
  if (!raw) return res.status(404).json({ error: 'Sem dados no servidor — faça o upload do dre_data.json' });

  try {
    res.status(200).json(typeof raw === 'string' ? JSON.parse(raw) : raw);
  } catch {
    res.status(500).json({ error: 'Dados corrompidos no servidor — faça o upload novamente' });
  }
}
