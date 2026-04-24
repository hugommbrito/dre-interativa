import { kv } from '@vercel/kv';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-dre-password');

  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método não permitido' });

  const senha = req.headers['x-dre-password'] || '';
  if (!process.env.DRE_PASSWORD || senha !== process.env.DRE_PASSWORD)
    return res.status(401).json({ error: 'Senha inválida' });

  const data = req.body;
  if (!data || !data.entradas || !data.saidas)
    return res.status(400).json({ error: 'JSON inválido — campos entradas e saidas são obrigatórios' });

  await kv.set('dre_data', JSON.stringify(data));
  return res.status(200).json({ ok: true });
}
