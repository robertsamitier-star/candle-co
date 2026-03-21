const PRINTIFY_API = 'https://api.printify.com/v1';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const token = process.env.PRINTIFY_API_TOKEN;
  if (!token) {
    return res.status(500).json({ error: 'PRINTIFY_API_TOKEN not configured' });
  }

  const { file_name, contents } = req.body || {};

  if (!file_name || !contents) {
    return res.status(400).json({ error: 'file_name and contents (base64) are required' });
  }

  if (contents.length > 10 * 1024 * 1024) {
    return res.status(400).json({ error: 'Image exceeds 10MB limit' });
  }

  const allowedExtensions = ['.png', '.jpg', '.jpeg', '.webp'];
  const ext = file_name.toLowerCase().slice(file_name.lastIndexOf('.'));
  if (!allowedExtensions.includes(ext)) {
    return res.status(400).json({ error: 'Only PNG, JPG, and WebP files are accepted' });
  }

  try {
    const response = await fetch(`${PRINTIFY_API}/uploads/images.json`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ file_name, contents }),
    });

    const data = await response.json();

    if (!response.ok) {
      return res.status(response.status).json({ error: data.error || 'Printify upload failed', details: data });
    }

    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ error: 'Failed to upload image', message: err.message });
  }
}
