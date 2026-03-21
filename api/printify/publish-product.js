const PRINTIFY_API = 'https://api.printify.com/v1';
const SHOP_ID = '26855509';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const token = process.env.PRINTIFY_API_TOKEN;
  if (!token) {
    return res.status(500).json({ error: 'PRINTIFY_API_TOKEN not configured' });
  }

  const { product_id } = req.body || {};

  if (!product_id) {
    return res.status(400).json({ error: 'product_id is required' });
  }

  try {
    const response = await fetch(
      `${PRINTIFY_API}/shops/${SHOP_ID}/products/${product_id}/publish.json`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          title: true,
          description: true,
          images: true,
          variants: true,
          tags: true,
        }),
      }
    );

    if (response.status === 404) {
      return res.status(404).json({ error: 'Product not found. Check the product_id.' });
    }

    const contentType = response.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      const data = await response.json();
      if (!response.ok) {
        const isDisconnected = JSON.stringify(data).toLowerCase().includes('disconnect');
        if (isDisconnected) {
          return res.status(400).json({
            error: 'Etsy shop is not connected',
            message: 'Connect your Etsy shop in Printify Dashboard > Manage My Stores before publishing.',
          });
        }
        return res.status(response.status).json({ error: data.error || 'Publish failed', details: data });
      }
      return res.status(200).json({ success: true, data });
    }

    if (response.ok) {
      return res.status(200).json({ success: true });
    }

    return res.status(response.status).json({ error: 'Publish failed with unexpected response' });
  } catch (err) {
    return res.status(500).json({ error: 'Failed to publish product', message: err.message });
  }
}
