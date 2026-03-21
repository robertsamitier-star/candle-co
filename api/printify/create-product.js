const PRINTIFY_API = 'https://api.printify.com/v1';
const SHOP_ID = '26855509';
const PRINT_PROVIDER_ID = 219;

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const token = process.env.PRINTIFY_API_TOKEN;
  if (!token) {
    return res.status(500).json({ error: 'PRINTIFY_API_TOKEN not configured' });
  }

  const { title, description, tags, blueprint_id, variants, print_areas } = req.body || {};

  if (!title || !blueprint_id || !variants || !print_areas) {
    return res.status(400).json({ error: 'title, blueprint_id, variants, and print_areas are required' });
  }

  const validBlueprints = [1048, 1468, 1695, 2664];
  if (!validBlueprints.includes(blueprint_id)) {
    return res.status(400).json({ error: `Invalid blueprint_id. Must be one of: ${validBlueprints.join(', ')}` });
  }

  const enabledVariantIds = variants
    .filter(v => v.is_enabled)
    .map(v => v.id);

  if (enabledVariantIds.length === 0) {
    return res.status(400).json({ error: 'At least one variant must be enabled' });
  }

  const printifyPayload = {
    title,
    description: description || '',
    tags: tags || [],
    blueprint_id,
    print_provider_id: PRINT_PROVIDER_ID,
    variants: variants.map(v => ({
      id: v.id,
      price: v.price || 2999,
      is_enabled: v.is_enabled,
    })),
    print_areas: [
      {
        variant_ids: enabledVariantIds,
        placeholders: [
          {
            position: 'front',
            images: [
              {
                id: print_areas.front.image_id,
                x: print_areas.front.x ?? 0.5,
                y: print_areas.front.y ?? 0.5,
                scale: print_areas.front.scale ?? 1,
                angle: print_areas.front.angle ?? 0,
              },
            ],
          },
        ],
      },
    ],
  };

  try {
    const response = await fetch(`${PRINTIFY_API}/shops/${SHOP_ID}/products.json`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(printifyPayload),
    });

    const data = await response.json();

    if (!response.ok) {
      return res.status(response.status).json({ error: data.error || 'Product creation failed', details: data });
    }

    return res.status(200).json({
      id: data.id,
      title: data.title,
      images: data.images,
      variants_count: data.variants?.length || 0,
      enabled_count: enabledVariantIds.length,
    });
  } catch (err) {
    return res.status(500).json({ error: 'Failed to create product', message: err.message });
  }
}
