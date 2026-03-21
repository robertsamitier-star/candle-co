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

  try {
    const body = req.body || {};
    const { title, description, tags, blueprint_id, variants, print_areas } = body;

    if (!title || !blueprint_id || !variants || !print_areas) {
      return res.status(400).json({ error: 'title, blueprint_id, variants, and print_areas are required' });
    }

    const validBlueprints = [1048, 1468, 1695, 2664];
    if (!validBlueprints.includes(blueprint_id)) {
      return res.status(400).json({ error: `Invalid blueprint_id. Must be one of: ${validBlueprints.join(', ')}` });
    }

    if (!Array.isArray(variants) || variants.length === 0) {
      return res.status(400).json({ error: 'variants must be a non-empty array' });
    }

    // Build the Printify payload
    // print_areas can come in two formats:
    // 1. Already in Printify format: [{ variant_ids, placeholders }]
    // 2. Simplified: { front: { image_id, x, y, scale, angle } }
    let printifyPrintAreas;

    if (Array.isArray(print_areas)) {
      // Frontend sends Printify format directly — pass through
      printifyPrintAreas = print_areas;
    } else if (print_areas.front) {
      // Simplified format — transform to Printify format
      const enabledIds = variants.filter(v => v.is_enabled).map(v => v.id);
      const area = {
        variant_ids: enabledIds,
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
      };
      if (print_areas.background) {
        area.background = print_areas.background;
      }
      printifyPrintAreas = [area];
    } else {
      return res.status(400).json({ error: 'print_areas must be an array or have a "front" key' });
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
        is_enabled: v.is_enabled !== undefined ? v.is_enabled : true,
      })),
      print_areas: printifyPrintAreas,
    };

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
      return res.status(response.status).json({
        error: data.error || data.message || 'Product creation failed',
        details: data,
      });
    }

    return res.status(200).json({
      id: data.id,
      title: data.title,
      images: data.images,
      variants_count: data.variants?.length || 0,
      enabled_count: variants.filter(v => v.is_enabled).length,
    });
  } catch (err) {
    return res.status(500).json({ error: 'Failed to create product', message: err.message });
  }
}
