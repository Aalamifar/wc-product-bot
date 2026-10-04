# wc-product-bot

Bulk-create WooCommerce products from an Excel spreadsheet, using one
existing product as a "template" for price, type, categories and images,
and filling in the title / slug / description / short description from
your own customizable text templates.

No browser automation — it talks directly to the WooCommerce REST API.

## Why

If you sell many similar products that differ only in a few details
(a model number, a capacity, a part number, ...), this lets you keep one
row per product in a spreadsheet and generate all of them from a single
template product + a handful of text templates, instead of duplicating
and hand-editing each one in wp-admin.

## Requirements

- Python 3.9+
- A WooCommerce site with the REST API enabled (enabled by default)
- A WooCommerce REST API key with **Read/Write** permissions

## Setup

```bash
pip install -r requirements.txt
cp config.example.ini config.ini
```

Edit `config.ini`:

1. **`[woocommerce]`** — your site URL and REST API key. To create a key:
   WooCommerce → Settings → Advanced → REST API → Add key, permissions
   **Read/Write**. The secret is shown only once — copy it immediately.
2. **`[template_product]`** — the ID of an existing product to copy
   price/type/categories/images from (the number in `post.php?post=ID`
   when editing it in wp-admin), and which fields to copy.
3. **`[excel]`** — path to your spreadsheet, which column marks a row as
   already processed, and the header row number.
4. **`[columns]`** — map each column letter in your spreadsheet to a
   placeholder name. These names are what you use in your template files,
   e.g. mapping `brand = B` lets you write `{brand}` in a template.
5. **`[templates]`** — paths to your title / description / short
   description template files (plain text files with `{placeholder}`
   markers, filled with Python's `str.format`), and an optional slug
   template.
6. **`[seo]`** — optionally set the Rank Math "Focus Keyword" meta field
   on each new product.

The `templates/` folder has example templates for a "BIOS file" product
(in Gutenberg block format, so the content shows up as proper
heading/paragraph/list blocks instead of one raw HTML blob when you open
it in the block editor) — replace them with your own.

`sample/sample_products.xlsx` shows the expected spreadsheet shape for
the example config.

## Usage

Always test before running for real:

```bash
# Build the payloads and print them, without calling the API or touching the Excel file
python wc_product_bot.py --config config.ini --dry-run

# Create just the first pending row, as a draft (set publish_status = draft in config.ini first)
python wc_product_bot.py --config config.ini --limit 1

# Once you've checked the result in wp-admin and are happy with it,
# set publish_status = publish in config.ini, then run the rest
python wc_product_bot.py --config config.ini
```

The script skips rows whose status column already equals `done_value`
in your config, so re-running it only processes new rows. Failed rows
are marked with `error_value` so you can find and retry them.

## How a product is built

For each pending row:

1. The template product is fetched once (`GET /products/{id}`).
2. A new payload is built: title/slug/description/short description come
   from your templates filled with that row's data; price, type,
   categories, images and any other fields listed in `copy_fields` are
   copied from the template product.
3. The product is created (`POST /products`) with the status from
   `publish_status` (`draft` while testing, `publish` once you trust it).

**Note on downloadable products:** if your products are WooCommerce
downloadable products, the `downloads` field (the actual file URL) is
*not* set automatically — it isn't something a generic Excel column can
safely guess, and WordPress Media Library or the Woo API both need a
media/file URL. If every product has a different file, add a column for
it in your spreadsheet, a placeholder in `[columns]`, and in
`wc_product_bot.py` set `payload["downloads"]` from that column in
`build_payload()`. Otherwise, remember to attach the right file to each
new product manually before publishing.

## Troubleshooting

**`403 Forbidden` on every request, even with a correct key:**
Some hosts and security plugins (Wordfence and similar) block requests
whose `Authorization` header or default `python-requests` User-Agent
look bot-like. This script already works around both by sending the key
as a query parameter and a normal browser User-Agent
(see the `[request]` section in your config). If you still get 403:

- Confirm the key's permission is **Read/Write**, not Read-only.
- Confirm the key's associated user has an Administrator or Shop
  Manager role.
- Check for a firewall/security plugin blocking `/wp-json/` for
  non-browser clients, or an IP allowlist that doesn't include wherever
  you're running this from.

**Permalinks set to "Plain" (`?p=123`):** the REST API generally needs a
different permalink structure (e.g. "Post name"). Settings → Permalinks.

**Description shows as one big raw-HTML block in the editor instead of
separate blocks:** your WordPress version may not auto-convert the
`<!-- wp:... -->` comments on import. The content and formatting on the
live page are unaffected either way — it only affects how it looks if
you reopen it in the block editor.

## License

MIT — see [LICENSE](LICENSE). Contributions and adaptations for other
use cases welcome.
