# Catalog artwork

- `../banner.png`: the 1200×600 (2:1) catalog card and plugin-page hero (`image:` in the catalog entry).
- `../icon.png` / `../icon.svg`: the 512×512 Office Desk mark.

Both are rendered from these sources with headless Chrome, using only the neutral demo screenshots:

```bash
cd docs/media/src
google-chrome --headless=new --hide-scrollbars --window-size=1200,600 --screenshot=../banner.png "file://$PWD/banner.html"
```
