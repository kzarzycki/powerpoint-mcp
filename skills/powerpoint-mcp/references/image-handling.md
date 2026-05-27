# Image & Logo Handling

Rules for inserting images and logos into presentations via the MCP bridge.

## Aspect Ratio

**NEVER set both `width` and `height` on `insert_image`** — this stretches the image. Set only ONE dimension (usually `height` for logo rows, `width` for hero images) and let the other auto-calculate from the source image's native aspect ratio.

## Transparency

Always verify inserted images have transparent backgrounds. A white-background logo on an off-white (`#FBF7F7`) or colored strip is immediately visible and looks unprofessional.

**If only a white-bg version is available**, remove the background before inserting:

```python
from PIL import Image

img = Image.open("logo.png").convert("RGBA")
data = img.getdata()
new_data = []
for item in data:
    if item[0] > 240 and item[1] > 240 and item[2] > 240:
        new_data.append((255, 255, 255, 0))
    else:
        new_data.append(item)
img.putdata(new_data)
img.save("logo-transparent.png")
```

## Visual Sizing

Size logos by **perceived visual weight**, not bounding box dimensions. Some logos have generous padding, making same-height logos look very different in size.

After inserting all logos in a row:
1. `inspect_slide` to read each logo's actual width
2. Calculate even spacing based on actual dimensions
3. `execute_officejs` to reposition for balanced distribution

## Contextual Placement

Place logos where they semantically belong:
- **Partner logos** → with partnership credentials (credibility box)
- **Tech logos** → with technical capability descriptions
- **Client logos** → with case studies or references

Don't default to a disconnected generic row at the slide bottom.

## Sourcing

### fetch-logo.sh (automated, preferred)

Script at `scripts/fetch-logo.sh` in the project directory. Searches by **name** (not domain):

```bash
# Tech logos (gilbarbara — 1400+ SVGs, includes 60+ AWS services)
./scripts/fetch-logo.sh Kafka Kubernetes "AWS MSK" Terraform

# Data/AI/ML logos (LF AI Landscape — 470+ SVGs)
./scripts/fetch-logo.sh "Apache Iceberg" "Delta Lake" Dagster Trino MLflow

# Company/vendor logos (Brandfetch — PNG, needs BRANDFETCH_CLIENT_ID env var)
./scripts/fetch-logo.sh Acme Globex Initech Umbrella
./scripts/fetch-logo.sh --source bf "dbt Labs"  # force Brandfetch for company version

# Options
./scripts/fetch-logo.sh --icon dbt              # symbol-only (no text)
./scripts/fetch-logo.sh --search-only Flink     # browse matches without downloading
./scripts/fetch-logo.sh --outdir /tmp/logos Kafka  # custom output directory
```

Output goes to `logos/` in the project directory. Naming: `{slug}.svg` (gilbarbara/LF AI) or `{domain}-logo.png` (Brandfetch).

Sources are tried in order: gilbarbara → LF AI → Brandfetch. First match wins.

**Known gaps** (not in any automated source — download manually from cloud architecture icon packs):
- GCP sub-services: BigQuery, Dataflow, Pub/Sub, Dataproc
- Azure sub-services: Data Factory, Synapse, Purview

### Manual sourcing (fallback)

If `fetch-logo.sh` doesn't find the logo:
1. Official press/brand pages (highest quality, correct colors)
2. GitHub repos — many OSS projects have brand assets in `/assets` or `/branding`
3. Fallback aggregators: logo.wine, stickpng, svgcdn, brandpnglogo
4. Wikimedia Commons (last resort, may be outdated)

Prefer **SVG** (vector, scales perfectly) or **PNG with transparent background**. Both work in Office.js. Save downloaded logos to `logos/` in the project dir for reuse across slides.

## Workflow

Before adding visual elements to a slide:
1. **Inspect** — `screenshot_slide` to see current state
2. **Propose** — offer the user visual options (style treatment, count, labels, placement)
3. **Confirm** — get user preference
4. **Execute** — build per their choice

Don't jump straight to execution mechanics.
