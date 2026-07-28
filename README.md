# pptx-to-html

Convert PowerPoint (`.pptx`) slides into self-contained HTML files (SVG +
`foreignObject` text), with embedded fonts, extracted images/media, and an
`info.json` describing the deck.

Node.js port of the C# `PptxToHtml` tool; the output is byte-for-byte identical
to the reference implementation (apart from the `exportedAt` timestamp).

## Install

```bash
npm install pptx-to-html
```

Requires Node.js >= 18.

## Usage

ESM:

```js
import { exportPptx, getSlidesCount } from 'pptx-to-html';

exportPptx('deck.pptx', 'out/deck');
console.log(getSlidesCount('deck.pptx'));
```

CommonJS:

```js
const { exportPptx, getSlidesCount } = require('pptx-to-html');

exportPptx('deck.pptx', 'out/deck', null /* or a path to eot2ttf */);
```

TypeScript (declarations ship with the package — nothing extra to install):

```ts
import { exportPptx, getSlidesCount, getVersion } from 'pptx-to-html';

const slides: number = getSlidesCount('deck.pptx');
exportPptx('deck.pptx', 'out/deck');
```

## CLI

```bash
npx pptx-to-html --file <input.pptx> [--out <output-folder>] [--eot2ttf <path>]
```

- `--file` — path to the `.pptx` file (required)
- `--out` — destination folder (default: `out/<deck-name>`)
- `--eot2ttf` — path to the `eot2ttf` helper, used to convert older EOT-format
  embedded fonts to TTF (optional)

## API

| Export | Description |
| --- | --- |
| `exportPptx(pptxPath, outputDirectory, eot2ttfPath?)` | Convert a deck; creates the output folder if missing. |
| `getSlidesCount(pptxPath)` | Number of slides, without converting. |
| `getVersion()` | Tool version, as written into `info.json`. |
| `PptxConverter` | The converter class, for driving a run directly. |

Each run writes `slide1.html`, `slide2.html`, … plus per-animation-stage files
(`slide1_1.html`, …) and `info.json`, along with `fonts/`, `images/` and
`media/` asset folders. Every shape group carries a deterministic
`data-id="<slideId>-<n>"` handle, unique within its slide.

Supported: full text formatting with placeholder inheritance (slide → layout →
master → theme) and per-font metrics-derived line height; preset and custom
shape geometry with gradient/pattern/blip fills, strokes, shadow/glow and
scene3d; pictures with clip paths; embedded video and audio; tables, bar charts,
connectors and groups; slide transitions (including morph); entrance/exit
animations as per-stage HTML plus CSS keyframes; and embedded + system font
extraction with `@font-face` embedding.

## License

MIT
