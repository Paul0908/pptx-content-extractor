# PPTX Content Extractor

**PPTX Content Extractor** is a Node.js library for extracting slides, notes, and media content (e.g., images) from `.pptx` files. This tool leverages `JSZip` for unpacking `.pptx` archives and `xml2js` for parsing XML-based content.

## Features

- Extract text content from PowerPoint slides (`.pptx`).
- Retrieve media files (e.g., images) embedded in the presentation.
- Extract speaker notes for each slide.
- Modular structure for extracting specific content types (slides, media, or notes).

---

## Installation

Install the library via npm:

```bash
npm install --save pptx-content-extractor
```

## Usage

### Full Extraction

Extract all slides, media, and notes from a `.pptx` file:

```typescript
import { extractPptx } from 'pptx-content-extractor';

(async () => {
  const result = await extractPptx('/path/to/presentation.pptx');
  console.log('Slides:', result.slides);
  console.log('Media:', result.media);
  console.log('Notes:', result.notes);
})();
```

---

### Extract specific content

#### Slides

```typescript
import { extractPptxSlides } from 'pptx-content-extractor';

(async () => {
  const slides = await extractPptxSlides('/path/to/presentation.pptx');
  console.log('Slides:', slides);
})();
```

---

#### Media

```typescript
import { extractPptxMedia } from 'pptx-content-extractor';

(async () => {
  const media = await extractPptxMedia('/path/to/presentation.pptx');
  console.log('Media:', media);
})();
```

---

#### Notes

```typescript
import { extractPptxNotes } from 'pptx-content-extractor';

(async () => {
  const notes = await extractPptxNotes('/path/to/presentation.pptx');
  console.log('Notes:', notes);
})();
```

---

## API

### `extractPptx(filePath: string): Promise<ParsedPowerPoint>`

Extracts slides, media, and notes from a `.pptx` file.

- **`filePath`**: Path to the `.pptx` file.
- **Returns**: A `Promise<ParsedPowerPoint>` containing:
  - `slides`: An array of parsed slides.
  - `media`: An array of media content.
  - `notes`: An array of parsed notes.

---

### `extractPptxSlides(filePath: string): Promise<ParsedSlide[]>`

Extracts only the slides.

- **`filePath`**: Path to the `.pptx` file.
- **Returns**: A `Promise<ParsedSlide[]>` containing parsed slides.

---

### `extractPptxMedia(filePath: string): Promise<ParsedMedia[]>`

Extracts only the media content.

- **`filePath`**: Path to the `.pptx` file.
- **Returns**: A `Promise<ParsedMedia[]>` containing media content.

---

### `extractPptxNotes(filePath: string): Promise<ParsedNote[]>`

Extracts only the notes.

- **`filePath`**: Path to the `.pptx` file.
- **Returns**: A `Promise<ParsedNote[]>` containing parsed notes.

---

## Types

### `ParsedContent`

Base interface for parsed content.

```typescript
export interface ParsedContent {
  name: string;
  content: unknown;
}
```

---

### `ParsedPptx`

```typescript
export interface ParsedPowerPoint {
  slides: ParsedSlide[];
  media: ParsedMedia[];
  notes: ParsedNote[];
}
```

---

### `ParsedSlide`

```typescript
export interface ParsedSlide extends ParsedContent {
  content: { id: string; type: string; text: string[] }[];
}
```

---

### `ParsedMedia`

```typescript
export interface ParsedMedia extends ParsedContent {
  content: string; // Base64-encoded media content
}
```

---

### `ParsedNote`

```typescript
export interface ParsedNote extends ParsedContent {
  content: string;
}
```

---
