export interface ParsedContent {
    name: string;
    content: unknown;
}
export interface ParsedSlide extends ParsedContent {
    content: {
        id: string;
        type: string;
        text: string[];
    }[];
}
/**
 * @property content is base64 encoded
 */
export interface ParsedMedia extends ParsedContent {
    content: string;
}
export interface ParsedNote extends ParsedContent {
    content: string;
}
export interface ParsedPptx {
    notes: ParsedNote[];
    media: ParsedMedia[];
    slides: ParsedSlide[];
}
/**
 * Extract text (slides + notes) and images from a .pptx file.
 * @param filePath path to the .pptx file on disk
 * @returns Promise of ParsedPowerPoint with notes, media and slides
 */
export declare function extractPptx(filePath: string): Promise<ParsedPptx>;
export declare function extractPptxSlides(filePath: string): Promise<ParsedSlide[]>;
export declare function extractPptxMedia(filePath: string): Promise<ParsedMedia[]>;
export declare function extractPptxNotes(filePath: string): Promise<ParsedNote[]>;
//# sourceMappingURL=main.d.ts.map