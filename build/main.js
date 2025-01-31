"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.extractPptx = extractPptx;
exports.extractPptxSlides = extractPptxSlides;
exports.extractPptxMedia = extractPptxMedia;
exports.extractPptxNotes = extractPptxNotes;
const fs_1 = __importDefault(require("fs"));
const jszip_1 = __importDefault(require("jszip"));
const xml2js_1 = require("xml2js");
/**
 * Extract text (slides + notes) and images from a .pptx file.
 * @param filePath path to the .pptx file on disk
 * @returns Promise of ParsedPowerPoint with notes, media and slides
 */
function extractPptx(filePath) {
    return __awaiter(this, void 0, void 0, function* () {
        const files = yield loadFilesOfFile(filePath);
        const { slides: rawSlides, media: rawMedia, notes: rawNotes } = extractParts(files);
        const slides = yield parsePart(rawSlides, parseSlideContent);
        const media = yield parsePart(rawMedia, parseMediaContent);
        const notes = yield parsePart(rawNotes, parseNotesContent);
        return {
            slides,
            media,
            notes
        };
    });
}
function extractPptxSlides(filePath) {
    return __awaiter(this, void 0, void 0, function* () {
        const files = yield loadFilesOfFile(filePath);
        const rawSlides = getSlides(files);
        return yield parsePart(rawSlides, parseSlideContent);
    });
}
function extractPptxMedia(filePath) {
    return __awaiter(this, void 0, void 0, function* () {
        const files = yield loadFilesOfFile(filePath);
        const rawSlides = getMedia(files);
        return yield parsePart(rawSlides, parseMediaContent);
    });
}
function extractPptxNotes(filePath) {
    return __awaiter(this, void 0, void 0, function* () {
        const files = yield loadFilesOfFile(filePath);
        const rawSlides = getNotes(files);
        return yield parsePart(rawSlides, parseNotesContent);
    });
}
function loadFilesOfFile(filePath) {
    return __awaiter(this, void 0, void 0, function* () {
        const fileBuffer = readFileAsBuffer(filePath);
        return (yield loadPpt(fileBuffer)).files;
    });
}
function readFileAsBuffer(filePath) {
    const fileBuffer = fs_1.default.readFileSync(filePath);
    if (!fileBuffer) {
        throw new Error("Failed to read file");
    }
    return fileBuffer;
}
function loadPpt(fileBuffer) {
    return __awaiter(this, void 0, void 0, function* () {
        return jszip_1.default.loadAsync(fileBuffer).catch((e) => {
            console.error(e);
            throw new Error("Failed to load .pptx file");
        });
    });
}
function parsePart(toParse, parser) {
    return __awaiter(this, void 0, void 0, function* () {
        return yield Promise.all(toParse.map((part) => __awaiter(this, void 0, void 0, function* () { return yield parser(part); })));
    });
}
function parseNotesContent(note) {
    return __awaiter(this, void 0, void 0, function* () {
        const content = yield note.async('string');
        // TODO
        return { name: note.name, content };
    });
}
function parseMediaContent(media) {
    return __awaiter(this, void 0, void 0, function* () {
        const binaries = yield media.async('base64');
        const fileName = media.name.split('/').pop() || media.name;
        const mediaType = fileName.split('.').pop() || 'unknown';
        return {
            name: media.name,
            content: `data:image/${mediaType};base64,${binaries}`
        };
    });
}
function parseSlideContent(slide) {
    return __awaiter(this, void 0, void 0, function* () {
        var _a, _b, _c, _d, _e;
        const xml = yield slide.async('string');
        const parsed = yield (0, xml2js_1.parseStringPromise)(xml);
        const results = [];
        const shapes = (_e = (_d = (_c = (_b = (_a = parsed['p:sld']) === null || _a === void 0 ? void 0 : _a['p:cSld']) === null || _b === void 0 ? void 0 : _b[0]) === null || _c === void 0 ? void 0 : _c['p:spTree']) === null || _d === void 0 ? void 0 : _d[0]) === null || _e === void 0 ? void 0 : _e['p:sp'];
        if (shapes) {
            shapes.forEach((shape) => {
                var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p;
                const cNvPr = (_d = (_c = (_b = (_a = shape['p:nvSpPr']) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b['p:cNvPr']) === null || _c === void 0 ? void 0 : _c[0]) === null || _d === void 0 ? void 0 : _d['$'];
                const phType = ((_l = (_k = (_j = (_h = (_g = (_f = (_e = shape['p:nvSpPr']) === null || _e === void 0 ? void 0 : _e[0]) === null || _f === void 0 ? void 0 : _f['p:nvPr']) === null || _g === void 0 ? void 0 : _g[0]) === null || _h === void 0 ? void 0 : _h['p:ph']) === null || _j === void 0 ? void 0 : _j[0]) === null || _k === void 0 ? void 0 : _k['$']) === null || _l === void 0 ? void 0 : _l['type']) || 'unknown';
                const texts = ((_p = (_o = (_m = shape['p:txBody']) === null || _m === void 0 ? void 0 : _m[0]) === null || _o === void 0 ? void 0 : _o['a:p']) === null || _p === void 0 ? void 0 : _p.map((paragraph) => {
                    var _a;
                    return ((_a = paragraph['a:r']) === null || _a === void 0 ? void 0 : _a.map(run => { var _a; return (_a = run['a:t']) === null || _a === void 0 ? void 0 : _a[0]; }).join(' ')) || '';
                }).filter((text) => text)) || [];
                if (cNvPr && texts.length > 0) {
                    results.push({
                        id: cNvPr.id,
                        type: phType,
                        text: texts,
                    });
                }
            });
        }
        return {
            name: slide.name,
            content: results
        };
    });
}
function extractNumberFromName(fileName, pattern) {
    const match = fileName.match(pattern);
    return match ? parseInt(match[1], 10) : Number.MAX_SAFE_INTEGER;
}
function getPartByBasePathAndPattern(files, basePath, pattern) {
    const partObjects = Object.keys(files)
        .filter((fileName) => fileName.startsWith(basePath))
        .map((fileName) => files[fileName]);
    partObjects.sort((a, b) => {
        const aNum = extractNumberFromName(a.name, pattern);
        const bNum = extractNumberFromName(b.name, pattern);
        return aNum - bNum;
    });
    return partObjects;
}
function getSlides(files) {
    const slidesBasePath = "ppt/slides/";
    const slidePattern = /slide(\d+)\.xml(\.rels)?$/;
    return getPartByBasePathAndPattern(files, slidesBasePath, slidePattern);
}
function getMedia(files) {
    const mediaBasePath = "ppt/media/";
    const mediaPattern = /(\d+)\.(jpg|jpeg|png|gif)$/;
    return getPartByBasePathAndPattern(files, mediaBasePath, mediaPattern);
}
function getNotes(files) {
    const notesBasePath = "ppt/notesSlides/";
    const notesPattern = /notesSlide(\d+)\.xml(\.rels)?$/;
    return getPartByBasePathAndPattern(files, notesBasePath, notesPattern);
}
function extractParts(files) {
    return {
        slides: getSlides(files),
        media: getMedia(files),
        notes: getNotes(files),
    };
}
