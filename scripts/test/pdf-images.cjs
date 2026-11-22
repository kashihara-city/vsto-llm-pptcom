// How the task pane takes the images of an attached PDF (pptcom/UI/web/pdfdoc.js): at most 60 are kept, and none is
// made into a PNG once the document has that many - a page's images were all made first, and those over the number
// thrown away; an image that comes twice counts once; one whose long side is over 4000 pixels is taken smaller, its
// pixels averaged - each by how much of it shows - without anything of its own size being made on the way.
// pdfdoc.js runs as it is but for its first lines: it is a module that imports PDF.js, so its import is replaced by
// a stand-in for PDF.js (a document of pages with images painted on them), and the browser's canvas by one that
// counts. No packages.
//
//   node scripts/test/pdf-images.cjs     exit code 0 = all passed
"use strict";
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");

let source = fs.readFileSync(path.join(__dirname, "..", "..", "pptcom", "UI", "web", "pdfdoc.js"), "utf8");
let passed = 0, failed = 0;

function check(name, ok, detail) {
  if (ok) { passed++; return; }
  failed++;
  console.log("FAIL " + name);
  console.log("     " + detail);
}

// the module's own words, as a script: each replacement must find its place, or the file has changed
function replace(what, by, times) {
  const parts = source.split(what);
  if (parts.length - 1 !== times) throw new Error(`expected ${times} of ${JSON.stringify(what)} in pdfdoc.js, found ${parts.length - 1}`);
  source = parts.join(by);
}
replace('import * as pdfjs from "./vendor/pdfjs/pdf.mjs";', "const pdfjs = __pdfjs;", 1);
replace("import.meta.url", '"https://pptcom.invalid/pdfdoc.js"', 2);
replace("export async function", "async function", 2);

const OPS = { save: 1, restore: 2, transform: 3, paintImageXObject: 4 };

// A page that paints these images (each {id, width, height} and `bitmap` or `data` with `kind`), one after the other.
function page(images) {
  const byId = new Map(images.map(i => [i.id, i]));
  return {
    view: [0, 0, 600, 800],
    getTextContent: async () => ({ items: [] }),
    getOperatorList: async () => ({ fnArray: images.map(() => OPS.paintImageXObject), argsArray: images.map(i => [i.id]) }),
    objs: { get: (id, done) => done(byId.get(id)) },
    commonObjs: { get: (id, done) => done(byId.get(id)) },
    cleanup() {},
  };
}

// pdfdoc.js with a document of these pages. `made`: the size of each PNG made, in order.
async function convert(pages) {
  const made = [];
  class Canvas {
    constructor(width, height) { this.width = width; this.height = height; this.painted = ""; }
    getContext() {
      return {
        fillRect() {},
        drawImage: (bitmap, x, y, w, h) => { this.painted = `bitmap ${bitmap.id} at ${w}x${h}`; },
        putImageData: data => { this.painted = `pixels ${Array.from(data.data.subarray(0, 16)).join(",")} of ${data.width}x${data.height}`; },
      };
    }
    async convertToBlob() {
      made.push(`${this.width}x${this.height}`);
      const bytes = Buffer.from(`${this.width}x${this.height} ${this.painted}`);
      return { arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) };
    }
  }
  const context = vm.createContext({
    console, URL, Uint8Array, Uint8ClampedArray, Set, Promise, Math, Array, setTimeout,
    crypto: require("node:crypto").webcrypto,
    OffscreenCanvas: Canvas,
    ImageData: class { constructor(data, width, height) { this.data = data; this.width = width; this.height = height; } },
    __pdfjs: {
      GlobalWorkerOptions: {}, OPS,
      getDocument: () => ({ promise: Promise.resolve({
        numPages: pages.length,
        getPage: async n => pages[n - 1],
        loadingTask: { destroy: async () => {} },
      }) }),
    },
  });
  vm.runInContext(source, context);
  const result = await vm.runInContext("convertPdf", context)(new Uint8Array([1]));
  return { result, made, context };
}

const bitmaps = (pageNumber, count, width = 200, height = 200) =>
  Array.from({ length: count }, (_, i) => ({ id: `img_p${pageNumber}_${i}`, width, height, bitmap: { id: `p${pageNumber}_${i}` } }));

(async () => {
  // three pages of 25 images: the third page's first ten fill the document's sixty
  const many = await convert([page(bitmaps(1, 25)), page(bitmaps(2, 25)), page(bitmaps(3, 25))]);
  check("a document keeps 60 images, and makes no PNG of an image after its sixtieth",
        many.result.ok && many.result.images.length === 60 && many.made.length === 60,
        `${many.result.images && many.result.images.length} kept, ${many.made.length} PNGs made`);

  // the same image on two pages counts once, and the ones after it are still taken
  const twice = await convert([page([{ id: "a", width: 200, height: 200, bitmap: { id: "same" } }]),
                               page([{ id: "b", width: 200, height: 200, bitmap: { id: "same" } }, ...bitmaps(2, 2)])]);
  check("an image that comes twice is kept once, and the images after it are taken",
        twice.result.images.length === 3 && twice.made.length === 4, `${twice.result.images.length} kept, ${twice.made.length} PNGs made`);

  // too small to be a picture: left out without a PNG
  const small = await convert([page([{ id: "icon", width: 32, height: 32, bitmap: { id: "icon" } }, ...bitmaps(1, 1)])]);
  check("an image too small to be a picture is left out, and no PNG is made of it", small.result.images.length === 1 && small.made.join() === "200x200", small.made.join());

  // long side over 4000: taken smaller, both ways PDF.js hands an image over
  const rgb = new Uint8Array(8000 * 64 * 3).fill(200);
  const large = await convert([page([{ id: "wide", width: 8000, height: 2000, bitmap: { id: "wide" } },
                                     { id: "raw", width: 8000, height: 64, kind: 2, data: rgb },
                                     { id: "fits", width: 4000, height: 3000, bitmap: { id: "fits" } }])]);
  check("an image whose long side is over 4000 pixels is taken smaller, to 4000; one that is not keeps its size",
        large.made.join() === "4000x1000,4000x32,4000x3000", large.made.join());

  // the averaging itself: each pixel of the smaller picture is the mean of the pixels it stands for
  const pixels = vm.runInContext("pixels", large.context);
  const rgbHalf = pixels({ width: 4, height: 2, kind: 2, data: new Uint8Array([0, 0, 0, 100, 100, 100, 10, 20, 30, 30, 40, 50,
                                                                              0, 0, 0, 100, 100, 100, 10, 20, 30, 30, 40, 50]) }, 2, 1);
  const rgbaSame = pixels({ width: 2, height: 1, kind: 3, data: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]) }, 2, 1);
  const bits = pixels({ width: 4, height: 1, kind: 1, data: new Uint8Array([0b10100000]) }, 4, 1);
  const bitsHalf = pixels({ width: 4, height: 1, kind: 1, data: new Uint8Array([0b11100000]) }, 2, 1);
  check("raw pixels, made smaller, are averaged: RGB halved in width", Array.from(rgbHalf).join() === "50,50,50,255,20,30,40,255", Array.from(rgbHalf).join());
  check("raw pixels at their own size stay as they are: RGBA, and one bit a pixel",
        Array.from(rgbaSame).join() === "1,2,3,4,5,6,7,8" && Array.from(bits).join() === "255,255,255,255,0,0,0,255,255,255,255,255,0,0,0,255",
        Array.from(rgbaSame).join() + " / " + Array.from(bits).join());
  check("one bit a pixel, halved: white and black average to gray", Array.from(bitsHalf).join() === "255,255,255,255,128,128,128,255", Array.from(bitsHalf).join());
  check("a kind of raw pixels that is none of the three gives nothing", pixels({ width: 1, height: 1, kind: 9, data: new Uint8Array(4) }, 1, 1) === null, "");

  // RGBA made smaller: a transparent pixel's color is not mixed into its neighbour's - red beside nothing stays red,
  // half as much of it shows; two colors that both show are averaged as before; nothing but transparent stays so
  const edge = pixels({ width: 2, height: 1, kind: 3, data: new Uint8Array([255, 0, 0, 255, 0, 0, 0, 0]) }, 1, 1);
  const both = pixels({ width: 2, height: 1, kind: 3, data: new Uint8Array([10, 20, 30, 255, 30, 40, 50, 255]) }, 1, 1);
  const faint = pixels({ width: 2, height: 1, kind: 3, data: new Uint8Array([200, 0, 0, 51, 0, 100, 0, 153]) }, 1, 1);
  const none = pixels({ width: 2, height: 1, kind: 3, data: new Uint8Array([9, 9, 9, 0, 7, 7, 7, 0]) }, 1, 1);
  check("RGBA made smaller: a pixel's color counts by how much of it shows, so a transparent neighbour does not darken an edge",
        Array.from(edge).join() === "255,0,0,128" && Array.from(both).join() === "20,30,40,255" && Array.from(faint).join() === "50,75,0,102" &&
        Array.from(none).join() === "0,0,0,0",
        [edge, both, faint, none].map(p => Array.from(p).join()).join(" / "));

  console.log();
  console.log(failed === 0 ? `ALL PASSED (${passed})` : `${failed} FAILED, ${passed} passed`);
  process.exitCode = failed === 0 ? 0 : 1;
})().catch(e => { console.error(e); process.exitCode = 1; });
