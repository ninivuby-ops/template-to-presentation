import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";

export type Placeholder = {
  id: string;
  slide: number;
  shape: string;
  kind: string;
  current: string;
  /** Sibling text on the same slide (headings, prescribed pointers) for context. */
  context: string[];
};

const PLACEHOLDER_RE =
  /(\{\{[^}]*\}\})|(\[[^\]]{1,80}\])|(<[^>]{1,80}>)|lorem ipsum|click to add|click to edit|placeholder|^x{3,}$|^_{3,}$|^tbd$|^lorem/i;

const xmlEscape = (s: string) =>
  s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

const xmlUnescape = (s: string) =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

const paraText = (p: string) =>
  xmlUnescape(
    Array.from(p.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g))
      .map((m) => m[1])
      .join(""),
  ).trim();

function matchBlocks(xml: string, tag: string) {
  const re = new RegExp(`<${tag}[ >][\\s\\S]*?</${tag}>`, "g");
  const out: { text: string; start: number; end: number }[] = [];
  for (const m of xml.matchAll(re)) {
    out.push({ text: m[0], start: m.index!, end: m.index! + m[0].length });
  }
  return out;
}

function slideEntries(files: Record<string, Uint8Array>) {
  return Object.keys(files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort(
      (a, b) =>
        Number(a.match(/(\d+)/)![1]) - Number(b.match(/(\d+)/)![1]),
    );
}

/** Extract every fillable placeholder without touching structure or styling. */
export function extractPlaceholders(bytes: Uint8Array) {
  const files = unzipSync(bytes);
  const slides = slideEntries(files);
  if (!slides.length) throw new Error("No slides found — is this a valid .pptx template?");

  const placeholders: Placeholder[] = [];

  slides.forEach((name, si) => {
    const xml = strFromU8(files[name]!);
    const shapes = matchBlocks(xml, "p:sp");
    const slideTexts = shapes
      .map((s) => matchBlocks(s.text, "a:p").map((p) => paraText(p.text)).join(" \n"))
      .map((t) => t.trim())
      .filter(Boolean);

    shapes.forEach((sp, shi) => {
      const shapeName = sp.text.match(/<p:cNvPr[^>]*name="([^"]*)"/)?.[1] ?? `Shape ${shi + 1}`;
      const phType = sp.text.match(/<p:ph[^>]*type="([^"]*)"/)?.[1] ?? "body";
      const isTitle = /title|ctrTitle/i.test(phType) || /title/i.test(shapeName);
      const paras = matchBlocks(sp.text, "a:p");

      paras.forEach((p, pi) => {
        const text = paraText(p.text);
        const empty = text.length === 0;
        const looksPlaceholder = !!text && PLACEHOLDER_RE.test(text);
        // Empty title paragraphs and headings are never rewritten.
        if (!looksPlaceholder && !(empty && !isTitle && /<p:ph/.test(sp.text))) return;

        placeholders.push({
          id: `s${si + 1}_sh${shi}_p${pi}`,
          slide: si + 1,
          shape: shapeName,
          kind: isTitle ? "title" : phType,
          current: text,
          context: slideTexts,
        });
      });
    });
  });

  return { placeholders, slideCount: slides.length };
}

/** Write generated text back into exactly the paragraphs we identified. */
export function fillTemplate(bytes: Uint8Array, values: Record<string, string>) {
  const files = unzipSync(bytes);
  const slides = slideEntries(files);
  let filled = 0;

  slides.forEach((name, si) => {
    let xml = strFromU8(files[name]!);
    const shapes = matchBlocks(xml, "p:sp");
    // Rebuild from the end so earlier offsets stay valid.
    for (let shi = shapes.length - 1; shi >= 0; shi--) {
      const sp = shapes[shi]!;
      let spXml = sp.text;
      const paras = matchBlocks(spXml, "a:p");
      for (let pi = paras.length - 1; pi >= 0; pi--) {
        const value = values[`s${si + 1}_sh${shi}_p${pi}`];
        if (!value) continue;
        const p = paras[pi]!;
        const replaced = setParagraphText(p.text, value);
        if (replaced === p.text) continue;
        spXml = spXml.slice(0, p.start) + replaced + spXml.slice(p.end);
        filled++;
      }
      if (spXml !== sp.text) xml = xml.slice(0, sp.start) + spXml + xml.slice(sp.end);
    }
    files[name] = strToU8(xml);
  });

  return { bytes: zipSync(files, { level: 6 }), filled };
}

/** Keeps run properties (font, size, color) of the template intact. */
function setParagraphText(para: string, value: string) {
  const text = xmlEscape(value);
  const runs = matchBlocks(para, "a:r");

  if (runs.length) {
    let out = para;
    for (let i = runs.length - 1; i >= 1; i--) {
      out = out.slice(0, runs[i]!.start) + out.slice(runs[i]!.end);
    }
    const first = matchBlocks(out, "a:r")[0]!;
    const newRun = first.text.replace(
      /<a:t>[\s\S]*?<\/a:t>/,
      `<a:t>${text}</a:t>`,
    );
    return out.slice(0, first.start) + newRun + out.slice(first.end);
  }

  // Empty paragraph: build a run from its end-run properties so styling matches.
  const endRpr = para.match(/<a:endParaRPr([^>]*)\/>/);
  const attrs = endRpr?.[1] ?? ' lang="en-US" dirty="0"';
  const run = `<a:r><a:rPr${attrs}/><a:t>${text}</a:t></a:r>`;
  if (endRpr) return para.replace(endRpr[0], run + endRpr[0]);
  return para.replace(/<\/a:p>$/, `${run}</a:p>`);
}

export type MediaItem = {
  /** zip path, e.g. ppt/media/image2.png */
  id: string;
  name: string;
  kind: "image" | "video" | "audio" | "other";
  slides: number[];
  bytes: number;
  /** small previews only, as a data URL */
  preview?: string | undefined;
};

const MIME: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  webp: "image/webp",
  emf: "image/emf",
  wmf: "image/wmf",
  mp4: "video/mp4",
  mov: "video/quicktime",
  m4v: "video/x-m4v",
  avi: "video/x-msvideo",
  wmv: "video/x-ms-wmv",
  mp3: "audio/mpeg",
  wav: "audio/wav",
  m4a: "audio/mp4",
};

const kindOf = (ext: string): MediaItem["kind"] => {
  const m = MIME[ext] ?? "";
  if (m.startsWith("image/")) return "image";
  if (m.startsWith("video/")) return "video";
  if (m.startsWith("audio/")) return "audio";
  return "other";
};

const toB64 = (bytes: Uint8Array) => {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk)
    bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(bin);
};

/** List every image/video/audio asset the slides reference. */
export function extractMedia(bytes: Uint8Array): MediaItem[] {
  const files = unzipSync(bytes);
  const slides = slideEntries(files);
  const usedBy = new Map<string, number[]>();

  slides.forEach((name, si) => {
    const relName = name.replace(/slides\/(slide\d+)\.xml$/, "slides/_rels/$1.xml.rels");
    const rel = files[relName];
    if (!rel) return;
    const xml = strFromU8(rel);
    for (const m of xml.matchAll(/Target="\.\.\/(media\/[^"]+)"/g)) {
      const path = `ppt/${m[1]}`;
      const list = usedBy.get(path) ?? [];
      if (!list.includes(si + 1)) list.push(si + 1);
      usedBy.set(path, list);
    }
  });

  const out: MediaItem[] = [];
  for (const path of Object.keys(files).filter((n) => n.startsWith("ppt/media/"))) {
    const data = files[path]!;
    const name = path.split("/").pop()!;
    const ext = (name.split(".").pop() ?? "").toLowerCase();
    const kind = kindOf(ext);
    out.push({
      id: path,
      name,
      kind,
      slides: usedBy.get(path) ?? [],
      bytes: data.length,
      preview:
        kind === "image" && data.length < 1_500_000 && MIME[ext]
          ? `data:${MIME[ext]};base64,${toB64(data)}`
          : undefined,
    });
  }
  return out.sort((a, b) => (a.slides[0] ?? 99) - (b.slides[0] ?? 99));
}

/** Fill text and optionally swap media assets in place, keeping every zip path. */
export function buildDeck(
  bytes: Uint8Array,
  values: Record<string, string>,
  media: Record<string, Uint8Array> = {},
) {
  const { bytes: filledBytes, filled } = fillTemplate(bytes, values);
  const mediaKeys = Object.keys(media);
  if (!mediaKeys.length) return { bytes: filledBytes, filled, replacedMedia: 0 };

  const files = unzipSync(filledBytes);
  let replacedMedia = 0;
  for (const key of mediaKeys) {
    if (!files[key]) continue;
    files[key] = new Uint8Array(media[key]!);
    replacedMedia++;
  }
  return { bytes: new Uint8Array(zipSync(files, { level: 6 })), filled, replacedMedia };
}

/* ---------------------------------------------------------------------------
 * Complex slide objects: tables, charts and SmartArt diagrams.
 * All of these already exist in the template — we only rewrite their text and
 * cached numbers, never their styling, size, position or structure.
 * ------------------------------------------------------------------------ */

export type TableStruct = {
  id: string;
  slide: number;
  rows: string[][];
};

export type ChartStruct = {
  id: string;
  slide: number;
  kind: string;
  title: string;
  categories: string[];
  series: { name: string; values: number[] }[];
};

export type DiagramStruct = {
  id: string;
  slide: number;
  nodes: string[];
};

const cellText = (tc: string) =>
  matchBlocks(tc, "a:p").map((p) => paraText(p.text)).join(" ").trim();

/** Read every table on every slide. */
export function extractTables(files: Record<string, Uint8Array>): TableStruct[] {
  const out: TableStruct[] = [];
  slideEntries(files).forEach((name, si) => {
    const xml = strFromU8(files[name]!);
    matchBlocks(xml, "a:tbl").forEach((tbl, ti) => {
      const rows = matchBlocks(tbl.text, "a:tr").map((tr) =>
        matchBlocks(tr.text, "a:tc").map((tc) => cellText(tc.text)),
      );
      if (rows.length) out.push({ id: `tbl_s${si + 1}_${ti}`, slide: si + 1, rows });
    });
  });
  return out;
}

function chartTargets(files: Record<string, Uint8Array>) {
  const map = new Map<string, number>();
  slideEntries(files).forEach((name, si) => {
    const rel = files[name.replace(/slides\/(slide\d+)\.xml$/, "slides/_rels/$1.xml.rels")];
    if (!rel) return;
    const xml = strFromU8(rel);
    for (const m of xml.matchAll(/Target="[^"]*?(charts\/chart\d+\.xml)"/g)) {
      map.set(`ppt/${m[1]}`, si + 1);
    }
    // charts are often reached through a graphicFrame -> chart part directly
    for (const m of xml.matchAll(/Target="[^"]*?(drawings\/[^"]+\.xml)"/g)) void m;
  });
  return map;
}

const firstVal = (block: string) => xmlUnescape(block.match(/<c:v>([\s\S]*?)<\/c:v>/)?.[1] ?? "").trim();
const allVals = (block: string) =>
  Array.from(block.matchAll(/<c:v>([\s\S]*?)<\/c:v>/g)).map((m) => xmlUnescape(m[1]!).trim());

/** Read every chart's cached title, categories and series values. */
export function extractCharts(files: Record<string, Uint8Array>): ChartStruct[] {
  const slideOf = chartTargets(files);
  const out: ChartStruct[] = [];
  for (const path of Object.keys(files).filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    const xml = strFromU8(files[path]!);
    const kind = xml.match(/<c:(\w+Chart)[ >]/)?.[1]?.replace(/Chart$/, "") ?? "chart";
    const titleBlock = matchBlocks(xml, "c:title")[0]?.text ?? "";
    const title = xmlUnescape(
      Array.from(titleBlock.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)).map((m) => m[1]).join(""),
    ).trim();

    const sers = matchBlocks(xml, "c:ser");
    let categories: string[] = [];
    const series = sers.map((s) => {
      const tx = matchBlocks(s.text, "c:tx")[0]?.text ?? "";
      const cat = matchBlocks(s.text, "c:cat")[0]?.text ?? "";
      const val = matchBlocks(s.text, "c:val")[0]?.text ?? "";
      if (!categories.length) categories = allVals(cat);
      return {
        name: firstVal(tx),
        values: allVals(val).map((v) => Number(v) || 0),
      };
    });
    out.push({ id: path, slide: slideOf.get(path) ?? 0, kind, title, categories, series });
  }
  return out.sort((a, b) => a.slide - b.slide);
}

/** Read SmartArt / diagram node labels. */
export function extractDiagrams(files: Record<string, Uint8Array>): DiagramStruct[] {
  const slideOf = new Map<string, number>();
  slideEntries(files).forEach((name, si) => {
    const rel = files[name.replace(/slides\/(slide\d+)\.xml$/, "slides/_rels/$1.xml.rels")];
    if (!rel) return;
    for (const m of strFromU8(rel).matchAll(/Target="[^"]*?(diagrams\/data\d+\.xml)"/g)) {
      slideOf.set(`ppt/${m[1]}`, si + 1);
    }
  });

  const out: DiagramStruct[] = [];
  for (const path of Object.keys(files).filter((n) => /^ppt\/diagrams\/data\d+\.xml$/.test(n))) {
    const xml = strFromU8(files[path]!);
    const nodes = matchBlocks(xml, "a:p")
      .map((p) => paraText(p.text))
      .filter((t) => t.length > 0 || true);
    out.push({ id: path, slide: slideOf.get(path) ?? 0, nodes });
  }
  return out.sort((a, b) => a.slide - b.slide);
}

/** Write table cell text back, keeping every cell's own formatting. */
function fillTables(files: Record<string, Uint8Array>, tables: Record<string, string[][]>) {
  let filled = 0;
  slideEntries(files).forEach((name, si) => {
    let xml = strFromU8(files[name]!);
    let changed = false;
    const tbls = matchBlocks(xml, "a:tbl");
    for (let ti = tbls.length - 1; ti >= 0; ti--) {
      const wanted = tables[`tbl_s${si + 1}_${ti}`];
      if (!wanted) continue;
      const tbl = tbls[ti]!;
      let tblXml = tbl.text;
      const trs = matchBlocks(tblXml, "a:tr");
      for (let ri = trs.length - 1; ri >= 0; ri--) {
        const row = wanted[ri];
        if (!row) continue;
        let trXml = trs[ri]!.text;
        const tcs = matchBlocks(trXml, "a:tc");
        for (let ci = tcs.length - 1; ci >= 0; ci--) {
          const value = row[ci];
          if (value === undefined) continue;
          let tcXml = tcs[ci]!.text;
          const paras = matchBlocks(tcXml, "a:p");
          if (!paras.length) continue;
          // collapse to the first paragraph so a cell never grows extra lines
          for (let pi = paras.length - 1; pi >= 1; pi--) {
            tcXml = tcXml.slice(0, paras[pi]!.start) + tcXml.slice(paras[pi]!.end);
          }
          const first = matchBlocks(tcXml, "a:p")[0]!;
          const replaced = setParagraphText(first.text, value);
          tcXml = tcXml.slice(0, first.start) + replaced + tcXml.slice(first.end);
          trXml = trXml.slice(0, tcs[ci]!.start) + tcXml + trXml.slice(tcs[ci]!.end);
          filled++;
        }
        tblXml = tblXml.slice(0, trs[ri]!.start) + trXml + tblXml.slice(trs[ri]!.end);
      }
      xml = xml.slice(0, tbl.start) + tblXml + xml.slice(tbl.end);
      changed = true;
    }
    if (changed) files[name] = strToU8(xml);
  });
  return filled;
}

function replaceVals(block: string, values: string[]) {
  let i = 0;
  return block.replace(/<c:v>[\s\S]*?<\/c:v>/g, (m) => {
    const next = values[i++];
    return next === undefined ? m : `<c:v>${xmlEscape(next)}</c:v>`;
  });
}

/** Write chart titles, categories and cached values back into the chart part. */
function fillCharts(
  files: Record<string, Uint8Array>,
  charts: Record<string, { title?: string; categories?: string[]; series?: { name?: string; values?: number[] }[] }>,
) {
  let filled = 0;
  for (const [path, spec] of Object.entries(charts)) {
    const raw = files[path];
    if (!raw) continue;
    let xml = strFromU8(raw);

    if (spec.title) {
      const t = matchBlocks(xml, "c:title")[0];
      if (t) {
        let done = false;
        const newTitle = t.text.replace(/<a:t>[\s\S]*?<\/a:t>/, () => {
          done = true;
          return `<a:t>${xmlEscape(spec.title!)}</a:t>`;
        });
        if (done) {
          xml = xml.slice(0, t.start) + newTitle + xml.slice(t.end);
          filled++;
        }
      }
    }

    const sers = matchBlocks(xml, "c:ser");
    for (let i = sers.length - 1; i >= 0; i--) {
      const s = sers[i]!;
      const spec_i = spec.series?.[i];
      let sXml = s.text;

      if (spec.categories?.length) {
        const cat = matchBlocks(sXml, "c:cat")[0];
        if (cat) {
          sXml =
            sXml.slice(0, cat.start) + replaceVals(cat.text, spec.categories) + sXml.slice(cat.end);
          filled++;
        }
      }
      if (spec_i?.values?.length) {
        const val = matchBlocks(sXml, "c:val")[0];
        if (val) {
          sXml =
            sXml.slice(0, val.start) +
            replaceVals(val.text, spec_i.values.map((v) => String(v))) +
            sXml.slice(val.end);
          filled++;
        }
      }
      if (spec_i?.name) {
        const tx = matchBlocks(sXml, "c:tx")[0];
        if (tx) {
          sXml = sXml.slice(0, tx.start) + replaceVals(tx.text, [spec_i.name]) + sXml.slice(tx.end);
          filled++;
        }
      }
      xml = xml.slice(0, s.start) + sXml + xml.slice(s.end);
    }
    files[path] = strToU8(xml);
  }
  return filled;
}

/** Write SmartArt node labels back, by position. */
function fillDiagrams(files: Record<string, Uint8Array>, diagrams: Record<string, string[]>) {
  let filled = 0;
  for (const [path, nodes] of Object.entries(diagrams)) {
    const raw = files[path];
    if (!raw) continue;
    let xml = strFromU8(raw);
    const paras = matchBlocks(xml, "a:p");
    for (let i = paras.length - 1; i >= 0; i--) {
      const value = nodes[i];
      if (value === undefined) continue;
      const replaced = setParagraphText(paras[i]!.text, value);
      if (replaced === paras[i]!.text) continue;
      xml = xml.slice(0, paras[i]!.start) + replaced + xml.slice(paras[i]!.end);
      filled++;
    }
    files[path] = strToU8(xml);
  }
  return filled;
}

/** Everything the AI can fill in one pass, read straight out of the template. */
export function extractStructures(bytes: Uint8Array) {
  const files = unzipSync(bytes);
  return {
    tables: extractTables(files),
    charts: extractCharts(files),
    diagrams: extractDiagrams(files),
  };
}

export type ComplexEdits = {
  tables?: Record<string, string[][]>;
  charts?: Record<string, { title?: string; categories?: string[]; series?: { name?: string; values?: number[] }[] }>;
  diagrams?: Record<string, string[]>;
};

/** Text + media + tables/charts/diagrams, all written into the same file. */
export function buildDeckFull(
  bytes: Uint8Array,
  values: Record<string, string>,
  media: Record<string, Uint8Array> = {},
  complex: ComplexEdits = {},
) {
  const { bytes: filledBytes, filled } = fillTemplate(bytes, values);
  const files = unzipSync(filledBytes);

  let replacedMedia = 0;
  for (const key of Object.keys(media)) {
    if (!files[key]) continue;
    files[key] = new Uint8Array(media[key]!);
    replacedMedia++;
  }

  const filledComplex =
    fillTables(files, complex.tables ?? {}) +
    fillCharts(files, complex.charts ?? {}) +
    fillDiagrams(files, complex.diagrams ?? {});

  return {
    bytes: new Uint8Array(zipSync(files, { level: 6 })),
    filled,
    replacedMedia,
    filledComplex,
  };
}
