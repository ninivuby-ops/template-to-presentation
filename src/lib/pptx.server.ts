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
