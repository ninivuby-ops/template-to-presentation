import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

const Input = z.object({
  fileBase64: z.string().min(10),
  fileName: z.string().default("template.pptx"),
  topic: z.string().min(3).max(300),
  audience: z.string().max(200).optional().default(""),
  details: z.string().max(300000).optional().default(""),
});

const b64ToBytes = (b64: string) => {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
};

const bytesToB64 = (bytes: Uint8Array) => {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
};

/** Step 1 — inspect the template only. Nothing is modified. */
export const inspectTemplate = createServerFn({ method: "POST" })
  .inputValidator((d: { fileBase64: string }) =>
    z.object({ fileBase64: z.string().min(10) }).parse(d),
  )
  .handler(async ({ data }) => {
    const { extractPlaceholders } = await import("./pptx.server");
    const { placeholders, slideCount } = extractPlaceholders(b64ToBytes(data.fileBase64));
    return {
      slideCount,
      placeholders: placeholders.map(({ id, slide, shape, kind, current }) => ({
        id,
        slide,
        shape,
        kind,
        current,
      })),
    };
  });

/** Step 2 — draft the content. Nothing is written to the file yet: the user edits first. */
export const generateDeck = createServerFn({ method: "POST" })
  .inputValidator((d: unknown) => Input.parse(d))
  .handler(async ({ data }) => {
    const { extractPlaceholders, extractMedia, extractStructures } = await import("./pptx.server");
    const bytes = b64ToBytes(data.fileBase64);
    const { placeholders, slideCount } = extractPlaceholders(bytes);
    const media = extractMedia(bytes);
    const { tables, charts, diagrams } = extractStructures(bytes);

    const empty = {
      slideCount,
      media,
      placeholders: [] as { id: string; slide: number; shape: string; kind: string; text: string }[],
      tables,
      charts,
      diagrams,
    };

    if (!placeholders.length && !tables.length && !charts.length && !diagrams.length) {
      return {
        ...empty,
        note: "No empty or marked placeholders, tables, charts or diagrams were found in this template.",
      };
    }

    const apiKey = process.env["LOVABLE_API_KEY"];
    if (!apiKey) throw new Error("AI is not configured for this project.");

    const brief = placeholders.map((p) => ({
      id: p.id,
      slide: p.slide,
      role: p.kind,
      shape: p.shape,
      existing: p.current,
      slide_context: p.context.slice(0, 12),
    }));

    const prompt = [
      `You are filling in an existing PowerPoint template. Topic: "${data.topic}".`,
      data.audience ? `Audience: ${data.audience}.` : "",
      data.details ? `Extra source material provided by the user:\n${data.details}` : "",
      "",
      "STRICT RULES:",
      "- Only produce replacement content for the ids given below.",
      "- Never invent new slides, rows, columns, series or nodes; every structure is fixed in size.",
      "- Respect each slide's existing headings and prescribed pointers (given in slide_context) and stay on that subject.",
      "- A `title` role gets a short line (max 8 words). Body placeholders get one concise sentence or bullet (max 28 words).",
      "- Plain text only: no markdown, no bullet characters, no quotes around the text, no line breaks.",
      "- Tables: return exactly the same number of rows and cells as given. Keep header rows as headers. Table cells are short (max 6 words).",
      "- Charts: keep the same number of categories and the same number of values per series as given. Category labels are short. Values are plain numbers, realistic and internally consistent.",
      "- Diagrams: return exactly the same number of nodes, in order. Empty nodes stay empty (\"\"). Node labels are max 6 words.",
      "- Only use figures the user supplied; where none exist, use clearly plausible round illustrative numbers and never present them as sourced facts.",
      "- Do not fabricate specific statistics, dates, customer names or citations in text.",
      "",
      "Text placeholders:",
      JSON.stringify(brief),
      "",
      "Tables:",
      JSON.stringify(tables),
      "",
      "Charts:",
      JSON.stringify(charts),
      "",
      "Diagrams:",
      JSON.stringify(diagrams),
      "",
      'Reply with JSON only, exactly: {"items":[{"id":"<id>","text":"<replacement>"}],"tables":[{"id":"<id>","rows":[["<cell>"]]}],"charts":[{"id":"<id>","title":"<title>","categories":["<label>"],"series":[{"name":"<name>","values":[0]}]}],"diagrams":[{"id":"<id>","nodes":["<label>"]}]} covering every id.',
    ]
      .filter(Boolean)
      .join("\n");

    const res = await fetch("https://ai.gateway.lovable.dev/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "openai/gpt-5.6-sol",
        input: [{ role: "user", content: prompt }],
      }),
    });

    if (res.status === 429) throw new Error("Too many requests right now — please retry in a moment.");
    if (res.status === 402) throw new Error("AI credits are exhausted for this workspace.");
    if (!res.ok) throw new Error(`AI request failed (${res.status}).`);

    const json: any = await res.json();
    const text: string =
      json.output_text ??
      (json.output ?? [])
        .flatMap((o: any) => o.content ?? [])
        .map((c: any) => c.text ?? "")
        .join("");

    const raw = text.match(/\{[\s\S]*\}/)?.[0] ?? "{}";
    const parsed = z
      .object({
        items: z.array(z.object({ id: z.string(), text: z.string() })).default([]),
        tables: z
          .array(z.object({ id: z.string(), rows: z.array(z.array(z.string())) }))
          .default([]),
        charts: z
          .array(
            z.object({
              id: z.string(),
              title: z.string().default(""),
              categories: z.array(z.string()).default([]),
              series: z
                .array(z.object({ name: z.string().default(""), values: z.array(z.number()).default([]) }))
                .default([]),
            }),
          )
          .default([]),
        diagrams: z.array(z.object({ id: z.string(), nodes: z.array(z.string()) })).default([]),
      })
      .safeParse(JSON.parse(raw));
    if (!parsed.success) throw new Error("The AI response could not be read. Please try again.");

    const values: Record<string, string> = {};
    for (const item of parsed.data.items) values[item.id] = item.text.trim();

    const byId = <T extends { id: string }>(list: T[]) =>
      Object.fromEntries(list.map((x) => [x.id, x])) as Record<string, T>;
    const tableOut = byId(parsed.data.tables);
    const chartOut = byId(parsed.data.charts);
    const diagramOut = byId(parsed.data.diagrams);

    return {
      slideCount,
      media,
      placeholders: placeholders.map((p) => ({
        id: p.id,
        slide: p.slide,
        shape: p.shape,
        kind: p.kind,
        text: values[p.id] ?? p.current,
      })),
      tables: tables.map((t) => ({
        ...t,
        rows: t.rows.map((row, ri) =>
          row.map((cell, ci) => tableOut[t.id]?.rows?.[ri]?.[ci] ?? cell),
        ),
      })),
      charts: charts.map((c) => {
        const g = chartOut[c.id];
        return {
          ...c,
          title: g?.title || c.title,
          categories: c.categories.map((cat, i) => g?.categories?.[i] ?? cat),
          series: c.series.map((s2, i) => ({
            name: g?.series?.[i]?.name || s2.name,
            values: s2.values.map((v, vi) => g?.series?.[i]?.values?.[vi] ?? v),
          })),
        };
      }),
      diagrams: diagrams.map((d) => ({
        ...d,
        nodes: d.nodes.map((n, i) => (n.trim() ? diagramOut[d.id]?.nodes?.[i] ?? n : n)),
      })),
      note: "",
    };
  });

const BuildInput = z.object({
  fileBase64: z.string().min(10),
  items: z.array(z.object({ id: z.string(), text: z.string().max(2000) })).default([]),
  media: z
    .array(z.object({ id: z.string(), fileBase64: z.string().min(10) }))
    .max(40)
    .default([]),
  tables: z.array(z.object({ id: z.string(), rows: z.array(z.array(z.string())) })).default([]),
  charts: z
    .array(
      z.object({
        id: z.string(),
        title: z.string().default(""),
        categories: z.array(z.string()).default([]),
        series: z.array(z.object({ name: z.string().default(""), values: z.array(z.number()).default([]) })).default([]),
      }),
    )
    .default([]),
  diagrams: z.array(z.object({ id: z.string(), nodes: z.array(z.string()) })).default([]),
});

/** Step 3 — write the edited text and swapped media into the very same file. */
export const buildDeckFile = createServerFn({ method: "POST" })
  .inputValidator((d: unknown) => BuildInput.parse(d))
  .handler(async ({ data }) => {
    const { buildDeckFull } = await import("./pptx.server");
    const bytes = b64ToBytes(data.fileBase64);

    const values: Record<string, string> = {};
    for (const item of data.items) if (item.text.trim()) values[item.id] = item.text.trim();

    const media: Record<string, Uint8Array> = {};
    for (const m of data.media) media[m.id] = b64ToBytes(m.fileBase64);

    const complex = {
      tables: Object.fromEntries(data.tables.map((t) => [t.id, t.rows])),
      charts: Object.fromEntries(
        data.charts.map((c) => [c.id, { title: c.title, categories: c.categories, series: c.series }]),
      ),
      diagrams: Object.fromEntries(data.diagrams.map((d) => [d.id, d.nodes])),
    };

    const { bytes: out, filled, replacedMedia, filledComplex } = buildDeckFull(
      bytes,
      values,
      media,
      complex,
    );
    return { filled, replacedMedia, filledComplex, fileBase64: bytesToB64(out) };
  });
