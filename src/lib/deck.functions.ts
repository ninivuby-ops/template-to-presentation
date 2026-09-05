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
    const { extractPlaceholders, extractMedia } = await import("./pptx.server");
    const bytes = b64ToBytes(data.fileBase64);
    const { placeholders, slideCount } = extractPlaceholders(bytes);
    const media = extractMedia(bytes);

    if (!placeholders.length) {
      return {
        slideCount,
        media,
        placeholders: [] as { id: string; slide: number; shape: string; kind: string; text: string }[],
        note: "No empty or marked placeholders were found in this template, so there is no text to write.",
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
      "- Only produce replacement text for the placeholder ids given below.",
      "- Never invent new slides, headings, or bullet groups; the deck structure is fixed.",
      "- Respect each slide's existing headings and prescribed pointers (given in slide_context) and stay on that subject.",
      "- A `title` role gets a short line (max 8 words). Body placeholders get one concise sentence or bullet (max 28 words).",
      "- Plain text only: no markdown, no bullet characters, no quotes around the text, no line breaks.",
      "- Do not fabricate specific statistics, dates, customer names or citations.",
      "",
      "Placeholders:",
      JSON.stringify(brief),
      "",
      'Reply with JSON only, exactly: {"items":[{"id":"<id>","text":"<replacement>"}]} covering every id.',
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
      .object({ items: z.array(z.object({ id: z.string(), text: z.string() })).default([]) })
      .safeParse(JSON.parse(raw));
    if (!parsed.success) throw new Error("The AI response could not be read. Please try again.");

    const values: Record<string, string> = {};
    for (const item of parsed.data.items) values[item.id] = item.text.trim();

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
});

/** Step 3 — write the edited text and swapped media into the very same file. */
export const buildDeckFile = createServerFn({ method: "POST" })
  .inputValidator((d: unknown) => BuildInput.parse(d))
  .handler(async ({ data }) => {
    const { buildDeck } = await import("./pptx.server");
    const bytes = b64ToBytes(data.fileBase64);

    const values: Record<string, string> = {};
    for (const item of data.items) if (item.text.trim()) values[item.id] = item.text.trim();

    const media: Record<string, Uint8Array> = {};
    for (const m of data.media) media[m.id] = b64ToBytes(m.fileBase64);

    const { bytes: out, filled, replacedMedia } = buildDeck(bytes, values, media);
    return { filled, replacedMedia, fileBase64: bytesToB64(out) };
  });
