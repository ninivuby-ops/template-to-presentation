import { createFileRoute } from "@tanstack/react-router";
import { useServerFn } from "@tanstack/react-start";
import { useEffect, useState } from "react";
import { toast } from "sonner";
import { Toaster } from "@/components/ui/sonner";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import {
  UploadCloud,
  FileDown,
  Loader2,
  ShieldCheck,
  Image as ImageIcon,
  Film,
  Music,
  RotateCcw,
  Table as TableIcon,
  BarChart3,
  Workflow,
  Sparkles,
} from "lucide-react";
import { generateDeck, buildDeckFile } from "@/lib/deck.functions";
import sihTemplate from "@/assets/sih2026-template.pptx.asset.json";
import hero from "@/assets/hero.jpg";
import architecture from "@/assets/architecture.jpg";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "SIH 2026 Deck Builder — AI fills the official idea format" },
      {
        name: "description",
        content:
          "Write your topic and get a finished Smart India Hackathon 2026 idea presentation, built on the official format with every heading, footer and layout untouched.",
      },
      { property: "og:title", content: "SIH 2026 Deck Builder" },
      {
        property: "og:description",
        content:
          "Generate a complete Smart India Hackathon 2026 idea presentation in the official format with one click.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

type Result = Awaited<ReturnType<typeof generateDeck>>;
type Tables = Result["tables"];
type Charts = Result["charts"];
type Diagrams = Result["diagrams"];
type MediaSwap = { name: string; fileBase64: string; preview?: string | undefined };

const wordCount = (s: string) => (s.trim() ? s.trim().split(/\s+/).length : 0);

const kb = (n: number) => `${Math.max(1, Math.round(n / 1024))} KB`;

const rules = [
  "Always the official SIH 2026 idea format — no other layout.",
  "Slide order, section headings and pointers stay untouched.",
  "No slide is added or removed.",
  "Text goes only into the format's own placeholders.",
  "Fonts, colours and layout are preserved run-for-run.",
];

function Index() {
  const run = useServerFn(generateDeck);
  const build = useServerFn(buildDeckFile);
  const [file, setFile] = useState<File | null>(null);
  const [templateError, setTemplateError] = useState(false);
  const [topic, setTopic] = useState("");
  const [audience, setAudience] = useState("");
  const [details, setDetails] = useState("");
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<Result | null>(null);
  const [draft, setDraft] = useState<Record<string, string>>({});
  const [swaps, setSwaps] = useState<Record<string, MediaSwap>>({});
  const [building, setBuilding] = useState(false);
  const [baseFile, setBaseFile] = useState("");
  const [tables, setTables] = useState<Tables>([]);
  const [charts, setCharts] = useState<Charts>([]);
  const [diagrams, setDiagrams] = useState<Diagrams>([]);

  useEffect(() => {
    let alive = true;
    fetch(sihTemplate.url)
      .then((r) => (r.ok ? r.blob() : Promise.reject(new Error("load"))))
      .then((b) => {
        if (alive)
          setFile(
            new File([b], "SIH2026-Idea-Presentation.pptx", {
              type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            }),
          );
      })
      .catch(() => alive && setTemplateError(true));
    return () => {
      alive = false;
    };
  }, []);

  const readBase64 = (f: File) =>

    new Promise<string>((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result).split(",")[1] ?? "");
      r.onerror = () => reject(new Error("Could not read that file."));
      r.readAsDataURL(f);
    });

  const onSubmit = async (e: React.FormEvent, oneClick = false) => {
    e.preventDefault();
    if (!file) {
      toast.error("The SIH 2026 format is still loading — try again in a second.");
      return;
    }
    if (wordCount(details) > 40000) {
      toast.error("Please keep the facts and notes under 40,000 words.");
      return;
    }
    if (topic.trim().length < 3) {
      toast.error("Please describe the topic.");
      return;
    }
    setBusy(true);
    setResult(null);
    try {
      const fileBase64 = await readBase64(file);
      const res = await run({
        data: { fileBase64, fileName: file.name, topic, audience, details },
      });
      setResult(res);
      setBaseFile(fileBase64);
      setDraft(Object.fromEntries(res.placeholders.map((p) => [p.id, p.text])));
      setTables(res.tables);
      setCharts(res.charts);
      setDiagrams(res.diagrams);
      setSwaps({});
      if (oneClick) {
        await doBuild({
          fileBase64,
          items: res.placeholders.map((p) => ({ id: p.id, text: p.text })),
          media: [],
          tables: res.tables.map((t) => ({ id: t.id, rows: t.rows })),
          charts: res.charts.map((c) => ({
            id: c.id,
            title: c.title,
            categories: c.categories,
            series: c.series,
          })),
          diagrams: res.diagrams.map((d) => ({ id: d.id, nodes: d.nodes })),
        });
        return;
      }
      toast.success(
        `Drafted ${res.placeholders.length} text blocks, ${res.tables.length} tables, ${res.charts.length} charts and ${res.diagrams.length} diagrams — edit below, then build.`,
      );
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Something went wrong.");
    } finally {
      setBusy(false);
    }
  };

  const pickMedia = (id: string, f: File | null) => {
    if (!f) return;
    if (f.size > 12_000_000) {
      toast.error("Please choose a file under 12 MB.");
      return;
    }
    const r = new FileReader();
    r.onload = () => {
      const url = String(r.result);
      setSwaps((s) => ({
        ...s,
        [id]: {
          name: f.name,
          fileBase64: url.split(",")[1] ?? "",
          preview: f.type.startsWith("image/") ? url : undefined,
        },
      }));
    };
    r.readAsDataURL(f);
  };

  type BuildPayload = {
    fileBase64: string;
    items: { id: string; text: string }[];
    media: { id: string; fileBase64: string }[];
    tables: { id: string; rows: string[][] }[];
    charts: { id: string; title: string; categories: string[]; series: { name: string; values: number[] }[] }[];
    diagrams: { id: string; nodes: string[] }[];
  };

  const doBuild = async (payload: BuildPayload) => {
    if (!file) return;
    setBuilding(true);
    try {
      const out = await build({ data: payload });
      const bin = atob(out.fileBase64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const url = URL.createObjectURL(
        new Blob([bytes], {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        }),
      );
      const a = document.createElement("a");
      a.href = url;
      a.download = file.name.replace(/\.pptx$/i, "") + "-filled.pptx";
      a.click();
      URL.revokeObjectURL(url);
      toast.success(
        `Deck ready — ${out.filled} text blocks, ${out.filledComplex} table/chart/diagram edits` +
          (out.replacedMedia ? ` and ${out.replacedMedia} media files replaced.` : "."),
      );
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Could not build the deck.");
    } finally {
      setBuilding(false);
    }
  };

  const buildAndDownload = () =>
    doBuild({
      fileBase64: baseFile,
      items: Object.entries(draft).map(([id, text]) => ({ id, text })),
      media: Object.entries(swaps).map(([id, m]) => ({ id, fileBase64: m.fileBase64 })),
      tables: tables.map((t) => ({ id: t.id, rows: t.rows })),
      charts: charts.map((c) => ({
        id: c.id,
        title: c.title,
        categories: c.categories,
        series: c.series,
      })),
      diagrams: diagrams.map((d) => ({ id: d.id, nodes: d.nodes })),
    });

  return (
    <main className="min-h-screen bg-background">
      <Toaster />
      <section className="bg-hero border-b border-border">
        <div className="mx-auto grid max-w-6xl items-center gap-10 px-6 py-16 md:grid-cols-2 md:py-24">
          <div>
            <p className="mb-4 inline-flex items-center gap-2 rounded-full border border-border bg-card/60 px-3 py-1 text-xs uppercase tracking-widest text-accent">
              <ShieldCheck className="h-3.5 w-3.5" /> SIH 2026 format only
            </p>
            <h1 className="text-4xl font-semibold leading-tight tracking-tight text-foreground md:text-5xl">
              Your idea, written into the official SIH 2026 deck.
            </h1>
            <p className="mt-5 max-w-lg text-muted-foreground">
              Describe your problem statement and solution — every section of the Smart India Hackathon
              2026 idea format comes back written, with the headings, footers and styling untouched.
            </p>

            <ul className="mt-6 space-y-2 text-sm text-muted-foreground">
              {rules.map((r) => (
                <li key={r} className="flex gap-2">
                  <span className="text-primary">—</span>
                  {r}
                </li>
              ))}
            </ul>
          </div>
          <img
            src={hero}
            alt="Presentation slides being populated automatically"
            width={1280}
            height={800}
            className="rounded-2xl border border-border shadow-panel"
          />
        </div>
      </section>

      <section className="mx-auto max-w-3xl px-6 py-14">
        <form
          onSubmit={onSubmit}
          className="space-y-5 rounded-2xl border border-border bg-card p-6 shadow-panel md:p-8"
        >
          <div>
            <Label>Presentation format</Label>
            <div className="mt-2 flex items-center gap-3 rounded-xl border border-border bg-secondary/40 px-4 py-5 text-sm">
              <ShieldCheck className="h-5 w-5 shrink-0 text-primary" />
              <span className={templateError ? "text-destructive" : "text-muted-foreground"}>
                {templateError
                  ? "The SIH 2026 format could not be loaded — please refresh the page."
                  : file
                    ? "Official SIH 2026 idea format — ready"
                    : "Loading the official SIH 2026 idea format…"}
              </span>
            </div>
          </div>


          <div className="grid gap-4 md:grid-cols-2">
            <div>
              <Label htmlFor="topic">Topic</Label>
              <Input
                id="topic"
                value={topic}
                onChange={(e) => setTopic(e.target.value)}
                placeholder="AI-based flood early warning system"
                className="mt-2"
              />
            </div>
            <div>
              <Label htmlFor="aud">Audience (optional)</Label>
              <Input
                id="aud"
                value={audience}
                onChange={(e) => setAudience(e.target.value)}
                placeholder="SIH evaluation jury"
                className="mt-2"
              />
            </div>
          </div>

          <div>
            <Label htmlFor="det">Facts and notes to use (optional)</Label>
            <Textarea
              id="det"
              value={details}
              onChange={(e) => setDetails(e.target.value)}
              rows={7}
              placeholder="Paste any real figures, names or context you want included."
              className="mt-2"
            />
            <p className="mt-1 text-right text-xs text-muted-foreground">
              {wordCount(details).toLocaleString()} / 40,000 words
            </p>
          </div>

          <div className="grid gap-3 sm:grid-cols-2">
            <Button
              type="button"
              disabled={busy || building}
              size="lg"
              onClick={(e) => onSubmit(e, true)}
            >
              {busy || building ? (
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
              ) : (
                <Sparkles className="mr-2 h-4 w-4" />
              )}
              {busy ? "Writing your deck…" : building ? "Building…" : "Generate & download deck"}
            </Button>
            <Button type="submit" disabled={busy || building} size="lg" variant="secondary">
              Generate, then let me edit
            </Button>
          </div>

        </form>

        {result ? (
          <div className="mt-8 space-y-6">
            <div className="rounded-2xl border border-border bg-card p-6 shadow-panel">
              <div className="flex flex-wrap items-center justify-between gap-4">
                <div>
                  <h2 className="text-lg font-semibold text-foreground">
                    Review & edit before building · {result.slideCount} slides kept
                  </h2>
                  <p className="text-sm text-muted-foreground">
                    {result.note || "Nothing is written into your file until you build it."}
                  </p>
                </div>
                <Button onClick={buildAndDownload} disabled={building} variant="secondary">
                  {building ? (
                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  ) : (
                    <FileDown className="mr-2 h-4 w-4" />
                  )}
                  {building ? "Building…" : "Build & download .pptx"}
                </Button>
              </div>

              {result.placeholders.length ? (
                <div className="mt-6 space-y-4">
                  {result.placeholders.map((p) => (
                    <div key={p.id} className="rounded-xl bg-secondary/40 p-3">
                      <div className="mb-2 flex items-center justify-between gap-3">
                        <span className="text-xs uppercase tracking-wide text-accent">
                          Slide {p.slide} · {p.shape} · {p.kind}
                        </span>
                        <button
                          type="button"
                          onClick={() => setDraft((d) => ({ ...d, [p.id]: p.text }))}
                          className="inline-flex items-center gap-1 text-xs text-muted-foreground hover:text-foreground"
                        >
                          <RotateCcw className="h-3 w-3" /> Reset
                        </button>
                      </div>
                      <Textarea
                        rows={p.kind === "title" ? 1 : 2}
                        value={draft[p.id] ?? ""}
                        onChange={(e) => setDraft((d) => ({ ...d, [p.id]: e.target.value }))}
                      />
                    </div>
                  ))}
                </div>
              ) : null}
            </div>

            {result.media.length ? (
              <div className="rounded-2xl border border-border bg-card p-6 shadow-panel">
                <h3 className="text-base font-semibold text-foreground">Images & video</h3>
                <p className="mt-1 text-sm text-muted-foreground">
                  Swap any picture, video or audio already in the template. Use the same file type so the
                  slide stays intact. Nothing is added or removed.
                </p>
                <div className="mt-5 grid gap-4 sm:grid-cols-2">
                  {result.media.map((m) => {
                    const swap = swaps[m.id];
                    const Icon = m.kind === "video" ? Film : m.kind === "audio" ? Music : ImageIcon;
                    return (
                      <div key={m.id} className="rounded-xl bg-secondary/40 p-3">
                        <div className="flex items-center gap-2 text-xs uppercase tracking-wide text-accent">
                          <Icon className="h-3.5 w-3.5" />
                          {m.kind} · {m.slides.length ? `slide ${m.slides.join(", ")}` : "shared"} ·{" "}
                          {kb(m.bytes)}
                        </div>
                        {swap?.preview || m.preview ? (
                          <img
                            src={swap?.preview ?? m.preview}
                            alt={`Template asset ${m.name}`}
                            loading="lazy"
                            className="mt-3 h-32 w-full rounded-lg border border-border object-contain"
                          />
                        ) : null}
                        <label
                          htmlFor={`media-${m.id}`}
                          className="mt-3 flex cursor-pointer items-center gap-2 rounded-lg border border-dashed border-border px-3 py-2 text-xs text-muted-foreground hover:border-primary"
                        >
                          <UploadCloud className="h-4 w-4 text-primary" />
                          {swap ? swap.name : `Replace ${m.name}`}
                        </label>
                        <input
                          id={`media-${m.id}`}
                          type="file"
                          accept="image/*,video/*,audio/*"
                          className="sr-only"
                          onChange={(e) => pickMedia(m.id, e.target.files?.[0] ?? null)}
                        />
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : null}

            {tables.length ? (
              <div className="rounded-2xl border border-border bg-card p-6 shadow-panel">
                <h3 className="flex items-center gap-2 text-base font-semibold text-foreground">
                  <TableIcon className="h-4 w-4 text-primary" /> Tables
                </h3>
                <p className="mt-1 text-sm text-muted-foreground">
                  Same rows and columns as your template — only the wording changes.
                </p>
                <div className="mt-5 space-y-6">
                  {tables.map((t, ti) => (
                    <div key={t.id} className="overflow-x-auto rounded-xl bg-secondary/40 p-3">
                      <div className="mb-2 text-xs uppercase tracking-wide text-accent">
                        Slide {t.slide}
                      </div>
                      <table className="w-full border-separate border-spacing-1">
                        <tbody>
                          {t.rows.map((row, ri) => (
                            <tr key={ri}>
                              {row.map((cell, ci) => (
                                <td key={ci}>
                                  <Input
                                    value={cell}
                                    aria-label={`Row ${ri + 1} column ${ci + 1}`}
                                    onChange={(e) =>
                                      setTables((prev) =>
                                        prev.map((tt, i) =>
                                          i !== ti
                                            ? tt
                                            : {
                                                ...tt,
                                                rows: tt.rows.map((rr, j) =>
                                                  j !== ri
                                                    ? rr
                                                    : rr.map((cc, k) => (k === ci ? e.target.value : cc)),
                                                ),
                                              },
                                        ),
                                      )
                                    }
                                    className="h-9 min-w-36 text-sm"
                                  />
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ))}
                </div>
              </div>
            ) : null}

            {charts.length ? (
              <div className="rounded-2xl border border-border bg-card p-6 shadow-panel">
                <h3 className="flex items-center gap-2 text-base font-semibold text-foreground">
                  <BarChart3 className="h-4 w-4 text-primary" /> Charts
                </h3>
                <p className="mt-1 text-sm text-muted-foreground">
                  Labels and numbers go straight into the chart already on the slide — its type, colours
                  and size stay as designed.
                </p>
                <div className="mt-5 space-y-6">
                  {charts.map((c, ci) => (
                    <div key={c.id} className="rounded-xl bg-secondary/40 p-3">
                      <div className="mb-2 text-xs uppercase tracking-wide text-accent">
                        {c.slide ? `Slide ${c.slide}` : "Chart"} · {c.kind}
                      </div>
                      <Input
                        value={c.title}
                        aria-label="Chart title"
                        placeholder="Chart title"
                        onChange={(e) =>
                          setCharts((prev) =>
                            prev.map((cc, i) => (i === ci ? { ...cc, title: e.target.value } : cc)),
                          )
                        }
                        className="mb-3 h-9 text-sm"
                      />
                      <div className="overflow-x-auto">
                        <table className="border-separate border-spacing-1 text-sm">
                          <thead>
                            <tr>
                              <th className="text-left text-xs font-normal text-muted-foreground">
                                Series
                              </th>
                              {c.categories.map((cat, k) => (
                                <th key={k}>
                                  <Input
                                    value={cat}
                                    aria-label={`Category ${k + 1}`}
                                    onChange={(e) =>
                                      setCharts((prev) =>
                                        prev.map((cc, i) =>
                                          i !== ci
                                            ? cc
                                            : {
                                                ...cc,
                                                categories: cc.categories.map((x, j) =>
                                                  j === k ? e.target.value : x,
                                                ),
                                              },
                                        ),
                                      )
                                    }
                                    className="h-9 min-w-28 text-sm"
                                  />
                                </th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {c.series.map((s, si) => (
                              <tr key={si}>
                                <td>
                                  <Input
                                    value={s.name}
                                    aria-label={`Series ${si + 1} name`}
                                    onChange={(e) =>
                                      setCharts((prev) =>
                                        prev.map((cc, i) =>
                                          i !== ci
                                            ? cc
                                            : {
                                                ...cc,
                                                series: cc.series.map((ss, j) =>
                                                  j === si ? { ...ss, name: e.target.value } : ss,
                                                ),
                                              },
                                        ),
                                      )
                                    }
                                    className="h-9 min-w-28 text-sm"
                                  />
                                </td>
                                {s.values.map((v, vi) => (
                                  <td key={vi}>
                                    <Input
                                      type="number"
                                      value={String(v)}
                                      aria-label={`Series ${si + 1} value ${vi + 1}`}
                                      onChange={(e) =>
                                        setCharts((prev) =>
                                          prev.map((cc, i) =>
                                            i !== ci
                                              ? cc
                                              : {
                                                  ...cc,
                                                  series: cc.series.map((ss, j) =>
                                                    j !== si
                                                      ? ss
                                                      : {
                                                          ...ss,
                                                          values: ss.values.map((x, k) =>
                                                            k === vi ? Number(e.target.value) || 0 : x,
                                                          ),
                                                        },
                                                  ),
                                                },
                                          ),
                                        )
                                      }
                                      className="h-9 w-28 text-sm"
                                    />
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ) : null}

            {diagrams.length ? (
              <div className="rounded-2xl border border-border bg-card p-6 shadow-panel">
                <h3 className="flex items-center gap-2 text-base font-semibold text-foreground">
                  <Workflow className="h-4 w-4 text-primary" /> Diagrams
                </h3>
                <p className="mt-1 text-sm text-muted-foreground">
                  Each shape in your existing diagram keeps its place; only the label changes.
                </p>
                <div className="mt-5 space-y-5">
                  {diagrams.map((d, di) => (
                    <div key={d.id} className="rounded-xl bg-secondary/40 p-3">
                      <div className="mb-2 text-xs uppercase tracking-wide text-accent">
                        {d.slide ? `Slide ${d.slide}` : "Diagram"} · {d.nodes.length} shapes
                      </div>
                      <div className="grid gap-2 sm:grid-cols-2">
                        {d.nodes.map((n, ni) => (
                          <Input
                            key={ni}
                            value={n}
                            aria-label={`Diagram label ${ni + 1}`}
                            onChange={(e) =>
                              setDiagrams((prev) =>
                                prev.map((dd, i) =>
                                  i !== di
                                    ? dd
                                    : {
                                        ...dd,
                                        nodes: dd.nodes.map((x, j) => (j === ni ? e.target.value : x)),
                                      },
                                ),
                              )
                            }
                            className="h-9 text-sm"
                          />
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ) : null}
          </div>

        ) : null}
      </section>

      <section className="border-t border-border bg-surface">
        <div className="mx-auto max-w-6xl px-6 py-14">
          <h2 className="text-2xl font-semibold text-foreground">How it works</h2>
          <p className="mt-2 max-w-2xl text-sm text-muted-foreground">
            We open the official SIH 2026 idea format, read its placeholders, write only that text with
            AI, put it back in place, and hand the finished file to you.
          </p>
          <img
            src={architecture}
            alt="Workflow: official SIH format, parse placeholders, generate content, inject text, download"
            loading="lazy"
            width={1600}
            height={704}
            className="mt-6 w-full rounded-2xl border border-border bg-card"
          />
        </div>
      </section>
    </main>
  );
}
