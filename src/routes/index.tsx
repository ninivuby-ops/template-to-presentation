import { createFileRoute } from "@tanstack/react-router";
import { useServerFn } from "@tanstack/react-start";
import { useState } from "react";
import { toast } from "sonner";
import { Toaster } from "@/components/ui/sonner";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { UploadCloud, FileDown, Loader2, ShieldCheck } from "lucide-react";
import { generateDeck } from "@/lib/deck.functions";
import hero from "@/assets/hero.jpg";
import architecture from "@/assets/architecture.jpg";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "DeckFill — AI fills your PPTX template, exactly as designed" },
      {
        name: "description",
        content:
          "Upload a PowerPoint template, give a topic, and get every placeholder filled by AI with the original slides, fonts and layout untouched.",
      },
      { property: "og:title", content: "DeckFill — AI fills your PPTX template" },
      {
        property: "og:description",
        content:
          "Upload a PowerPoint template and download a fully populated deck that keeps your exact structure, fonts and colours.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

type Result = Awaited<ReturnType<typeof generateDeck>>;

const rules = [
  "Only your uploaded file is used — no template of ours.",
  "Slide order, section headings and pointers stay untouched.",
  "No slide is added or removed.",
  "Text goes only into the template's own placeholders.",
  "Fonts, colours and layout are preserved run-for-run.",
];

function Index() {
  const run = useServerFn(generateDeck);
  const [file, setFile] = useState<File | null>(null);
  const [topic, setTopic] = useState("");
  const [audience, setAudience] = useState("");
  const [details, setDetails] = useState("");
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<Result | null>(null);

  const readBase64 = (f: File) =>
    new Promise<string>((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(String(r.result).split(",")[1] ?? "");
      r.onerror = () => reject(new Error("Could not read that file."));
      r.readAsDataURL(f);
    });

  const onSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) {
      toast.error("Please choose a .pptx template first.");
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
      toast.success(
        res.filled ? `Filled ${res.filled} placeholders across ${res.slideCount} slides.` : "Finished.",
      );
    } catch (err) {
      toast.error(err instanceof Error ? err.message : "Something went wrong.");
    } finally {
      setBusy(false);
    }
  };

  const download = () => {
    if (!result || !file) return;
    const bin = atob(result.fileBase64);
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
  };

  return (
    <main className="min-h-screen bg-background">
      <Toaster />
      <section className="bg-hero border-b border-border">
        <div className="mx-auto grid max-w-6xl items-center gap-10 px-6 py-16 md:grid-cols-2 md:py-24">
          <div>
            <p className="mb-4 inline-flex items-center gap-2 rounded-full border border-border bg-card/60 px-3 py-1 text-xs uppercase tracking-widest text-accent">
              <ShieldCheck className="h-3.5 w-3.5" /> Template-faithful
            </p>
            <h1 className="text-4xl font-semibold leading-tight tracking-tight text-foreground md:text-5xl">
              Your template. Filled in by AI. Nothing else moved.
            </h1>
            <p className="mt-5 max-w-lg text-muted-foreground">
              Upload a PowerPoint file, tell us the subject, and every blank or marked placeholder comes
              back written — with the same slides, order, fonts and colours you designed.
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
            <Label htmlFor="tpl">PowerPoint template (.pptx)</Label>
            <label
              htmlFor="tpl"
              className="mt-2 flex cursor-pointer items-center gap-3 rounded-xl border border-dashed border-border bg-secondary/40 px-4 py-6 text-sm text-muted-foreground transition-colors hover:border-primary"
            >
              <UploadCloud className="h-5 w-5 text-primary" />
              {file ? file.name : "Choose or drop your template file"}
            </label>
            <input
              id="tpl"
              type="file"
              accept=".pptx"
              className="sr-only"
              onChange={(e) => setFile(e.target.files?.[0] ?? null)}
            />
          </div>

          <div className="grid gap-4 md:grid-cols-2">
            <div>
              <Label htmlFor="topic">Topic</Label>
              <Input
                id="topic"
                value={topic}
                onChange={(e) => setTopic(e.target.value)}
                placeholder="Q3 supply chain review"
                className="mt-2"
              />
            </div>
            <div>
              <Label htmlFor="aud">Audience (optional)</Label>
              <Input
                id="aud"
                value={audience}
                onChange={(e) => setAudience(e.target.value)}
                placeholder="Board of directors"
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
              rows={5}
              placeholder="Paste any real figures, names or context you want included."
              className="mt-2"
            />
          </div>

          <Button type="submit" disabled={busy} className="w-full" size="lg">
            {busy ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}
            {busy ? "Writing your deck…" : "Generate populated deck"}
          </Button>
        </form>

        {result ? (
          <div className="mt-8 rounded-2xl border border-border bg-card p-6 shadow-panel">
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div>
                <h2 className="text-lg font-semibold text-foreground">
                  {result.filled} placeholders filled · {result.slideCount} slides kept
                </h2>
                <p className="text-sm text-muted-foreground">
                  {result.note || "Structure, fonts and layout are unchanged."}
                </p>
              </div>
              <Button onClick={download} variant="secondary">
                <FileDown className="mr-2 h-4 w-4" /> Download .pptx
              </Button>
            </div>
            {result.placeholders.length ? (
              <ul className="mt-5 max-h-80 space-y-2 overflow-y-auto text-sm">
                {result.placeholders.map((p) => (
                  <li key={p.id} className="rounded-lg bg-secondary/50 px-3 py-2">
                    <span className="text-xs uppercase tracking-wide text-accent">
                      Slide {p.slide} · {p.shape}
                    </span>
                    <p className="text-foreground">{p.text}</p>
                  </li>
                ))}
              </ul>
            ) : null}
          </div>
        ) : null}
      </section>

      <section className="border-t border-border bg-surface">
        <div className="mx-auto max-w-6xl px-6 py-14">
          <h2 className="text-2xl font-semibold text-foreground">How it works</h2>
          <p className="mt-2 max-w-2xl text-sm text-muted-foreground">
            Upload, read the placeholders out of the file, write only that text with AI, put it back in
            place, and hand the same file back to you.
          </p>
          <img
            src={architecture}
            alt="Workflow: upload template, parse placeholders, generate content, inject text, download"
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
