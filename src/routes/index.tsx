import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useRef, useState } from "react";
import { Upload, FileSpreadsheet, X, AlertCircle, CircleCheck, Download, ShieldCheck, HelpCircle } from "lucide-react";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import kpmgLogo from "@/assets/kpmg-logo.png";
import { processSoxFiles, classifyError } from "@/lib/soxProcessor";
import { Send } from "lucide-react";

export const Route = createFileRoute("/")({
  component: Index,
});

type View = "upload" | "loading" | "success";

const LOADING_STEPS = [
  "מנתח מבנה הגיליונות...",
  "מזהה עמודות סטטוס ומפתח...",
  "מחשב סיכומים...",
  "בונה גיליונות פלט...",
  "שומר קובץ...",
];

function formatBytes(b: number): string {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / (1024 * 1024)).toFixed(1)} MB`;
}

function Index() {
  const [view, setView] = useState<View>("upload");
  const [files, setFiles] = useState<File[]>([]);
  const [error, setError] = useState<string>("");
  const [dragging, setDragging] = useState(false);
  const [progress, setProgress] = useState(0);
  const [stepIdx, setStepIdx] = useState(0);
  const [download, setDownload] = useState<{ url: string; name: string } | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    return () => {
      if (download?.url) URL.revokeObjectURL(download.url);
    };
  }, [download]);

  // Global drag-and-drop on the whole page (active in upload + success views).
  useEffect(() => {
    if (view === "loading") return;
    const onDragOver = (e: DragEvent) => {
      if (!e.dataTransfer?.types?.includes("Files")) return;
      e.preventDefault();
      setDragging(true);
    };
    const onDragLeave = (e: DragEvent) => {
      if ((e as DragEvent).relatedTarget === null) setDragging(false);
    };
    const onDrop = (e: DragEvent) => {
      if (!e.dataTransfer?.files?.length) return;
      e.preventDefault();
      setDragging(false);
      if (view === "success") {
        if (download?.url) URL.revokeObjectURL(download.url);
        setDownload(null);
        setFiles([]);
        setView("upload");
      }
      addFiles(e.dataTransfer.files);
    };
    window.addEventListener("dragover", onDragOver);
    window.addEventListener("dragleave", onDragLeave);
    window.addEventListener("drop", onDrop);
    return () => {
      window.removeEventListener("dragover", onDragOver);
      window.removeEventListener("dragleave", onDragLeave);
      window.removeEventListener("drop", onDrop);
    };
  }, [view, download]);

  const addFiles = (incoming: FileList | File[]) => {
    const arr = Array.from(incoming);
    const valid: File[] = [];
    for (const f of arr) {
      const ext = f.name.split(".").pop()?.toLowerCase();
      if (ext !== "xlsx") {
        setError("יש להעלות קבצים מסוג xlsx בלבד");
        return;
      }
      valid.push(f);
    }
    setError("");
    setFiles((prev) => {
      const seen = new Set(prev.map((p) => `${p.name}_${p.size}`));
      const merged = [...prev];
      for (const f of valid) {
        const k = `${f.name}_${f.size}`;
        if (!seen.has(k)) {
          seen.add(k);
          merged.push(f);
        }
      }
      return merged;
    });
  };

  const removeFile = (idx: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== idx));
  };

  const clearFiles = () => {
    setFiles([]);
    setError("");
    if (inputRef.current) inputRef.current.value = "";
  };

  const handleSubmit = async () => {
    if (!files.length) return;
    setView("loading");
    setProgress(0);
    setStepIdx(0);

    const tick = window.setInterval(() => {
      setProgress((p) => {
        const next = p + Math.random() * 14 + 5;
        return next > 92 ? 92 : next;
      });
      setStepIdx((s) => {
        const next = s + 1;
        return next < LOADING_STEPS.length ? next : s;
      });
    }, 600);

    try {
      await new Promise((r) => setTimeout(r, 50));
      const { blob, filename } = await processSoxFiles(files);
      window.clearInterval(tick);
      setProgress(100);

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      setDownload({ url, name: filename });
      window.setTimeout(() => setView("success"), 400);
    } catch (e) {
      window.clearInterval(tick);
      setError(classifyError(e));
      setView("upload");
    }
  };

  const reset = () => {
    setView("upload");
    clearFiles();
    if (download?.url) URL.revokeObjectURL(download.url);
    setDownload(null);
  };

  const redownload = () => {
    if (!download) return;
    const a = document.createElement("a");
    a.href = download.url;
    a.download = download.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <main className="min-h-screen flex flex-col items-center px-4 py-8 bg-background">
      <div className="w-full max-w-[520px] flex items-center justify-between mb-6">
        <Dialog>
          <DialogTrigger asChild>
            <button
              type="button"
              aria-label="עזרה"
              className="w-9 h-9 inline-flex items-center justify-center rounded-full border border-border text-textc-secondary hover:text-textc-primary hover:bg-muted transition-colors"
            >
              <HelpCircle className="w-5 h-5" />
            </button>
          </DialogTrigger>
          <DialogContent dir="rtl" className="max-w-md">
            <DialogHeader>
              <DialogTitle className="text-right">איך זה עובד?</DialogTitle>
            </DialogHeader>
            <Tabs defaultValue="how" dir="rtl" className="w-full">
              <TabsList className="grid grid-cols-2 w-full">
                <TabsTrigger value="how">עזרה</TabsTrigger>
                <TabsTrigger value="template">תבנית הקובץ</TabsTrigger>
              </TabsList>

              <TabsContent
                value="how"
                className="text-right text-sm text-textc-secondary space-y-3 leading-relaxed mt-4 h-[420px] overflow-y-auto pl-1"
              >
                <p>
                  הכלי מקבל קבצי <strong>Excel (xlsx בלבד)</strong> של מטריצות בקרות SOX, מזהה אוטומטית את שורת הכותרת,
                  עמודות הסטטוס, עמודת המפתח ועמודת התהליך בכל גיליון.
                </p>
                <div>
                  <div className="font-medium text-textc-primary mb-1">מה אפשר להעלות:</div>
                  <ul className="list-disc pr-5 space-y-1">
                    <li>קובץ xlsx אחד או יותר (גרירה או בחירה).</li>
                    <li>גיליונות בעברית או באנגלית — מעורבים גם יחד.</li>
                    <li>טבלאות עם עמודות סטטוס מרובות (סבבים).</li>
                  </ul>
                </div>
                <div>
                  <div className="font-medium text-textc-primary mb-1">מה תקבל בפלט:</div>
                  <ul className="list-disc pr-5 space-y-1">
                    <li>קובץ xlsx יחיד עם כל הגיליונות המקוריים.</li>
                    <li>גיליון "סיכום" אחד עם שורה לכל (גיליון × תהליך × סבב).</li>
                    <li>אם זוהו שתי שפות — נוצר סיכום באנגלית.</li>
                  </ul>
                </div>
                <p className="text-sm font-medium text-textc-primary">הקובץ המעובד יורד אוטומטית בסיום העיבוד.</p>
              </TabsContent>

              <TabsContent
                value="template"
                className="text-right text-sm text-textc-secondary space-y-3 leading-relaxed mt-4 h-[420px] overflow-y-auto pl-1"
              >
                <p>
                  כדי שגיליון יעובד, עליו להכיל את הרכיבים הבאים. גיליונות שאינם תואמים לתבנית —{" "}
                  <strong>ידולגו אוטומטית</strong>, והשאר יעובדו כרגיל.
                </p>

                <div>
                  <div className="font-medium text-textc-primary mb-1">שורת כותרת</div>
                  <p>
                    שורה אחת לפחות עם כותרות עמודות בעברית או באנגלית. ההתאמה של שמות העמודות היא מדויקת (אחרי הסרת
                    רווחים).
                  </p>
                </div>

                <div>
                  <div className="font-medium text-textc-primary mb-1">עמודות חובה</div>
                  <ul className="list-disc pr-5 space-y-2">
                    <li>
                      <span className="font-medium text-textc-primary">עמודת מסקנה / סטטוס</span> — אחת מהכותרות הבאות:
                      <div className="mt-1 text-[12px] bg-muted border border-border rounded px-2 py-1 leading-relaxed">
                        מסקנה · סטטוס · בקרה · מסקנה סבב א' · מסקנה סבב ב' · status · conclusion · control · results
                      </div>
                      <div className="text-[12px] text-textc-tertiary mt-1">
                        ניתן להופיע מספר עמודות סטטוס (סבב א', סבב ב' וכו') — כולן יסוכמו.
                      </div>
                    </li>
                    <li>
                      <span className="font-medium text-textc-primary">שורות נתונים</span> — בכל שורה מתחת לכותרת חייב
                      להיות ערך באחת מ-3 העמודות הראשונות (A / B / C). שורות ריקות מסיימות את הטבלה.
                    </li>
                  </ul>
                </div>

                <div>
                  <div className="font-medium text-textc-primary mb-1">עמודות רשות (משפרות את הסיכום)</div>
                  <ul className="list-disc pr-5 space-y-2">
                    <li>
                      <span className="font-medium text-textc-primary">תהליך</span> — כותרת: <code>תהליך</code> או{" "}
                      <code>process</code>.
                      <div className="text-[12px] text-textc-tertiary mt-1">
                        לא יזוהו עמודות כמו "Sub Process" / "שם תהליך משנה". אם חסרה — כל הבקרות יסוכמו תחת תהליך יחיד.
                      </div>
                    </li>
                    <li>
                      <span className="font-medium text-textc-primary">בקרת מפתח</span> — אחת מהכותרות:
                      <div className="mt-1 text-[12px] bg-muted border border-border rounded px-2 py-1 leading-relaxed">
                        מפתח · בקרת מפתח · key · key control
                      </div>
                      <div className="text-[12px] text-textc-tertiary mt-1">
                        בקרות עם הערך <code>Yes</code> / <code>Y / כן</code> ייספרו כבקרות מפתח.
                      </div>
                    </li>
                    <li>
                      <span className="font-medium text-textc-primary">מס׳ בקרה / ID</span> — לדוגמה:
                      <div className="mt-1 text-[12px] bg-muted border border-border rounded px-2 py-1 leading-relaxed">
                        מס · מס' · מספר · מס בקרה · control number · control no · id · no
                      </div>
                    </li>
                  </ul>
                </div>

                <div>
                  <div className="font-medium text-textc-primary mb-1">ערכים שלא נספרים בסטטוס</div>
                  <p>הערכים הבאים בעמודת הסטטוס לא ייספרו כסטטוס שבוצע:</p>
                  <div className="mt-1 text-[12px] bg-muted border border-border rounded px-2 py-1 leading-relaxed">
                    ריק · לא מפתח · not key · IT · שנתי · annual
                  </div>
                </div>
              </TabsContent>
            </Tabs>

            <div className="pt-3 mt-2 border-t border-border flex items-center justify-between gap-3">
              <span className="text-sm text-textc-secondary">נתקלת בבעיות?</span>
              <a
                href="mailto:refaelbar@KPMG.com;ybetito@KPMG.com?subject=נתקלתי בבעיה בכלי סכימת בקרות"
                className="inline-flex items-center gap-1.5 py-1.5 px-3 bg-blue-50 text-blue-800 border border-blue-200 rounded-md text-sm font-medium transition-colors hover:bg-blue-100"
              >
                פניה לצוות
              </a>
            </div>
          </DialogContent>
        </Dialog>
        <img src={kpmgLogo} alt="KPMG" className="h-8 w-auto opacity-90" />
      </div>

      <header className="w-full flex flex-col items-center text-center mb-10">
        <div className="inline-flex items-center justify-center gap-1.5 bg-blue-50 text-blue-800 text-xs font-medium px-3.5 py-1 rounded-full border border-blue-200 mb-3 tracking-wider">
          <ShieldCheck className="w-3.5 h-3.5" />
          SOX
        </div>
        <h1 className="text-[26px] font-semibold text-textc-primary leading-tight">
          סכימת <span className="text-blue-600">סטטוסים</span>
        </h1>
        <p className="text-sm text-textc-secondary mt-1.5">העלו קובץ/י אקסל לעיבוד אוטומטי של בקרות</p>
      </header>

      <section className="w-full max-w-[520px] bg-card border border-border rounded-xl p-8">
        {view === "upload" && (
          <div>
            <div
              role="button"
              tabIndex={0}
              aria-label="גררו לכאן קבצי Excel או לחצו לעיון"
              onClick={(e) => {
                if ((e.target as HTMLElement).tagName !== "BUTTON") inputRef.current?.click();
              }}
              onKeyDown={(e) => {
                if (e.key === "Enter" || e.key === " ") inputRef.current?.click();
              }}
              onDragOver={(e) => {
                e.preventDefault();
                setDragging(true);
              }}
              onDragLeave={() => setDragging(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragging(false);
                if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
              }}
              className={`border border-dashed rounded-lg py-10 px-6 flex flex-col items-center gap-2.5 cursor-pointer text-center transition-colors ${
                dragging
                  ? "bg-blue-50 border-blue-400"
                  : "bg-muted border-border hover:bg-blue-50 hover:border-blue-400"
              }`}
            >
              <div className="w-12 h-12 bg-blue-50 rounded-full flex items-center justify-center text-blue-600">
                <Upload className="w-5 h-5" />
              </div>
              <p className="text-[15px] font-medium text-textc-primary">גררו לכאן את הקבצים</p>
              <p className="text-[13px] text-textc-secondary">
                או{" "}
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    inputRef.current?.click();
                  }}
                  className="text-blue-600 underline hover:text-blue-800 font-medium"
                >
                  לחצו לעיון
                </button>{" "}
                בקבצים
              </p>
              <p className="text-[11px] text-textc-tertiary">ניתן לבחור מספר קבצי xlsx</p>
            </div>

            <input
              ref={inputRef}
              type="file"
              accept=".xlsx"
              multiple
              className="hidden"
              onChange={(e) => {
                if (e.target.files?.length) addFiles(e.target.files);
              }}
            />

            {files.length > 0 && (
              <ul className="mt-4 flex flex-col gap-2">
                {files.map((f, i) => (
                  <li
                    key={`${f.name}_${f.size}_${i}`}
                    className="flex items-center gap-2.5 py-2.5 px-3.5 bg-muted border border-border rounded-md"
                  >
                    <FileSpreadsheet className="w-5 h-5 text-teal-600 shrink-0" />
                    <div className="flex-1 min-w-0">
                      <div className="text-[13px] font-medium text-textc-primary truncate">{f.name}</div>
                      <div className="text-[11px] text-textc-secondary mt-0.5">{formatBytes(f.size)}</div>
                    </div>
                    <button
                      aria-label="הסר קובץ"
                      onClick={() => removeFile(i)}
                      className="text-textc-secondary hover:text-brandred-400 p-0.5 rounded"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  </li>
                ))}
              </ul>
            )}

            {error && (
              <div className="flex items-start gap-2 mt-4 py-2.5 px-3.5 bg-brandred-50 border border-brandred-400 rounded-md text-[13px] text-brandred-600">
                <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                <span className="whitespace-pre-line text-right leading-relaxed">{error}</span>
              </div>
            )}

            <button
              type="button"
              disabled={!files.length}
              onClick={handleSubmit}
              className="w-full mt-5 py-[11px] bg-blue-600 text-white border-0 rounded-md text-[15px] font-medium gap-2 transition-colors hover:bg-blue-800 disabled:opacity-45 disabled:cursor-not-allowed disabled:hover:bg-blue-600 flex items-center justify-center"
            >
              <Send className="w-4 h-4" strokeWidth={2} />
              {files.length > 1 ? `שליחת ${files.length} קבצים לעיבוד` : "שליחה לעיבוד"}
            </button>
          </div>
        )}

        {view === "loading" && (
          <div className="flex flex-col items-center gap-5 py-10 px-6 text-center">
            <div
              role="status"
              aria-label="מעבד..."
              className="w-12 h-12 rounded-full border-[3px] border-blue-100 border-t-blue-600"
              style={{ animation: "spin-slow 0.8s linear infinite" }}
            />
            <div className="text-base font-medium text-textc-primary">מעבד את הקובץ...</div>
            <div className="text-[13px] text-textc-secondary">{LOADING_STEPS[stepIdx]}</div>
            <div className="w-full h-1 bg-border rounded-sm overflow-hidden mt-1">
              <div
                className="h-full bg-blue-600 rounded-sm transition-[width] duration-300"
                style={{ width: `${Math.round(progress)}%` }}
              />
            </div>
          </div>
        )}

        {view === "success" && (
          <div className="flex flex-col items-center gap-4 py-10 px-6 text-center">
            <div className="w-14 h-14 bg-teal-50 rounded-full flex items-center justify-center text-teal-600">
              <CircleCheck className="w-7 h-7" />
            </div>
            <div className="text-base font-medium text-textc-primary">העיבוד הושלם בהצלחה!</div>
            <div className="text-[13px] text-textc-secondary">הקובץ נוצר עם גיליונות הסטטוסים המעודכנים</div>
            <button
              type="button"
              onClick={redownload}
              className="inline-flex items-center gap-1.5 mt-1 py-2.5 px-6 bg-teal-50 text-teal-800 border border-teal-400 rounded-md text-sm font-medium transition-colors hover:bg-teal-400 hover:text-white"
            >
              <Download className="w-4 h-4" />
              הורד קובץ
            </button>
            <button
              type="button"
              onClick={reset}
              className="text-[13px] text-textc-secondary underline hover:text-textc-primary mt-1"
            >
              עיבוד קובץ נוסף
            </button>
          </div>
        )}
      </section>
      <footer className="fixed bottom-0 left-0 right-0 py-3 text-center text-[11px] tracking-wide text-textc-secondary/70 bg-gradient-to-t from-background via-background/80 to-transparent pointer-events-none">
        SOX Summary of Controls · v0.1
      </footer>
    </main>
  );
}
