import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { QRCodeCanvas, QRCodeSVG } from "qrcode.react";
import jsPDF from "jspdf";
import QRCodeLib from "qrcode";

export const Route = createFileRoute("/")({
  component: Index,
});

type Item = {
  id: string;
  codigo: string;
  produto: string;
  qtdEsperada: string;
  qrValue: string;
};

type Warning = { type: string; detail: string };
type SheetRow = Array<unknown>;

const STORAGE_READ = "qr-read-ids";
const STORAGE_ITEMS = "qr-items";

function isHeaderRow(row: SheetRow) {
  const joined = row.map((c) => String(c ?? "").toLowerCase()).join(" ");
  return /c[oó]digo|produto|qtde|quantidade/.test(joined);
}

function parseSheet(file: File): Promise<{ items: Item[]; warnings: Warning[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array", cellText: true, raw: false });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<SheetRow>(ws, {
          header: 1,
          raw: false,
          defval: "",
          blankrows: false,
        });
        const items: Item[] = [];
        const warnings: Warning[] = [];
        const seen = new Map<string, number>();
        const start = rows.length > 0 && isHeaderRow(rows[0]) ? 1 : 0;
        for (let i = start; i < rows.length; i++) {
          const row = rows[i] || [];
          const codigo = String(row[0] ?? "").trim();
          const produto = String(row[1] ?? "").trim();
          const qtdEsperada = String(row[4] ?? "").trim();
          if (!codigo && !produto && !qtdEsperada) continue;
          if (!codigo) warnings.push({ type: "missing-code", detail: `Linha ${i + 1} sem código` });
          if (!produto)
            warnings.push({ type: "missing-product", detail: `Linha ${i + 1} sem produto` });
          if (!qtdEsperada)
            warnings.push({ type: "missing-qty", detail: `Linha ${i + 1} sem quantidade` });
          if (codigo) {
            if (seen.has(codigo)) {
              warnings.push({
                type: "duplicate",
                detail: `Código duplicado: ${codigo} (linhas ${seen.get(codigo)} e ${i + 1})`,
              });
            } else seen.set(codigo, i + 1);
          }
          items.push({
            id: `${i}-${codigo || produto}`,
            codigo,
            produto,
            qtdEsperada,
            qrValue: codigo ? codigo + "0001" : "",
          });
        }
        resolve({ items, warnings });
      } catch (err) {
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function Index() {
  const [items, setItems] = useState<Item[]>([]);
  const [warnings, setWarnings] = useState<Warning[]>([]);
  const [fileName, setFileName] = useState("");
  const [query, setQuery] = useState("");
  const [loading, setLoading] = useState(false);
  const [tab, setTab] = useState<"pendentes" | "lidos" | "todos">("pendentes");
  const [readIds, setReadIds] = useState<Set<string>>(new Set());

  // load from storage
  useEffect(() => {
    try {
      const r = localStorage.getItem(STORAGE_READ);
      if (r) setReadIds(new Set(JSON.parse(r)));
      const it = localStorage.getItem(STORAGE_ITEMS);
      if (it) {
        const parsed = JSON.parse(it);
        if (parsed?.items) {
          setItems(parsed.items);
          setFileName(parsed.fileName || "");
        }
      }
    } catch (error) {
      console.error("Erro ao carregar dados do armazenamento local.", error);
    }
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_READ, JSON.stringify(Array.from(readIds)));
  }, [readIds]);

  useEffect(() => {
    if (items.length) localStorage.setItem(STORAGE_ITEMS, JSON.stringify({ items, fileName }));
  }, [items, fileName]);

  function toggleRead(id: string) {
    setReadIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function markAll(read: boolean) {
    if (read) setReadIds(new Set(items.map((i) => i.id)));
    else setReadIds(new Set());
  }

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    return items.filter((i) => {
      const isRead = readIds.has(i.id);
      if (tab === "pendentes" && isRead) return false;
      if (tab === "lidos" && !isRead) return false;
      if (!q) return true;
      return i.codigo.toLowerCase().includes(q) || i.produto.toLowerCase().includes(q);
    });
  }, [items, query, tab, readIds]);

  const stats = useMemo(() => {
    const total = items.length;
    const qrCount = items.filter((i) => i.qrValue).length;
    const withQty = items.filter((i) => Number(i.qtdEsperada.replace(",", ".")) > 0).length;
    const pendentes = items.filter((i) => !readIds.has(i.id)).length;
    const lidos = items.length - pendentes;
    return { total, qrCount, withQty, pendentes, lidos };
  }, [items, readIds]);

  async function onFile(f: File | undefined | null) {
    if (!f) return;
    setLoading(true);
    try {
      const { items: parsed, warnings } = await parseSheet(f);
      if (!parsed.length) {
        alert(
          "Nenhum item encontrado na planilha. Verifique se a primeira aba contém as colunas A (Código), B (Produto) e E (Qtde Esperada).",
        );
      }
      setItems(parsed);
      setWarnings(warnings);
      setFileName(f.name);
      setReadIds(new Set());
    } catch (err) {
      console.error(err);
      alert("Erro ao ler planilha: " + (err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  function clearAll() {
    if (!confirm("Limpar todos os dados carregados?")) return;
    setItems([]);
    setWarnings([]);
    setFileName("");
    setReadIds(new Set());
    localStorage.removeItem(STORAGE_ITEMS);
    localStorage.removeItem(STORAGE_READ);
  }

  function exportCSV() {
    const header = ["Codigo", "Produto", "QtdEsperada", "QRCodeValue", "Lido"];
    const lines = [header.join(",")].concat(
      items.map((i) =>
        [
          i.codigo,
          `"${i.produto.replace(/"/g, '""')}"`,
          i.qtdEsperada,
          i.qrValue,
          readIds.has(i.id) ? "Sim" : "Não",
        ].join(","),
      ),
    );
    const blob = new Blob(["\uFEFF" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "qrcodes-almoxarifado.csv";
    a.click();
    URL.revokeObjectURL(url);
  }

  async function exportPDF() {
    const list = visible.filter((i) => i.qrValue);
    if (!list.length) {
      alert("Nada para exportar.");
      return;
    }

    const pdf = new jsPDF({ unit: "mm", format: "a4" });
    const pageW = 210,
      pageH = 297,
      margin = 8;
    const cols = 3,
      rows = 4;
    const cellW = (pageW - margin * 2) / cols;
    const cellH = (pageH - margin * 2) / rows;

    try {
      for (const [idx, it] of list.entries()) {
        const slot = idx % (cols * rows);
        if (idx > 0 && slot === 0) pdf.addPage();
        const c = slot % cols,
          r = Math.floor(slot / cols);
        const x = margin + c * cellW,
          y = margin + r * cellH;
        const img = await QRCodeLib.toDataURL(it.qrValue, {
          width: 320,
          margin: 1,
          errorCorrectionLevel: "M",
        });
        pdf.setDrawColor(180);
        pdf.roundedRect(x + 1, y + 1, cellW - 2, cellH - 2, 2, 2);
        const qrSize = Math.min(cellW, cellH) * 0.5;
        pdf.addImage(img, "PNG", x + (cellW - qrSize) / 2, y + 3, qrSize, qrSize);
        pdf.setFontSize(8);
        pdf.setTextColor(20);
        const ty = y + qrSize + 7;
        const productText = it.produto || "—";
        const prod = pdf.splitTextToSize(productText, cellW - 6);
        pdf.setFont("helvetica", "bold");
        pdf.text(`Código: ${it.codigo}`, x + 3, ty);
        pdf.setFont("helvetica", "normal");
        pdf.text(`Leitura: ${it.qrValue}`, x + 3, ty + 4);
        pdf.text(prod, x + 3, ty + 8);
        pdf.setFont("helvetica", "bold");
        pdf.text(`Qtd Esperada: ${it.qtdEsperada}`, x + 3, ty + 8 + prod.length * 3.5 + 2);
      }

      pdf.save("etiquetas-qrcodes.pdf");
    } catch (err) {
      console.error(err);
      alert("Erro ao gerar PDF: " + (err as Error).message);
    }
  }

  return (
    <div className="min-h-screen">
      <header className="no-print sticky top-0 z-20 border-b border-border/40 bg-background/70 backdrop-blur-xl">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between gap-4 flex-wrap">
          <div className="flex items-center gap-3">
            <div className="size-11 rounded-2xl bg-gradient-to-br from-primary via-primary to-accent grid place-items-center text-primary-foreground font-bold shadow-lg shadow-primary/40">
              <svg
                width="22"
                height="22"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2.2"
              >
                <rect x="3" y="3" width="7" height="7" rx="1" />
                <rect x="14" y="3" width="7" height="7" rx="1" />
                <rect x="3" y="14" width="7" height="7" rx="1" />
                <path d="M14 14h3v3h-3zM18 18h3v3h-3zM14 18h3v3h-3zM18 14h3v3h-3z" />
              </svg>
            </div>
            <div>
              <h1 className="text-base sm:text-lg font-bold leading-tight tracking-tight">
                Sistema de QR Code
              </h1>
              <p className="text-[11px] text-muted-foreground uppercase tracking-widest">
                Almoxarifado ADM
              </p>
            </div>
          </div>
          <div className="flex gap-2 flex-wrap">
            <Btn onClick={exportCSV} disabled={!items.length} variant="ghost">
              CSV
            </Btn>
            <Btn onClick={exportPDF} disabled={!items.length} variant="ghost">
              PDF
            </Btn>
            <Btn onClick={() => window.print()} disabled={!items.length}>
              Imprimir
            </Btn>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-8 space-y-6">
        {/* Upload */}
        <section className="no-print">
          <div className="rounded-3xl bg-card/60 border border-border/60 p-6 shadow-2xl shadow-black/30 backdrop-blur-sm relative overflow-hidden">
            <div className="absolute -top-24 -right-24 size-64 rounded-full bg-primary/20 blur-3xl pointer-events-none" />
            <div className="absolute -bottom-24 -left-24 size-64 rounded-full bg-accent/15 blur-3xl pointer-events-none" />
            <div className="relative flex flex-col md:flex-row items-center gap-5">
              <label className="flex-1 w-full flex items-center gap-4 border-2 border-dashed border-border/70 rounded-2xl p-5 cursor-pointer hover:border-primary/70 hover:bg-primary/5 transition group">
                <div className="size-14 shrink-0 rounded-2xl bg-gradient-to-br from-primary/30 to-accent/30 text-primary grid place-items-center group-hover:scale-105 transition">
                  <svg
                    width="24"
                    height="24"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                  >
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                    <polyline points="17 8 12 3 7 8" />
                    <line x1="12" y1="3" x2="12" y2="15" />
                  </svg>
                </div>
                <div className="flex-1 min-w-0">
                  <p className="font-semibold truncate">
                    {fileName || "Selecionar planilha .xlsx"}
                  </p>
                  <p className="text-xs text-muted-foreground mt-0.5">
                    Colunas esperadas: A=Código · B=Produto · E=Qtde Esperada
                  </p>
                </div>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={(e) => onFile(e.target.files?.[0])}
                />
                {loading && (
                  <span className="text-xs text-primary animate-pulse">Processando…</span>
                )}
              </label>
              {items.length > 0 && (
                <button
                  onClick={clearAll}
                  className="text-xs text-muted-foreground hover:text-destructive transition px-3 py-2"
                >
                  Limpar tudo
                </button>
              )}
            </div>
          </div>
        </section>

        {items.length > 0 && (
          <>
            {/* Stats */}
            <section className="no-print grid grid-cols-2 lg:grid-cols-5 gap-3">
              <StatCard label="Itens" value={stats.total} tone="primary" />
              <StatCard label="QR Codes" value={stats.qrCount} tone="accent" />
              <StatCard label="Com Qtd > 0" value={stats.withQty} tone="success" />
              <StatCard label="Pendentes" value={stats.pendentes} tone="warning" />
              <StatCard label="Lidos" value={stats.lidos} tone="muted" />
            </section>

            {/* Tabs + filtros */}
            <section className="no-print rounded-3xl bg-card/60 border border-border/60 backdrop-blur-sm overflow-hidden">
              <div className="p-4 flex flex-col sm:flex-row gap-3 items-stretch sm:items-center border-b border-border/40">
                <div className="inline-flex rounded-xl bg-secondary/60 p-1 self-start">
                  {(["pendentes", "lidos", "todos"] as const).map((t) => {
                    const count =
                      t === "pendentes"
                        ? stats.pendentes
                        : t === "lidos"
                          ? stats.lidos
                          : stats.total;
                    return (
                      <button
                        key={t}
                        onClick={() => setTab(t)}
                        className={`px-4 py-1.5 text-sm rounded-lg capitalize transition font-medium ${tab === t ? "bg-background text-foreground shadow" : "text-muted-foreground hover:text-foreground"}`}
                      >
                        {t} <span className="ml-1 text-xs opacity-70">({count})</span>
                      </button>
                    );
                  })}
                </div>
                <div className="flex-1 relative">
                  <svg
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground"
                    width="16"
                    height="16"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                  >
                    <circle cx="11" cy="11" r="8" />
                    <path d="m21 21-4.3-4.3" />
                  </svg>
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Buscar por código ou produto…"
                    className="w-full bg-input/50 border border-border/60 rounded-xl pl-10 pr-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-primary/60 transition"
                  />
                </div>
                <div className="flex gap-2">
                  <Btn variant="ghost" onClick={() => markAll(true)}>
                    Marcar todos lidos
                  </Btn>
                  <Btn variant="ghost" onClick={() => markAll(false)}>
                    Desmarcar
                  </Btn>
                </div>
              </div>

              {/* Cards grid */}
              <div className="p-4">
                {visible.length === 0 ? (
                  <div className="text-center py-16 text-muted-foreground">
                    <p className="text-sm">
                      {tab === "pendentes"
                        ? "🎉 Todos os itens foram lidos!"
                        : tab === "lidos"
                          ? "Nenhum item marcado como lido ainda."
                          : "Nenhum resultado."}
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                    {visible.map((it) => {
                      const isRead = readIds.has(it.id);
                      return (
                        <article
                          key={it.id}
                          className={`group relative rounded-2xl border p-4 transition-all duration-300 ${
                            isRead
                              ? "bg-success/5 border-success/30"
                              : "bg-secondary/30 border-border/60 hover:border-primary/60 hover:-translate-y-0.5 hover:shadow-xl hover:shadow-primary/10"
                          }`}
                        >
                          <button
                            onClick={() => toggleRead(it.id)}
                            title={isRead ? "Desmarcar como lido" : "Marcar como lido"}
                            className={`absolute top-3 right-3 size-7 rounded-full grid place-items-center border transition ${
                              isRead
                                ? "bg-success border-success text-success-foreground"
                                : "border-border/70 text-muted-foreground hover:border-primary hover:text-primary"
                            }`}
                          >
                            <svg
                              width="14"
                              height="14"
                              viewBox="0 0 24 24"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="3"
                            >
                              <polyline points="20 6 9 17 4 12" />
                            </svg>
                          </button>
                          <div className="flex justify-center mb-3">
                            <div className="bg-white p-2.5 rounded-xl shadow-inner">
                              {it.qrValue ? (
                                <QRCodeCanvas value={it.qrValue} size={120} level="M" />
                              ) : (
                                <div className="size-[120px] grid place-items-center text-xs text-muted-foreground">
                                  sem código
                                </div>
                              )}
                            </div>
                          </div>
                          <div className="space-y-1 text-sm">
                            <p className="font-mono text-base font-bold text-foreground">
                              {it.codigo || "—"}
                            </p>
                            <p className="text-xs text-accent font-mono truncate">{it.qrValue}</p>
                            <p className="font-medium text-foreground line-clamp-2 leading-snug min-h-[2.5rem]">
                              {it.produto || "—"}
                            </p>
                            <div className="flex items-center justify-between pt-2 border-t border-border/40 mt-2">
                              <span className="text-xs text-muted-foreground">Qtd Esperada</span>
                              <span className="font-bold text-primary">
                                {it.qtdEsperada || "—"}
                              </span>
                            </div>
                          </div>
                        </article>
                      );
                    })}
                  </div>
                )}
              </div>
            </section>

            {/* Warnings */}
            {warnings.length > 0 && (
              <details className="no-print rounded-2xl bg-warning/5 border border-warning/30 p-4">
                <summary className="cursor-pointer text-sm font-semibold text-warning">
                  Avisos ({warnings.length})
                </summary>
                <ul className="text-xs space-y-1 mt-3 max-h-48 overflow-auto text-muted-foreground">
                  {warnings.slice(0, 100).map((w, i) => (
                    <li key={i}>• {w.detail}</li>
                  ))}
                  {warnings.length > 100 && (
                    <li className="italic">…e mais {warnings.length - 100}</li>
                  )}
                </ul>
              </details>
            )}

            {/* Print-only labels (visible items only) */}
            <section className="hidden print:block">
              <div className="print-grid">
                {visible
                  .filter((i) => i.qrValue)
                  .map((it) => (
                    <div key={it.id} className="print-card p-3 text-center">
                      <div className="grid place-items-center">
                        <QRCodeSVG value={it.qrValue} size={140} level="M" />
                      </div>
                      <div className="mt-2 text-xs text-left space-y-0.5">
                        <p>
                          <b>Código:</b> {it.codigo}
                        </p>
                        <p>
                          <b>Leitura:</b> {it.qrValue}
                        </p>
                        <p>
                          <b>Produto:</b> {it.produto}
                        </p>
                        <p>
                          <b>Qtd Esperada:</b> {it.qtdEsperada}
                        </p>
                      </div>
                    </div>
                  ))}
              </div>
            </section>
          </>
        )}

        {!items.length && !loading && (
          <div className="no-print text-center text-muted-foreground py-20">
            <p className="text-sm">Importe uma planilha .xlsx para começar a gerar QR Codes.</p>
          </div>
        )}
      </main>
    </div>
  );
}

function Btn({
  children,
  onClick,
  disabled,
  variant = "primary",
}: {
  children: React.ReactNode;
  onClick?: React.ButtonHTMLAttributes<HTMLButtonElement>["onClick"];
  disabled?: boolean;
  variant?: "primary" | "ghost";
}) {
  const cls =
    variant === "primary"
      ? "bg-gradient-to-r from-primary to-accent text-primary-foreground shadow-lg shadow-primary/30 hover:opacity-90"
      : "bg-secondary/70 text-foreground hover:bg-secondary border border-border/60";
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className={`px-4 py-2 text-sm rounded-xl font-medium transition disabled:opacity-40 disabled:cursor-not-allowed ${cls}`}
    >
      {children}
    </button>
  );
}

function StatCard({
  label,
  value,
  tone,
}: {
  label: string;
  value: number;
  tone: "primary" | "accent" | "success" | "warning" | "muted";
}) {
  const tones = {
    primary: "from-primary/25 to-primary/5 border-primary/30",
    accent: "from-accent/25 to-accent/5 border-accent/30",
    success: "from-success/25 to-success/5 border-success/30",
    warning: "from-warning/25 to-warning/5 border-warning/30",
    muted: "from-muted/40 to-muted/10 border-border/60",
  }[tone];
  return (
    <div className={`rounded-2xl border bg-gradient-to-br ${tones} p-4 backdrop-blur-sm`}>
      <p className="text-[10px] uppercase tracking-widest text-muted-foreground font-semibold">
        {label}
      </p>
      <p className="text-2xl font-bold mt-1 text-foreground tabular-nums">{value}</p>
    </div>
  );
}
