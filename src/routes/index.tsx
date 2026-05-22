import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { QRCodeCanvas, QRCodeSVG } from "qrcode.react";
import jsPDF from "jspdf";
import { Linkedin, Mail, MessageCircleMore } from "lucide-react";
import QRCodeLib from "qrcode";
import { LoadingGlyphArt, PortalBackdropArt } from "@/components/game-art";

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

type Role = "adm" | "rh";
type Warning = { type: string; detail: string };
type BusyState = { title: string; detail: string };
type SheetRow = Array<unknown>;

const STORAGE_READ = "qr-read-ids";
const STORAGE_ITEMS = "qr-items";

const ROLE_LABELS: Record<Role, string> = {
  adm: "ADM",
  rh: "RH",
};

const ROLE_HELPERS: Record<Role, string> = {
  adm: "Planilha padrão: A=Código, B=Produto e E=Qtde Esperada.",
  rh: "Planilha do RH: o sistema tenta encontrar o código mesmo se a ordem das colunas mudar.",
};

const DEVELOPER_NAME = "Leonardo Gonçalves da Silva";
const DEVELOPER_LINKEDIN_URL = "https://www.linkedin.com/in/leonardo-g-silva-353228210/";
const DEVELOPER_WHATSAPP_URL = "https://wa.me/5521982111477";
const DEVELOPER_EMAIL_URL = "mailto:leonardogsilv1@gmail.com";

function storageKey(prefix: string, role: Role) {
  return `${prefix}-${role}`;
}

function cellText(value: unknown) {
  return String(value ?? "").trim();
}

function normalizeText(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function rowValues(row: SheetRow) {
  return row.map(cellText).filter(Boolean);
}

function findHeaderRow(rows: SheetRow[], role: Role) {
  const keywords =
    role === "adm"
      ? ["codigo", "produto", "qtde", "quantidade", "qtd"]
      : [
          "codigo",
          "matricula",
          "nome",
          "funcionario",
          "colaborador",
          "produto",
          "descricao",
          "setor",
          "cargo",
          "quantidade",
          "qtde",
          "qtd",
        ];

  const limit = Math.min(rows.length, 6);
  for (let i = 0; i < limit; i++) {
    const joined = normalizeText(rows[i].map(cellText).join(" "));
    if (keywords.some((keyword) => joined.includes(keyword))) return i;
  }
  return -1;
}

function findColumnIndex(row: SheetRow, keywords: string[]) {
  return row.findIndex((cell) => {
    const normalized = normalizeText(cellText(cell));
    return keywords.some((keyword) => normalized.includes(keyword));
  });
}

function detectColumns(row: SheetRow, role: Role) {
  const codeKeywords =
    role === "adm"
      ? ["codigo", "cod"]
      : ["codigo", "cod", "matricula", "registro", "chapa", "badge", "id", "funcional"];
  const productKeywords =
    role === "adm"
      ? ["produto", "descricao"]
      : ["nome", "produto", "funcionario", "colaborador", "descricao", "setor", "cargo"];
  const qtyKeywords = ["qtde", "quantidade", "qtd"];

  const codigo = findColumnIndex(row, codeKeywords);
  const produto = findColumnIndex(row, productKeywords);
  const quantidade = findColumnIndex(row, qtyKeywords);

  return {
    codigo: codigo >= 0 ? codigo : role === "adm" ? 0 : -1,
    produto: produto >= 0 ? produto : role === "adm" ? 1 : -1,
    quantidade: quantidade >= 0 ? quantidade : role === "adm" ? 4 : -1,
  };
}

function scoreCodeCandidate(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return -1;
  let score = 0;
  if (/\d/.test(trimmed)) score += 3;
  if (!/\s/.test(trimmed)) score += 1;
  if (/^[A-Za-z0-9._/-]+$/.test(trimmed)) score += 1;
  if (trimmed.length <= 24) score += 1;
  return score;
}

function pickLikelyCode(values: string[]) {
  return [...values].sort((a, b) => scoreCodeCandidate(b) - scoreCodeCandidate(a))[0] ?? "";
}

function parseStoredPayload(rawItems: string | null, rawReadIds: string | null) {
  let items: Item[] = [];
  let fileName = "";
  let warnings: Warning[] = [];
  let readIds = new Set<string>();

  if (rawItems) {
    const parsed = JSON.parse(rawItems) as {
      items?: Item[];
      fileName?: string;
      warnings?: Warning[];
    };
    if (Array.isArray(parsed.items)) items = parsed.items;
    if (typeof parsed.fileName === "string") fileName = parsed.fileName;
    if (Array.isArray(parsed.warnings)) warnings = parsed.warnings;
  }

  if (rawReadIds) {
    const parsed = JSON.parse(rawReadIds) as unknown;
    if (Array.isArray(parsed)) readIds = new Set(parsed.map((id) => String(id)));
  }

  return { items, fileName, warnings, readIds };
}

function loadRoleState(role: Role) {
  const rawItems = localStorage.getItem(storageKey(STORAGE_ITEMS, role));
  const rawReadIds = localStorage.getItem(storageKey(STORAGE_READ, role));
  return parseStoredPayload(rawItems, rawReadIds);
}

function parseSheet(file: File, role: Role): Promise<{ items: Item[]; warnings: Warning[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error);
    reader.onload = async (e) => {
      try {
        await new Promise((resolveWait) => setTimeout(resolveWait, 0));

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
        const headerRow = findHeaderRow(rows, role);
        const columns = detectColumns(headerRow >= 0 ? rows[headerRow] || [] : [], role);
        const start = headerRow >= 0 ? headerRow + 1 : 0;

        for (let i = start; i < rows.length; i++) {
          const row = rows[i] || [];
          const values = rowValues(row);
          const codigo = columns.codigo >= 0 ? cellText(row[columns.codigo]) : "";
          const produto = columns.produto >= 0 ? cellText(row[columns.produto]) : "";
          const qtdEsperada = columns.quantidade >= 0 ? cellText(row[columns.quantidade]) : "";
          const resolvedCodigo =
            codigo || (role === "rh" ? pickLikelyCode(values) : values[0] || "");
          const resolvedProduto =
            produto || values.find((value) => value !== resolvedCodigo) || values[1] || "";
          const resolvedQtd = qtdEsperada || (role === "adm" ? values[2] || "" : "");

          if (!resolvedCodigo && !resolvedProduto && !resolvedQtd) continue;

          if (!resolvedCodigo) {
            warnings.push({
              type: "missing-code",
              detail: `Linha ${i + 1} sem código`,
            });
          }

          if (role === "adm") {
            if (!resolvedProduto) {
              warnings.push({
                type: "missing-product",
                detail: `Linha ${i + 1} sem produto`,
              });
            }
            if (!resolvedQtd) {
              warnings.push({
                type: "missing-qty",
                detail: `Linha ${i + 1} sem quantidade`,
              });
            }
          }

          if (resolvedCodigo) {
            if (seen.has(resolvedCodigo)) {
              warnings.push({
                type: "duplicate",
                detail: `Código duplicado: ${resolvedCodigo} (linhas ${seen.get(resolvedCodigo)} e ${i + 1})`,
              });
            } else {
              seen.set(resolvedCodigo, i + 1);
            }
          }

          items.push({
            id: `${role}-${i}-${resolvedCodigo || resolvedProduto}`,
            codigo: resolvedCodigo,
            produto: resolvedProduto,
            qtdEsperada: resolvedQtd,
            qrValue: resolvedCodigo ? `${resolvedCodigo}0001` : "",
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
  const [role, setRole] = useState<Role | null>(null);
  const [items, setItems] = useState<Item[]>([]);
  const [warnings, setWarnings] = useState<Warning[]>([]);
  const [fileName, setFileName] = useState("");
  const [query, setQuery] = useState("");
  const [busy, setBusy] = useState<BusyState | null>(null);
  const [busyVisible, setBusyVisible] = useState(false);
  const [tab, setTab] = useState<"pendentes" | "lidos" | "todos">("pendentes");
  const [readIds, setReadIds] = useState<Set<string>>(new Set());

  useEffect(() => {
    if (!busy) {
      setBusyVisible(false);
      return;
    }

    const timer = window.setTimeout(() => setBusyVisible(true), 300);
    return () => window.clearTimeout(timer);
  }, [busy]);

  useEffect(() => {
    if (!role) return;
    localStorage.setItem(storageKey(STORAGE_READ, role), JSON.stringify(Array.from(readIds)));
  }, [role, readIds]);

  useEffect(() => {
    if (!role) return;
    localStorage.setItem(
      storageKey(STORAGE_ITEMS, role),
      JSON.stringify({ items, fileName, warnings }),
    );
  }, [role, items, fileName, warnings]);

  function selectRole(nextRole: Role) {
    try {
      const loaded = loadRoleState(nextRole);
      setItems(loaded.items);
      setWarnings(loaded.warnings);
      setFileName(loaded.fileName);
      setReadIds(loaded.readIds);
    } catch (error) {
      console.error("Erro ao carregar dados do perfil.", error);
      setItems([]);
      setWarnings([]);
      setFileName("");
      setReadIds(new Set());
    }

    setQuery("");
    setTab("pendentes");
    setBusy(null);
    setRole(nextRole);
  }

  function resetRole() {
    setRole(null);
    setQuery("");
    setTab("pendentes");
    setBusy(null);
  }

  function toggleRead(id: string) {
    setReadIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function markAll(read: boolean) {
    if (read) setReadIds(new Set(items.map((item) => item.id)));
    else setReadIds(new Set());
  }

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    return items.filter((item) => {
      const isRead = readIds.has(item.id);
      if (tab === "pendentes" && isRead) return false;
      if (tab === "lidos" && !isRead) return false;
      if (!q) return true;
      return item.codigo.toLowerCase().includes(q) || item.produto.toLowerCase().includes(q);
    });
  }, [items, query, tab, readIds]);

  const stats = useMemo(() => {
    const total = items.length;
    const qrCount = items.filter((item) => item.qrValue).length;
    const withQty = items.filter((item) => Number(item.qtdEsperada.replace(",", ".")) > 0).length;
    const pendentes = items.filter((item) => !readIds.has(item.id)).length;
    const lidos = items.length - pendentes;
    return { total, qrCount, withQty, pendentes, lidos };
  }, [items, readIds]);

  async function onFile(file: File | undefined | null) {
    if (!file || !role) return;

    setBusy({
      title: `Lendo planilha ${ROLE_LABELS[role]}`,
      detail: "Organizando as linhas e preparando os QR Codes.",
    });

    try {
      await new Promise((resolveWait) => setTimeout(resolveWait, 0));
      const { items: parsed, warnings: nextWarnings } = await parseSheet(file, role);

      if (!parsed.length) {
        alert(
          role === "adm"
            ? "Nenhum item encontrado na planilha. Verifique se a primeira aba contém as colunas A (Código), B (Produto) e E (Qtde Esperada)."
            : "Nenhum item encontrado na planilha. Verifique se a aba contém os códigos que você quer transformar em QR Code.",
        );
      }

      setItems(parsed);
      setWarnings(nextWarnings);
      setFileName(file.name);
      setReadIds(new Set());
      setTab("pendentes");
      setQuery("");
    } catch (error) {
      console.error(error);
      alert("Erro ao ler planilha: " + (error as Error).message);
    } finally {
      setBusy(null);
    }
  }

  function clearAll() {
    if (!role) return;
    if (!confirm("Limpar todos os dados carregados?")) return;

    setItems([]);
    setWarnings([]);
    setFileName("");
    setReadIds(new Set());
    setQuery("");
    setTab("pendentes");
    localStorage.removeItem(storageKey(STORAGE_ITEMS, role));
    localStorage.removeItem(storageKey(STORAGE_READ, role));
  }

  function exportCSV() {
    if (!role) return;

    const header = ["Codigo", "Produto", "QtdEsperada", "QRCodeValue", "Lido"];
    const lines = [header.join(",")].concat(
      items.map((item) =>
        [
          item.codigo,
          `"${item.produto.replace(/"/g, '""')}"`,
          item.qtdEsperada,
          item.qrValue,
          readIds.has(item.id) ? "Sim" : "Nao",
        ].join(","),
      ),
    );
    const blob = new Blob(["\uFEFF" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `qrcodes-${role}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function exportPDF() {
    if (!role) return;

    const list = visible.filter((item) => item.qrValue);
    if (!list.length) {
      alert("Nada para exportar.");
      return;
    }

    setBusy({
      title: `Gerando PDF ${ROLE_LABELS[role]}`,
      detail: "Montando as etiquetas e convertendo os QR Codes.",
    });

    try {
      await new Promise((resolveWait) => setTimeout(resolveWait, 0));

      const pdf = new jsPDF({ unit: "mm", format: "a4" });
      const pageW = 210,
        pageH = 297,
        margin = 8;
      const cols = 3,
        rowsPerPage = 4;
      const cellW = (pageW - margin * 2) / cols;
      const cellH = (pageH - margin * 2) / rowsPerPage;

      for (const [idx, item] of list.entries()) {
        const slot = idx % (cols * rowsPerPage);
        if (idx > 0 && slot === 0) pdf.addPage();
        const c = slot % cols;
        const r = Math.floor(slot / cols);
        const x = margin + c * cellW;
        const y = margin + r * cellH;
        const img = await QRCodeLib.toDataURL(item.qrValue, {
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
        const productText = item.produto || "—";
        const prod = pdf.splitTextToSize(productText, cellW - 6);
        pdf.setFont("helvetica", "bold");
        pdf.text(`Codigo: ${item.codigo}`, x + 3, ty);
        pdf.setFont("helvetica", "normal");
        pdf.text(`Leitura: ${item.qrValue}`, x + 3, ty + 4);
        pdf.text(prod, x + 3, ty + 8);
        pdf.setFont("helvetica", "bold");
        pdf.text(`Qtd Esperada: ${item.qtdEsperada}`, x + 3, ty + 8 + prod.length * 3.5 + 2);
      }

      pdf.save(`etiquetas-qrcodes-${role}.pdf`);
    } catch (error) {
      console.error(error);
      alert("Erro ao gerar PDF: " + (error as Error).message);
    } finally {
      setBusy(null);
    }
  }

  if (!role) {
    return <RoleGate onSelect={selectRole} />;
  }

  return (
    <div className="flex min-h-screen flex-col">
      <header className="no-print sticky top-0 z-20 border-b border-border/40 bg-background/70 backdrop-blur-xl">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-4 px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="grid size-11 place-items-center rounded-2xl bg-gradient-to-br from-primary via-primary to-accent font-bold text-primary-foreground shadow-lg shadow-primary/40">
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
              <h1 className="text-base font-bold leading-tight tracking-tight sm:text-lg">
                Sistema de QR Code
              </h1>
              <p className="text-[11px] uppercase tracking-widest text-muted-foreground">
                Almoxarifado {ROLE_LABELS[role]}
              </p>
            </div>
          </div>
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={resetRole}
              className="inline-flex items-center gap-2 rounded-full border border-success/30 bg-success/5 px-4 py-2 text-sm font-medium text-success shadow-sm transition hover:bg-success/10"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="#f5b640" aria-hidden="true">
                <path d="M12 2.8l2.71 5.49 6.06.88-4.39 4.28 1.04 6.03L12 16.95 6.58 19.48l1.04-6.03-4.39-4.28 6.06-.88L12 2.8z" />
              </svg>
              <span>{ROLE_LABELS[role]}</span>
              <span className="text-xs opacity-70">trocar</span>
            </button>
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

      <main className="mx-auto flex-1 max-w-7xl space-y-6 px-6 py-8">
        <section className="no-print">
          <div className="rounded-3xl border border-border/60 bg-card/60 p-6 shadow-2xl shadow-black/30 backdrop-blur-sm relative overflow-hidden">
            <div className="pointer-events-none absolute -right-24 -top-24 size-64 rounded-full bg-primary/18 blur-3xl" />
            <div className="pointer-events-none absolute -bottom-24 -left-24 size-64 rounded-full bg-accent/14 blur-3xl" />
            <div className="relative flex flex-col gap-5 md:flex-row md:items-center">
              <label className="flex-1 w-full flex items-center gap-4 border-2 border-dashed border-border/70 rounded-2xl p-5 cursor-pointer hover:border-primary/70 hover:bg-primary/5 transition group">
                <div className="grid size-14 shrink-0 place-items-center rounded-2xl bg-gradient-to-br from-primary/30 to-accent/30 text-primary transition group-hover:scale-105">
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
                <div className="min-w-0 flex-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <p className="truncate font-semibold">
                      {fileName || `Selecionar planilha ${ROLE_LABELS[role]}.xlsx`}
                    </p>
                    <span className="rounded-full border border-primary/20 bg-primary/10 px-2.5 py-0.5 text-[10px] font-semibold uppercase tracking-[0.25em] text-primary">
                      {ROLE_LABELS[role]}
                    </span>
                  </div>
                  <p className="mt-0.5 text-xs text-muted-foreground">{ROLE_HELPERS[role]}</p>
                </div>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={(e) => onFile(e.target.files?.[0])}
                />
                {busy && <span className="text-xs text-primary animate-pulse">Processando...</span>}
              </label>
              {items.length > 0 && (
                <button
                  onClick={clearAll}
                  className="px-3 py-2 text-xs text-muted-foreground transition hover:text-destructive"
                >
                  Limpar tudo
                </button>
              )}
            </div>
          </div>
        </section>

        {items.length > 0 && (
          <>
            <section className="no-print grid grid-cols-2 gap-3 lg:grid-cols-5">
              <StatCard label="Itens" value={stats.total} tone="primary" />
              <StatCard label="QR Codes" value={stats.qrCount} tone="accent" />
              <StatCard label="Com Qtd > 0" value={stats.withQty} tone="success" />
              <StatCard label="Pendentes" value={stats.pendentes} tone="warning" />
              <StatCard label="Lidos" value={stats.lidos} tone="muted" />
            </section>

            <section className="no-print rounded-3xl bg-card/60 border border-border/60 backdrop-blur-sm overflow-hidden">
              <div className="flex flex-col gap-3 border-b border-border/40 p-4 sm:flex-row sm:items-center">
                <div className="inline-flex self-start rounded-[18px] bg-secondary/60 p-1">
                  {(["pendentes", "lidos", "todos"] as const).map((currentTab) => {
                    const count =
                      currentTab === "pendentes"
                        ? stats.pendentes
                        : currentTab === "lidos"
                          ? stats.lidos
                          : stats.total;
                    return (
                      <button
                        key={currentTab}
                        onClick={() => setTab(currentTab)}
                        className={`rounded-[14px] px-4 py-1.5 text-sm font-medium capitalize transition ${
                          tab === currentTab
                            ? "bg-background text-foreground shadow-[0_12px_30px_rgba(8,32,50,0.12)]"
                            : "text-muted-foreground hover:text-foreground"
                        }`}
                      >
                        {currentTab} <span className="ml-1 text-xs opacity-70">({count})</span>
                      </button>
                    );
                  })}
                </div>
                <div className="relative flex-1">
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
                    placeholder="Buscar por código ou produto..."
                    className="w-full rounded-xl border border-border/60 bg-input/50 py-2.5 pl-10 pr-4 text-sm transition focus:outline-none focus:ring-2 focus:ring-primary/60"
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

              <div className="p-4">
                {visible.length === 0 ? (
                  <div className="py-16 text-center text-muted-foreground">
                    <p className="text-sm">
                      {tab === "pendentes"
                        ? "Todos os itens foram lidos!"
                        : tab === "lidos"
                          ? "Nenhum item marcado como lido ainda."
                          : "Nenhum resultado."}
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
                    {visible.map((item) => {
                      const isRead = readIds.has(item.id);
                      return (
                        <article
                          key={item.id}
                          className={`group relative rounded-2xl border p-4 transition-all duration-300 ${
                            isRead
                              ? "border-success/30 bg-success/5"
                              : "border-border/60 bg-secondary/30 hover:-translate-y-0.5 hover:border-primary/60 hover:shadow-xl hover:shadow-primary/10"
                          }`}
                        >
                          <button
                            onClick={() => toggleRead(item.id)}
                            title={isRead ? "Desmarcar como lido" : "Marcar como lido"}
                            className={`absolute right-3 top-3 grid size-7 place-items-center rounded-full border transition ${
                              isRead
                                ? "border-success bg-success text-success-foreground"
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
                          <div className="mb-3 flex justify-center">
                            <div className="rounded-xl bg-white p-2.5 shadow-inner">
                              {item.qrValue ? (
                                <QRCodeCanvas value={item.qrValue} size={120} level="M" />
                              ) : (
                                <div className="grid size-[120px] place-items-center text-xs text-muted-foreground">
                                  sem código
                                </div>
                              )}
                            </div>
                          </div>
                          <div className="space-y-1 text-sm">
                            <p className="font-mono text-base font-bold text-foreground">
                              {item.codigo || "—"}
                            </p>
                            <p className="truncate font-mono text-xs text-accent">{item.qrValue}</p>
                            <p className="min-h-[2.5rem] leading-snug font-medium text-foreground line-clamp-2">
                              {item.produto || "—"}
                            </p>
                            <div className="mt-2 flex items-center justify-between border-t border-border/40 pt-2">
                              <span className="text-xs text-muted-foreground">Qtd Esperada</span>
                              <span className="font-bold text-primary">
                                {item.qtdEsperada || "—"}
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

            {warnings.length > 0 && (
              <details className="no-print rounded-2xl bg-warning/5 border border-warning/30 p-4">
                <summary className="cursor-pointer text-sm font-semibold text-warning">
                  Avisos ({warnings.length})
                </summary>
                <ul className="mt-3 max-h-48 space-y-1 overflow-auto text-xs text-muted-foreground">
                  {warnings.slice(0, 100).map((warning, i) => (
                    <li key={i}>• {warning.detail}</li>
                  ))}
                  {warnings.length > 100 && (
                    <li className="italic">...e mais {warnings.length - 100}</li>
                  )}
                </ul>
              </details>
            )}

            <section className="hidden print:block">
              <div className="print-grid">
                {visible
                  .filter((item) => item.qrValue)
                  .map((item) => (
                    <div key={item.id} className="print-card p-3 text-center">
                      <div className="grid place-items-center">
                        <QRCodeSVG value={item.qrValue} size={140} level="M" />
                      </div>
                      <div className="mt-2 space-y-0.5 text-left text-xs">
                        <p>
                          <b>Codigo:</b> {item.codigo}
                        </p>
                        <p>
                          <b>Leitura:</b> {item.qrValue}
                        </p>
                        <p>
                          <b>Produto:</b> {item.produto}
                        </p>
                        <p>
                          <b>Qtd Esperada:</b> {item.qtdEsperada}
                        </p>
                      </div>
                    </div>
                  ))}
              </div>
            </section>
          </>
        )}

        {!items.length && !busy && (
          <div className="no-print py-20 text-center text-muted-foreground">
            <p className="text-sm">Importe uma planilha .xlsx para começar a gerar QR Codes.</p>
          </div>
        )}
      </main>

      <footer className="no-print border-t border-border/30 bg-background/75 backdrop-blur">
        <div className="mx-auto flex max-w-7xl flex-col gap-3 px-6 py-3 text-[11px] tracking-[0.14em] text-slate-500 sm:flex-row sm:items-center sm:justify-between">
          <span className="uppercase">
            Desenvolvido por <span className="font-semibold text-slate-700">{DEVELOPER_NAME}</span>
          </span>
          <div className="flex flex-wrap items-center justify-center gap-2">
            <a
              href={DEVELOPER_LINKEDIN_URL}
              target="_blank"
              rel="noreferrer"
              className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white/80 px-3 py-1.5 font-semibold uppercase text-slate-500 shadow-sm transition hover:border-[#0a66c2]/30 hover:text-[#0a66c2] hover:shadow-md"
            >
              <Linkedin className="size-3.5" aria-hidden="true" />
              LinkedIn
            </a>
            <a
              href={DEVELOPER_WHATSAPP_URL}
              target="_blank"
              rel="noreferrer"
              className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white/80 px-3 py-1.5 font-semibold uppercase text-slate-500 shadow-sm transition hover:border-emerald-400/30 hover:text-emerald-600 hover:shadow-md"
            >
              <MessageCircleMore className="size-3.5" aria-hidden="true" />
              WhatsApp
            </a>
            <a
              href={DEVELOPER_EMAIL_URL}
              className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white/80 px-3 py-1.5 font-semibold uppercase text-slate-500 shadow-sm transition hover:border-rose-400/30 hover:text-rose-600 hover:shadow-md"
            >
              <Mail className="size-3.5" aria-hidden="true" />
              E-mail
            </a>
          </div>
        </div>
      </footer>

      {busy && busyVisible && <LoadingOverlay title={busy.title} detail={busy.detail} />}
    </div>
  );
}

function RoleGate({ onSelect }: { onSelect: (role: Role) => void }) {
  return <SimpleRoleGate onSelect={onSelect} />;

  return (
    <div className="relative min-h-screen overflow-hidden bg-[radial-gradient(circle_at_top_left,rgba(14,116,144,0.18),transparent_30%),radial-gradient(circle_at_80%_20%,rgba(30,136,229,0.16),transparent_24%),radial-gradient(circle_at_bottom,rgba(192,132,26,0.15),transparent_28%),linear-gradient(180deg,#f9fbff_0%,#edf4fb_56%,#f3f7f9_100%)]">
      <PortalBackdropArt className="pointer-events-none absolute left-1/2 top-[-8%] h-[120%] w-[120%] -translate-x-1/2 opacity-[0.18] blur-3xl" />
      <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(circle_at_20%_20%,rgba(255,255,255,0.72),transparent_18%),radial-gradient(circle_at_80%_0%,rgba(255,255,255,0.6),transparent_14%)]" />
      <div className="pointer-events-none absolute left-[-8rem] top-24 size-[26rem] rounded-full bg-primary/12 blur-3xl animate-[drift-slow_18s_ease-in-out_infinite]" />
      <div className="pointer-events-none absolute right-[-9rem] top-12 size-[24rem] rounded-full bg-accent/12 blur-3xl animate-[drift-slow_20s_ease-in-out_infinite_reverse]" />
      <div className="pointer-events-none absolute bottom-[-10rem] left-1/4 size-[30rem] rounded-full bg-warning/10 blur-3xl animate-[drift-slow_22s_ease-in-out_infinite]" />

      <main className="relative mx-auto flex min-h-screen max-w-7xl items-center px-6 py-10">
        <section className="grid w-full items-center gap-8 lg:grid-cols-[1.08fr_.92fr]">
          <div className="relative space-y-8">
            <div className="flex flex-wrap gap-3">
              <span className="game-chip inline-flex items-center gap-2 rounded-full px-4 py-2 text-[10px] font-semibold uppercase tracking-[0.34em] text-primary">
                <span className="size-2 rounded-full bg-primary animate-pulse" />
                Entrada mágica
              </span>
              <span className="game-chip rounded-full px-4 py-2 text-[10px] font-semibold uppercase tracking-[0.34em] text-accent">
                ADM + RH
              </span>
            </div>

            <div className="space-y-5">
              <p className="font-display text-sm uppercase tracking-[0.4em] text-muted-foreground">
                Bem vindo ao painel
              </p>
              <h1 className="max-w-3xl font-display text-5xl font-black leading-[0.95] tracking-tight text-foreground sm:text-6xl lg:text-7xl">
                Você é ADM ou RH?
              </h1>
              <p className="max-w-2xl text-base leading-relaxed text-muted-foreground sm:text-lg">
                Escolha seu portal e siga em frente. A experiência ganhou cara de jogo, mas a função
                continua a mesma: ler a planilha, achar o código e gerar os QR Codes.
              </p>
            </div>

            <div className="grid gap-3 sm:grid-cols-3">
              <div className="game-chip rounded-[22px] p-4">
                <p className="text-[10px] uppercase tracking-[0.32em] text-muted-foreground">
                  Passo 1
                </p>
                <p className="mt-2 font-display text-lg font-bold text-foreground">
                  Escolhe o perfil
                </p>
              </div>
              <div className="game-chip rounded-[22px] p-4">
                <p className="text-[10px] uppercase tracking-[0.32em] text-muted-foreground">
                  Passo 2
                </p>
                <p className="mt-2 font-display text-lg font-bold text-foreground">
                  Carrega a planilha
                </p>
              </div>
              <div className="game-chip rounded-[22px] p-4">
                <p className="text-[10px] uppercase tracking-[0.32em] text-muted-foreground">
                  Passo 3
                </p>
                <p className="mt-2 font-display text-lg font-bold text-foreground">Gera o QR</p>
              </div>
            </div>

            <div className="grid gap-4 md:grid-cols-2">
              <RoleChoiceCard
                title="Sou ADM"
                eyebrow="Planilha direta"
                description="Usa a estrutura padrão: código, produto e quantidade esperada já em ordem."
                tone="adm"
                onClick={() => onSelect("adm")}
              />
              <RoleChoiceCard
                title="Sou RH"
                eyebrow="Planilha flexível"
                description="Aceita planilhas mais soltas e tenta achar o código mesmo quando as colunas mudam."
                tone="rh"
                onClick={() => onSelect("rh")}
              />
            </div>

            <p className="max-w-2xl text-sm leading-relaxed text-muted-foreground">
              O modo ADM usa a estrutura padrão da planilha. O modo RH é mais flexível e tenta
              descobrir o código mesmo quando as colunas vêm embaralhadas.
            </p>
          </div>

          <div className="relative">
            <div className="game-panel relative overflow-hidden rounded-[36px] p-5 sm:p-6">
              <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(145deg,rgba(255,255,255,0.6),rgba(255,255,255,0.08))]" />
              <PortalBackdropArt className="pointer-events-none absolute left-1/2 top-1/2 h-[125%] w-[125%] -translate-x-1/2 -translate-y-1/2 opacity-20 blur-2xl" />
              <div className="pointer-events-none absolute -right-20 -top-24 size-56 rounded-full bg-primary/20 blur-3xl" />
              <div className="pointer-events-none absolute -bottom-24 -left-16 size-56 rounded-full bg-accent/16 blur-3xl" />

              <div className="relative space-y-5">
                <div className="flex items-start justify-between gap-4">
                  <div>
                    <p className="text-[10px] font-semibold uppercase tracking-[0.4em] text-muted-foreground">
                      Visão geral
                    </p>
                    <h2 className="font-display mt-2 text-3xl font-black text-foreground sm:text-4xl">
                      Portal de entrada
                    </h2>
                  </div>
                  <span className="game-chip rounded-full px-3 py-1 text-[10px] font-semibold uppercase tracking-[0.3em] text-success">
                    Ao vivo
                  </span>
                </div>

                <div className="grid gap-3 sm:grid-cols-2">
                  <div className="rounded-[24px] border border-primary/20 bg-primary/5 p-4 shadow-[0_18px_40px_rgba(8,32,50,0.08)]">
                    <p className="text-[10px] font-semibold uppercase tracking-[0.34em] text-primary">
                      ADM
                    </p>
                    <p className="mt-2 font-display text-xl font-bold text-foreground">
                      Estrutura padrão
                    </p>
                    <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                      Código, produto e quantidade já no formato esperado.
                    </p>
                  </div>

                  <div className="rounded-[24px] border border-accent/20 bg-accent/5 p-4 shadow-[0_18px_40px_rgba(8,32,50,0.08)]">
                    <p className="text-[10px] font-semibold uppercase tracking-[0.34em] text-accent">
                      RH
                    </p>
                    <p className="mt-2 font-display text-xl font-bold text-foreground">
                      Leitura adaptativa
                    </p>
                    <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                      O sistema tenta descobrir o código mesmo com colunas diferentes.
                    </p>
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-[140px_1fr] md:items-center rounded-[28px] border border-border/60 bg-background/75 p-4">
                  <div className="mx-auto w-full max-w-[140px]">
                    <LoadingGlyphArt className="w-full opacity-90 drop-shadow-[0_18px_32px_rgba(8,32,50,0.15)]" />
                  </div>
                  <div className="space-y-3">
                    <p className="font-display text-2xl font-bold text-foreground">
                      Loading inteligente
                    </p>
                    <p className="text-sm leading-relaxed text-muted-foreground">
                      Se a planilha ou a geração demorar, a interface mostra uma animação elegante
                      para manter tudo claro.
                    </p>
                    <div className="flex flex-wrap gap-2">
                      <span className="game-chip rounded-full px-3 py-1 text-[10px] font-semibold uppercase tracking-[0.28em] text-primary">
                        Leitura
                      </span>
                      <span className="game-chip rounded-full px-3 py-1 text-[10px] font-semibold uppercase tracking-[0.28em] text-accent">
                        Geração
                      </span>
                      <span className="game-chip rounded-full px-3 py-1 text-[10px] font-semibold uppercase tracking-[0.28em] text-warning">
                        QR pronto
                      </span>
                    </div>
                  </div>
                </div>

                <div className="rounded-[24px] border border-dashed border-border/70 bg-background/60 p-4">
                  <div className="flex items-center gap-3">
                    <div className="grid size-12 place-items-center rounded-2xl bg-gradient-to-br from-primary/25 to-accent/25 text-primary">
                      <svg
                        width="22"
                        height="22"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2.2"
                      >
                        <path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4z" />
                        <path d="M13 13h3v3h-3zM17 17h3v3h-3zM13 17h3v3h-3zM17 13h3v3h-3z" />
                      </svg>
                    </div>
                    <div>
                      <p className="font-semibold text-foreground">Pronto para o fluxo</p>
                      <p className="text-sm text-muted-foreground">
                        Depois de escolher o perfil, é só carregar a planilha.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="mt-4 flex flex-col items-center justify-center gap-3 text-[11px] tracking-[0.14em] text-slate-400">
            <span className="uppercase">
              Desenvolvido por{" "}
              <span className="font-semibold text-slate-600">{DEVELOPER_NAME}</span>
            </span>
            <div className="flex flex-wrap items-center justify-center gap-2 sm:gap-3">
              <a
                href={DEVELOPER_LINKEDIN_URL}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white/75 px-3 py-1.5 font-semibold uppercase text-slate-500 shadow-sm transition hover:border-[#0a66c2]/30 hover:text-[#0a66c2] hover:shadow-md"
              >
                <Linkedin className="size-3.5" aria-hidden="true" />
                LinkedIn
              </a>
              <a
                href={DEVELOPER_WHATSAPP_URL}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white/75 px-3 py-1.5 font-semibold uppercase text-slate-500 shadow-sm transition hover:border-emerald-400/30 hover:text-emerald-600 hover:shadow-md"
              >
                <MessageCircleMore className="size-3.5" aria-hidden="true" />
                WhatsApp
              </a>
              <a
                href={DEVELOPER_EMAIL_URL}
                className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white/75 px-3 py-1.5 font-semibold uppercase text-slate-500 shadow-sm transition hover:border-rose-400/30 hover:text-rose-600 hover:shadow-md"
              >
                <Mail className="size-3.5" aria-hidden="true" />
                E-mail
              </a>
            </div>
          </div>
        </section>
      </main>
    </div>
  );
}

function SimpleRoleGate({ onSelect }: { onSelect: (role: Role) => void }) {
  return (
    <div className="relative min-h-screen overflow-hidden bg-[radial-gradient(circle_at_top,rgba(17,24,39,0.05),transparent_42%),radial-gradient(circle_at_50%_20%,rgba(37,99,235,0.08),transparent_26%),radial-gradient(circle_at_75%_20%,rgba(34,197,94,0.08),transparent_26%),linear-gradient(180deg,#ffffff_0%,#f6f8fb_100%)] px-6 py-10">
      <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(circle_at_50%_14%,rgba(255,255,255,0.95),transparent_40%)]" />
      <div className="pointer-events-none absolute left-1/2 top-24 h-[26rem] w-[34rem] -translate-x-1/2 rounded-full bg-slate-200/45 blur-3xl" />

      <main className="relative mx-auto flex min-h-screen max-w-6xl items-center justify-center">
        <section className="w-full text-center">
          <p className="text-[11px] uppercase tracking-[0.3em] text-slate-500">
            Selecione seu perfil
          </p>
          <h1 className="mx-auto mt-4 max-w-4xl text-4xl font-black tracking-tight text-slate-900 sm:text-5xl lg:text-6xl">
            Você é <AnimatedRoleWord word="ADM" className="text-[#1f6fb0]" delay={0} /> ou{" "}
            <AnimatedRoleWord word="RH" className="text-[#37a564]" delay={0.24} />?
          </h1>
          <p className="mx-auto mt-3 max-w-2xl text-base text-slate-500 sm:text-lg">
            Escolha sua jornada para começar a aventura no almoxarifado
          </p>

          <div className="mt-12 flex flex-col items-center gap-6 lg:flex-row lg:justify-center lg:gap-5">
            <SimpleRoleChoiceCard
              tone="adm"
              title="ADM"
              eyebrow="Administração"
              description="Gerencia, importa e organiza o estoque"
              onClick={() => onSelect("adm")}
            />
            <SimpleRoleChoiceCard
              tone="rh"
              title="RH"
              eyebrow={"Recursos\nHumanos"}
              description="Acompanha, confere e marca itens lidos"
              onClick={() => onSelect("rh")}
            />
          </div>

          <p className="mt-10 text-[11px] uppercase tracking-[0.22em] text-slate-400">
            Sua escolha fica salva no navegador • pode trocar depois
          </p>
        </section>
      </main>
    </div>
  );
}

function AnimatedRoleWord({
  word,
  className,
  delay = 0,
}: {
  word: string;
  className: string;
  delay?: number;
}) {
  return (
    <span className={`inline-flex items-baseline ${className}`} role="text" aria-label={word}>
      {Array.from(word).map((letter, index) => (
        <span
          key={`${word}-${index}`}
          aria-hidden="true"
          className="role-word-letter"
          style={{ animationDelay: `${delay + index * 0.08}s` }}
        >
          {letter}
        </span>
      ))}
    </span>
  );
}

function SimpleRoleChoiceCard({
  tone,
  title,
  eyebrow,
  description,
  onClick,
}: {
  tone: "adm" | "rh";
  title: string;
  eyebrow: string;
  description: string;
  onClick: () => void;
}) {
  const isAdm = tone === "adm";
  const shellClass = isAdm
    ? "border-[#8eb9dc] bg-[linear-gradient(180deg,rgba(245,249,253,0.98),rgba(255,255,255,0.94))] shadow-[0_28px_70px_rgba(31,86,132,0.16)]"
    : "border-[#a5dbbb] bg-[linear-gradient(180deg,rgba(245,252,247,0.98),rgba(255,255,255,0.94))] shadow-[0_28px_70px_rgba(48,124,84,0.14)]";
  const badgeClass = isAdm
    ? "bg-[#1f6fb0] text-white shadow-[0_12px_22px_rgba(31,111,176,0.24)]"
    : "bg-[#37a564] text-white shadow-[0_12px_22px_rgba(55,165,100,0.24)]";
  const titleClass = isAdm ? "text-[#1f6fb0]" : "text-[#37a564]";
  const buttonClass = isAdm
    ? "bg-[linear-gradient(180deg,#1f6fb0_0%,#2f3f50_100%)] shadow-[0_16px_24px_rgba(24,74,119,0.28)]"
    : "bg-[linear-gradient(180deg,#37a564_0%,#4f8063_100%)] shadow-[0_16px_24px_rgba(37,118,74,0.24)]";
  const starFill = isAdm ? "#2e77b6" : "#42ad6c";
  const star = (
    <svg width="18" height="18" viewBox="0 0 24 24" fill={starFill} aria-hidden="true">
      <path d="M12 2.8l2.71 5.49 6.06.88-4.39 4.28 1.04 6.03L12 16.95 6.58 19.48l1.04-6.03-4.39-4.28 6.06-.88L12 2.8z" />
    </svg>
  );
  const icon = isAdm ? "\u{1F451}" : "\u2B50";

  return (
    <button
      type="button"
      onClick={onClick}
      style={{ animationDelay: isAdm ? "70ms" : "180ms" }}
      className={`group role-card-enter relative flex w-full max-w-[390px] flex-col items-center rounded-[30px] border px-6 pb-8 pt-8 text-center transition duration-300 hover:-translate-y-2 hover:scale-[1.015] hover:shadow-[0_40px_95px_rgba(15,23,42,0.2)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-slate-400/60 ${shellClass}`}
    >
      <div
        className={`absolute -top-5 left-1/2 -translate-x-1/2 rounded-full px-6 py-2 text-sm font-extrabold uppercase tracking-[0.06em] transition duration-300 group-hover:-translate-y-1 group-hover:scale-[1.03] ${badgeClass}`}
      >
        <span className="whitespace-pre-line leading-tight">{eyebrow}</span>
      </div>

      <div className="pointer-events-none absolute inset-0 overflow-hidden rounded-[30px]">
        <div className="absolute inset-y-0 left-[-35%] w-1/2 bg-gradient-to-r from-transparent via-white/65 to-transparent opacity-0 transition-opacity duration-300 group-hover:opacity-100 group-hover:animate-[role-card-sheen_1.15s_ease-out_1]" />
      </div>

      <div className="relative mt-4 flex min-h-[72px] items-center justify-center transition duration-300 group-hover:-translate-y-1">
        <div
          className={`transition-transform duration-300 ease-out group-hover:scale-[1.08] ${
            isAdm ? "group-hover:-rotate-6" : "group-hover:rotate-6"
          }`}
        >
          <span
            role="img"
            aria-label={isAdm ? "coroa" : "estrela"}
            className="role-card-float role-emoji select-none text-[3.55rem] leading-none drop-shadow-[0_10px_18px_rgba(0,0,0,0.12)]"
          >
            {icon}
          </span>
        </div>
      </div>

      <h2
        className={`mt-4 text-[2.75rem] font-black tracking-[-0.06em] transition duration-300 group-hover:-translate-y-0.5 ${titleClass}`}
      >
        {title}
      </h2>
      <p className="mt-3.5 max-w-[20rem] text-[15px] leading-6 text-slate-600 transition duration-300 group-hover:-translate-y-0.5 group-hover:text-slate-700">
        {description}
      </p>

      <div className="mt-5 flex items-center gap-[5px]" aria-hidden="true">
        {star}
        {star}
        {star}
      </div>

      <div
        className={`mt-7 inline-flex items-center gap-2.5 rounded-[16px] px-6 py-3 text-sm font-black uppercase tracking-[0.08em] text-white transition duration-300 group-hover:translate-y-[-1px] group-hover:scale-[1.03] ${buttonClass}`}
      >
        Iniciar
        <span className="text-lg leading-none transition duration-300 group-hover:translate-x-1">
          →
        </span>
      </div>
    </button>
  );
}

function LoadingOverlay({ title, detail }: BusyState) {
  return (
    <div
      aria-live="polite"
      className="fixed inset-0 z-50 flex items-center justify-center bg-background/70 px-6 backdrop-blur-sm"
    >
      <div className="w-full max-w-md rounded-3xl border border-border/60 bg-card/90 p-6 shadow-2xl shadow-black/20">
        <div className="mx-auto flex size-16 items-center justify-center rounded-2xl bg-gradient-to-br from-primary/15 to-accent/15">
          <div className="size-10 animate-spin rounded-full border-4 border-primary/20 border-t-primary" />
        </div>
        <h2 className="mt-4 text-center text-xl font-bold text-foreground">{title}</h2>
        <p className="mt-2 text-center text-sm text-muted-foreground">{detail}</p>
        <div className="mt-5 flex justify-center gap-2">
          <span className="size-2 rounded-full bg-primary animate-bounce" />
          <span className="size-2 rounded-full bg-accent animate-bounce [animation-delay:.15s]" />
          <span className="size-2 rounded-full bg-warning animate-bounce [animation-delay:.3s]" />
        </div>
        <div className="mt-6 h-2 overflow-hidden rounded-full bg-muted">
          <div className="h-full w-1/2 animate-pulse rounded-full bg-gradient-to-r from-primary via-accent to-warning" />
        </div>
      </div>
    </div>
  );
}

function RoleChoiceCard({
  title,
  eyebrow,
  description,
  tone,
  onClick,
}: {
  title: string;
  eyebrow: string;
  description: string;
  tone: "adm" | "rh";
  onClick: () => void;
}) {
  const accent =
    tone === "adm" ? "from-primary via-accent to-warning" : "from-accent via-primary to-success";
  const icon =
    tone === "adm" ? (
      <svg
        width="30"
        height="30"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
      >
        <path d="M4 5h16v14H4z" />
        <path d="M7 8h10" />
        <path d="M7 12h6" />
        <path d="M7 16h4" />
      </svg>
    ) : (
      <svg
        width="30"
        height="30"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
      >
        <path d="M12 3l8 4v10l-8 4-8-4V7l8-4z" />
        <path d="M12 7v10" />
        <path d="M7 9.5c1.4.9 2.9 1.3 5 1.3s3.6-.4 5-1.3" />
      </svg>
    );

  return (
    <button
      type="button"
      onClick={onClick}
      className="game-panel group relative flex w-full flex-col gap-4 overflow-hidden rounded-[28px] p-5 text-left transition duration-300 hover:-translate-y-1 hover:shadow-[0_30px_70px_rgba(8,32,50,0.18)]"
    >
      <div className={`absolute inset-x-0 top-0 h-1 bg-gradient-to-r ${accent}`} />
      <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(circle_at_top_right,rgba(255,255,255,0.72),transparent_38%)] opacity-0 transition duration-300 group-hover:opacity-100" />
      <div className="pointer-events-none absolute inset-x-0 bottom-0 h-16 bg-gradient-to-t from-foreground/5 to-transparent opacity-0 transition duration-300 group-hover:opacity-100" />
      <div className="pointer-events-none absolute right-4 top-4 text-foreground/10 transition duration-300 group-hover:scale-110 group-hover:text-foreground/15">
        {icon}
      </div>
      <div className="relative pr-14">
        <p className="text-[10px] font-semibold uppercase tracking-[0.34em] text-muted-foreground">
          {eyebrow}
        </p>
        <h2 className="font-display mt-2 text-3xl font-black text-foreground">{title}</h2>
        <p className="mt-3 max-w-xs text-sm leading-relaxed text-muted-foreground">{description}</p>
      </div>
      <div className="inline-flex w-fit items-center gap-2 rounded-full border border-border/60 bg-card/80 px-3 py-1 text-xs font-semibold text-foreground shadow-sm">
        <span className="size-2 rounded-full bg-primary" />
        Entrar
      </div>
    </button>
  );
}

function Btn({
  children,
  onClick,
  disabled,
  variant = "primary",
  className = "",
}: {
  children: React.ReactNode;
  onClick?: React.ButtonHTMLAttributes<HTMLButtonElement>["onClick"];
  disabled?: boolean;
  variant?: "primary" | "ghost";
  className?: string;
}) {
  const cls =
    variant === "primary"
      ? "bg-gradient-to-r from-primary to-accent text-primary-foreground shadow-lg shadow-primary/30 hover:opacity-90"
      : "border border-border/60 bg-secondary/70 text-foreground hover:bg-secondary";
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className={`rounded-xl px-4 py-2 text-sm font-medium transition disabled:cursor-not-allowed disabled:opacity-40 ${cls} ${className}`}
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
      <p className="text-[10px] font-semibold uppercase tracking-widest text-muted-foreground">
        {label}
      </p>
      <p className="mt-1 text-2xl font-bold tabular-nums text-foreground">{value}</p>
    </div>
  );
}
