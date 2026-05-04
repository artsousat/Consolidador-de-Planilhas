import { createFileRoute } from "@tanstack/react-router";
import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import * as XLSXStyle from "xlsx-js-style";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Upload, FileSpreadsheet, Download, CheckCircle2, AlertCircle, X } from "lucide-react";

export const Route = createFileRoute("/")({
  component: Index,
});

type Row = { nome: string; numero: string; parcela: string };
type FinalRow = Row & { origem: "CelCash" | "Maistodos" };
type StatusRow = Row & {
  cobrador: string;
  status: string;
  data: string;
  obs: string;
};
type OutRow = FinalRow & {
  cobrador: string;
  status: string;
  data: string;
  obs: string;
};

const normNome = (v: unknown) => String(v ?? "").trim().toLowerCase();
const normNum = (v: unknown) => String(v ?? "").replace(/\D/g, "").slice(-11);
const normStatus = (v: unknown) => String(v ?? "").trim().toUpperCase();

async function readSheet(file: File): Promise<unknown[][]> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, blankrows: false, defval: "" });
}

function rowsFrom(data: unknown[][]): Row[] {
  return data
    .slice(1)
    .map((r) => ({
      nome: String(r?.[0] ?? "").trim(),
      numero: String(r?.[1] ?? "").trim(),
      parcela: String(r?.[2] ?? "").trim(),
    }))
    .filter((r) => r.nome !== "" || r.numero !== "");
}

function statusRowsFrom(data: unknown[][]): StatusRow[] {
  return data
    .slice(1)
    .map((r) => ({
      nome: String(r?.[0] ?? "").trim(),
      numero: String(r?.[1] ?? "").trim(),
      parcela: String(r?.[2] ?? "").trim(),
      cobrador: String(r?.[3] ?? "").trim(),
      status: String(r?.[4] ?? "").trim(),
      data: r?.[5] instanceof Date
        ? (r[5] as Date).toLocaleDateString("pt-BR")
        : String(r?.[5] ?? "").trim(),
      obs: String(r?.[6] ?? "").trim(),
    }))
    .filter((r) => r.nome !== "" || r.numero !== "");
}

type Stats = {
  celTotal: number;
  celClean: number;
  mtTotal: number;
  mtClean: number;
  duplicados: number;
  finalTotal: number;
  statusTotal: number;
  statusAplicados: number;
  naoEncontrados: number;
  porStatus: Record<string, number>;
};

function processar(
  cel: Row[],
  mt: Row[],
  bug: Row[],
  status: StatusRow[],
): { final: OutRow[]; naoEncontrados: StatusRow[]; stats: Stats } {
  const bugNomes = new Set(bug.map((r) => normNome(r.nome)).filter(Boolean));
  const bugNums = new Set(bug.map((r) => normNum(r.numero)).filter(Boolean));

  const limpa = (rows: Row[]) =>
    rows.filter((r) => {
      const n = normNome(r.nome);
      const num = normNum(r.numero);
      if (n && bugNomes.has(n)) return false;
      if (num && bugNums.has(num)) return false;
      return true;
    });

  const celClean = limpa(cel);
  const mtClean = limpa(mt);

  // MaisTodos prioridade: remove do CelCash quem já está em MaisTodos (por número OU nome)
  const mtNums = new Set(mtClean.map((r) => normNum(r.numero)).filter(Boolean));
  const mtNomes = new Set(mtClean.map((r) => normNome(r.nome)).filter(Boolean));

  let duplicados = 0;
  const celFinal = celClean.filter((r) => {
    const num = normNum(r.numero);
    const n = normNome(r.nome);
    const dup = (num && mtNums.has(num)) || (n && mtNomes.has(n));
    if (dup) duplicados++;
    return !dup;
  });

  const consolidado: FinalRow[] = [
    ...mtClean.map((r) => ({ ...r, origem: "Maistodos" as const })),
    ...celFinal.map((r) => ({ ...r, origem: "CelCash" as const })),
  ];

  // Index status by número e nome
  const statusByNum = new Map<string, StatusRow>();
  const statusByNome = new Map<string, StatusRow>();
  for (const s of status) {
    const num = normNum(s.numero);
    const n = normNome(s.nome);
    if (num) statusByNum.set(num, s);
    if (n) statusByNome.set(n, s);
  }

  const usados = new Set<StatusRow>();
  let aplicados = 0;
  const final: OutRow[] = consolidado.map((c) => {
    const num = normNum(c.numero);
    const n = normNome(c.nome);
    const s = (num && statusByNum.get(num)) || (n && statusByNome.get(n)) || null;
    if (s) {
      usados.add(s);
      aplicados++;
      return {
        ...c,
        cobrador: s.cobrador,
        status: s.status,
        data: s.data,
        obs: s.obs,
      };
    }
    return { ...c, cobrador: "", status: "", data: "", obs: "" };
  });

  const naoEncontrados = status.filter((s) => !usados.has(s));

  const porStatus: Record<string, number> = {};
  for (const r of final) {
    const k = normStatus(r.status);
    if (!k) continue;
    porStatus[k] = (porStatus[k] ?? 0) + 1;
  }

  return {
    final,
    naoEncontrados,
    stats: {
      celTotal: cel.length,
      celClean: celClean.length,
      mtTotal: mt.length,
      mtClean: mtClean.length,
      duplicados,
      finalTotal: final.length,
      statusTotal: status.length,
      statusAplicados: aplicados,
      naoEncontrados: naoEncontrados.length,
      porStatus,
    },
  };
}

function FileSlot({
  label,
  file,
  onSelect,
  onClear,
  accent,
}: {
  label: string;
  file: File | null;
  onSelect: (f: File) => void;
  onClear: () => void;
  accent: string;
}) {
  const ref = useRef<HTMLInputElement>(null);
  return (
    <div
      className="group relative flex items-center gap-4 rounded-2xl border-2 border-dashed border-border bg-card p-5 transition-all hover:border-primary/50"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => {
        e.preventDefault();
        const f = e.dataTransfer.files?.[0];
        if (f) onSelect(f);
      }}
    >
      <div
        className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl text-white"
        style={{ background: accent }}
      >
        <FileSpreadsheet className="h-6 w-6" />
      </div>
      <div className="min-w-0 flex-1">
        <div className="text-sm font-semibold text-foreground">{label}</div>
        <div className="truncate text-xs text-muted-foreground">
          {file ? file.name : "Arraste o .xlsx aqui ou clique em Selecionar"}
        </div>
      </div>
      {file ? (
        <button
          onClick={onClear}
          className="rounded-lg p-2 text-muted-foreground hover:bg-muted hover:text-foreground"
          aria-label="Remover arquivo"
        >
          <X className="h-4 w-4" />
        </button>
      ) : (
        <Button variant="outline" size="sm" onClick={() => ref.current?.click()}>
          <Upload className="mr-2 h-4 w-4" />
          Selecionar
        </Button>
      )}
      <input
        ref={ref}
        type="file"
        accept=".xlsx"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onSelect(f);
          e.target.value = "";
        }}
      />
    </div>
  );
}

function App() {
  const [cel, setCel] = useState<File | null>(null);
  const [mt, setMt] = useState<File | null>(null);
  const [bug, setBug] = useState<File | null>(null);
  const [stsFile, setStsFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<{
    final: OutRow[];
    naoEncontrados: StatusRow[];
    stats: Stats;
  } | null>(null);

  const podeProcessar = cel && mt && bug && !loading;

  async function handleProcessar() {
    if (!cel || !mt || !bug) return;
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const [celData, mtData, bugData, stsData] = await Promise.all([
        readSheet(cel),
        readSheet(mt),
        readSheet(bug),
        stsFile ? readSheet(stsFile) : Promise.resolve([] as unknown[][]),
      ]);
      const celRows = rowsFrom(celData);
      const mtRows = rowsFrom(mtData);
      const bugRows = rowsFrom(bugData);
      const stsRows = statusRowsFrom(stsData);
      const r = processar(celRows, mtRows, bugRows, stsRows);
      setResult(r);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Erro ao processar planilhas");
    } finally {
      setLoading(false);
    }
  }

  function handleDownload() {
    if (!result) return;
    const header = [
      "Nome",
      "Número",
      "Parcela",
      "Cobrador",
      "Status",
      "Data",
      "Observação",
      "Origem",
    ];
    const aoa: unknown[][] = [
      header,
      ...result.final.map((r) => [
        r.nome,
        r.numero,
        r.parcela,
        r.cobrador,
        r.status,
        r.data,
        r.obs,
        r.origem,
      ]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [
      { wch: 32 },
      { wch: 16 },
      { wch: 8 },
      { wch: 12 },
      { wch: 12 },
      { wch: 10 },
      { wch: 28 },
      { wch: 12 },
    ];

    // Cores por status (preenchimento da linha A:H)
    const cores: Record<string, string> = {
      AP: "FFFF00", // amarelo
      CANCELADO: "FFA500", // laranja
      CHATA: "FF99CC", // rosa
      CHATINHA: "FF99CC",
    };
    for (let i = 0; i < result.final.length; i++) {
      const r = result.final[i];
      const key = normStatus(r.status);
      const cor = cores[key];
      if (!cor) continue;
      const rowIdx = i + 1; // +1 por causa do header
      for (let c = 0; c < header.length; c++) {
        const addr = XLSX.utils.encode_cell({ r: rowIdx, c });
        const cell = ws[addr] ?? { t: "s", v: "" };
        cell.s = {
          fill: { patternType: "solid", fgColor: { rgb: cor } },
          font: { color: { rgb: "000000" } },
        };
        ws[addr] = cell;
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Consolidado");

    if (result.naoEncontrados.length > 0) {
      const ne = [
        ["Nome", "Número", "Parcela", "Cobrador", "Status", "Data", "Observação"],
        ...result.naoEncontrados.map((r) => [
          r.nome,
          r.numero,
          r.parcela,
          r.cobrador,
          r.status,
          r.data,
          r.obs,
        ]),
      ];
      const wsNE = XLSX.utils.aoa_to_sheet(ne);
      wsNE["!cols"] = [
        { wch: 32 },
        { wch: 16 },
        { wch: 8 },
        { wch: 12 },
        { wch: 12 },
        { wch: 10 },
        { wch: 28 },
      ];
      XLSX.utils.book_append_sheet(wb, wsNE, "Não Encontrados");
    }

    XLSXStyle.writeFile(wb, "resultado.xlsx");
  }

  function reset() {
    setCel(null);
    setMt(null);
    setBug(null);
    setStsFile(null);
    setResult(null);
    setError(null);
  }

  return (
    <div className="min-h-screen bg-background">
      <div className="mx-auto max-w-3xl px-4 py-10 md:py-16">
        <header className="mb-10">
          <div className="mb-3 inline-flex items-center gap-2 rounded-full border border-border bg-card px-3 py-1 text-xs font-medium text-muted-foreground">
            <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
            100% no navegador · seus dados não saem da máquina
          </div>
          <h1 className="text-3xl font-bold tracking-tight text-foreground md:text-4xl">
            Consolidador de Planilhas
          </h1>
          <p className="mt-2 text-muted-foreground">
            Junta CelCash + MaisTodos, remove clientes Bugados e elimina duplicados
            (MaisTodos tem prioridade).
          </p>
        </header>

        <Card className="p-6 md:p-8">
          <div className="space-y-3">
            <FileSlot
              label="CelCash.xlsx"
              file={cel}
              onSelect={setCel}
              onClear={() => setCel(null)}
              accent="linear-gradient(135deg, #0ea5e9, #1e40af)"
            />
            <FileSlot
              label="MaisTodos.xlsx"
              file={mt}
              onSelect={setMt}
              onClear={() => setMt(null)}
              accent="linear-gradient(135deg, #10b981, #047857)"
            />
            <FileSlot
              label="Bugados.xlsx"
              file={bug}
              onSelect={setBug}
              onClear={() => setBug(null)}
              accent="linear-gradient(135deg, #f43f5e, #9f1239)"
            />
            <FileSlot
              label="Status.xlsx (opcional — AP, Cancelado, Chata...)"
              file={stsFile}
              onSelect={setStsFile}
              onClear={() => setStsFile(null)}
              accent="linear-gradient(135deg, #eab308, #b45309)"
            />
          </div>

          <div className="mt-6 flex items-center gap-3">
            <Button
              size="lg"
              className="flex-1"
              disabled={!podeProcessar}
              onClick={handleProcessar}
            >
              {loading ? "Processando..." : "Processar planilhas"}
            </Button>
            {(cel || mt || bug || result) && (
              <Button variant="ghost" size="lg" onClick={reset}>
                Limpar
              </Button>
            )}
          </div>

          {error && (
            <div className="mt-4 flex items-start gap-3 rounded-lg border border-destructive/30 bg-destructive/5 p-3 text-sm text-destructive">
              <AlertCircle className="mt-0.5 h-4 w-4 shrink-0" />
              <span>{error}</span>
            </div>
          )}

          {result && (
            <div className="mt-6 space-y-4 rounded-xl border border-border bg-muted/30 p-5">
              <div className="flex items-center gap-2 text-sm font-semibold text-foreground">
                <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                Pronto! Resumo do processamento
              </div>
              <dl className="grid grid-cols-2 gap-3 text-sm">
                <Stat label="CelCash" value={`${result.stats.celTotal} → ${result.stats.celClean}`} hint="após remover Bugados" />
                <Stat label="MaisTodos" value={`${result.stats.mtTotal} → ${result.stats.mtClean}`} hint="após remover Bugados" />
                <Stat label="Duplicados removidos" value={String(result.stats.duplicados)} hint="CelCash já em MaisTodos" />
                <Stat label="Total final" value={String(result.stats.finalTotal)} hint="clientes na planilha" highlight />
                {result.stats.statusTotal > 0 && (
                  <>
                    <Stat
                      label="Status aplicados"
                      value={`${result.stats.statusAplicados}/${result.stats.statusTotal}`}
                      hint="clientes com status definido"
                    />
                    <Stat
                      label="Não encontrados"
                      value={String(result.stats.naoEncontrados)}
                      hint="vão pra aba separada"
                    />
                  </>
                )}
              </dl>
              {Object.keys(result.stats.porStatus).length > 0 && (
                <div className="flex flex-wrap gap-2 pt-1">
                  {Object.entries(result.stats.porStatus).map(([k, v]) => (
                    <span
                      key={k}
                      className="rounded-full border border-border bg-background px-2.5 py-1 text-xs font-medium"
                    >
                      {k}: <span className="font-bold">{v}</span>
                    </span>
                  ))}
                </div>
              )}
              <Button size="lg" className="w-full" onClick={handleDownload}>
                <Download className="mr-2 h-4 w-4" />
                Baixar resultado.xlsx
              </Button>
            </div>
          )}
        </Card>

        <section className="mt-10">
          <h2 className="mb-3 text-lg font-semibold text-foreground">
            Como preparar suas planilhas
          </h2>
          <p className="mb-4 text-sm text-muted-foreground">
            Cada arquivo .xlsx deve ter cabeçalho na linha 1 e dados a partir da linha 2.
            O programa lê sempre a primeira aba do arquivo.
          </p>
          <div className="grid gap-3 md:grid-cols-2">
            <SheetSpec title="CelCash" cols={["A — Nome", "B — Número", "C — Parcela"]} />
            <SheetSpec title="MaisTodos" cols={["A — Nome", "B — Número", "C — Parcela"]} />
            <SheetSpec title="Bugados" cols={["A — Nome", "B — Número", "(C ignorada)"]} />
            <SheetSpec
              title="Status (opcional)"
              cols={[
                "A — Nome",
                "B — Número",
                "C — Parcela",
                "D — Cobrador",
                "E — Status (AP, Cancelado, Chata...)",
                "F — Data",
                "G — Observação",
              ]}
            />
          </div>
          <ul className="mt-5 space-y-1.5 text-xs text-muted-foreground">
            <li>• Comparação de número usa só os últimos 11 dígitos (ignora pontuação).</li>
            <li>• Comparação de nome ignora maiúsculas/minúsculas e espaços extras.</li>
            <li>• Linhas vazias são descartadas automaticamente.</li>
            <li>• Status: AP = amarelo, Cancelado = laranja, Chata/Chatinha = rosa.</li>
            <li>• Clientes da Status que não estão no Consolidado vão para a aba "Não Encontrados".</li>
          </ul>
        </section>
      </div>
    </div>
  );
}

function Stat({
  label,
  value,
  hint,
  highlight,
}: {
  label: string;
  value: string;
  hint: string;
  highlight?: boolean;
}) {
  return (
    <div
      className={`rounded-lg border p-3 ${
        highlight ? "border-primary/30 bg-primary/5" : "border-border bg-background"
      }`}
    >
      <dt className="text-xs uppercase tracking-wide text-muted-foreground">{label}</dt>
      <dd className="mt-1 text-xl font-bold text-foreground">{value}</dd>
      <dd className="text-[11px] text-muted-foreground">{hint}</dd>
    </div>
  );
}

function SheetSpec({ title, cols }: { title: string; cols: string[] }) {
  return (
    <div className="rounded-xl border border-border bg-card p-4">
      <div className="mb-2 text-sm font-semibold text-foreground">{title}.xlsx</div>
      <ul className="space-y-1 text-xs text-muted-foreground">
        {cols.map((c) => (
          <li key={c} className="font-mono">{c}</li>
        ))}
      </ul>
    </div>
  );
}

function Index() {
  return <App />;
}
