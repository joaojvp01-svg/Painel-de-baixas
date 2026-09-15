import React, { useState, useEffect, useCallback, useMemo } from "react";
import { useParams } from "react-router-dom";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, LabelList
} from "recharts";
import { TruckIcon, PackageX, Clock, Loader2, Info, ShieldCheck, Download } from "lucide-react";
import { storage } from "../lib/storage";
import { findCarrierBySlug } from "../lib/slug";

const STORAGE_KEY = "baixas-summary";
const LOGO_DATA_URI = "data:image/webp;base64,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";

const COLORS = {
  bg: "#f4f5f7", panel: "#ffffff", panelAlt: "#fafafa",
  border: "#e6e8ec", borderLight: "#d8dbe2",
  text: "#1f2430", textDim: "#667085", textFaint: "#98a2b3",
  orange: "#FF6901", orangeSoft: "#fff2e8",
  navy: "#2b3a55", teal: "#0d9488", amber: "#d97706", red: "#dc2626",
};

function fmtInt(n) { return (n ?? 0).toLocaleString("pt-BR"); }
function fmtPct(n) { return `${(n ?? 0).toLocaleString("pt-BR", { maximumFractionDigits: 1 })}%`; }
function fmtDate(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "—";
  return d.toLocaleDateString("pt-BR");
}
function urgencyStyle(dias) {
  if (dias >= 1) return { bg: "#fee2e2", border: COLORS.red, color: COLORS.red, label: "Crítico" };
  return { bg: "#e6f9f6", border: COLORS.teal, color: COLORS.teal, label: "No prazo" };
}
function mostCommonName(names) {
  const counts = {};
  let best = "Não informado", bestCount = 0;
  for (const n of names) {
    const key = n || "Não informado";
    counts[key] = (counts[key] || 0) + 1;
    if (counts[key] > bestCount) { bestCount = counts[key]; best = key; }
  }
  return best;
}

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <div style={{ background: "#fff", border: `1px solid ${COLORS.borderLight}`, borderRadius: 8, padding: "8px 12px", fontSize: 12, boxShadow: "0 4px 12px rgba(16,24,40,0.08)" }}>
      <div style={{ color: COLORS.textDim, marginBottom: 4, fontWeight: 600 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || COLORS.text, fontFamily: "ui-monospace, monospace" }}>{p.name}: {fmtInt(p.value)}</div>
      ))}
    </div>
  );
};

function KpiCard({ icon: Icon, label, value, sub, accent }) {
  return (
    <div style={{
      background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10,
      padding: "16px 18px", display: "flex", flexDirection: "column", gap: 8, flex: 1, minWidth: 0,
      position: "relative", overflow: "hidden", boxShadow: "0 1px 2px rgba(16,24,40,0.04)",
    }}>
      <div style={{ position: "absolute", top: 0, left: 0, width: 3, height: "100%", background: accent }} />
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: COLORS.textDim, fontWeight: 700 }}>{label}</span>
        <Icon size={16} color={accent} strokeWidth={2.2} />
      </div>
      <div style={{ fontSize: 26, fontWeight: 700, fontFamily: "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace", color: COLORS.text, letterSpacing: "-0.02em" }}>{value}</div>
      {sub && <div style={{ fontSize: 12, color: COLORS.textFaint }}>{sub}</div>}
    </div>
  );
}

function Panel({ title, subtitle, children, style, right }) {
  return (
    <div style={{
      background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10,
      padding: 18, display: "flex", flexDirection: "column", gap: 12, minWidth: 0,
      boxShadow: "0 1px 2px rgba(16,24,40,0.04)", ...style,
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, flexWrap: "wrap" }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.text }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11.5, color: COLORS.textFaint, marginTop: 2 }}>{subtitle}</div>}
        </div>
        {right}
      </div>
      {children}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Visão por placa — placa, motorista e notas pendentes agrupadas     */
/* ------------------------------------------------------------------ */
function PlacaTables({ notes, carrierName }) {
  const [openPlaca, setOpenPlaca] = useState(null);

  const groups = useMemo(() => {
    const byPlaca = {};
    for (const n of notes) {
      if (!byPlaca[n.placa]) byPlaca[n.placa] = [];
      byPlaca[n.placa].push(n);
    }
    return Object.entries(byPlaca)
      .map(([placa, list]) => ({
        placa,
        motorista: mostCommonName(list.map((n) => n.motorista)),
        notes: list.sort((a, b) => b.diasAberto - a.diasAberto),
        total: list.length,
      }))
      .sort((a, b) => b.total - a.total);
  }, [notes]);

  useEffect(() => { setOpenPlaca(groups[0]?.placa ?? null); }, [notes.length]); // eslint-disable-line

  const handleExport = () => {
    if (!groups.length) return;
    const wb = XLSX.utils.book_new();
    const usedNames = new Set();

    groups.forEach((g) => {
      const rows = g.notes.map((n) => ({
        "Nota Fiscal": n.nf,
        "Placa": g.placa,
        "Motorista": g.motorista,
        "Data da Rota": fmtDate(n.dataRota),
        "Mês da Rota": n.mesRota,
        "Dias em Aberto": n.diasAberto,
        "Cliente": n.cliente,
      }));
      const ws = XLSX.utils.json_to_sheet(rows);
      ws["!cols"] = [{ wch: 12 }, { wch: 12 }, { wch: 24 }, { wch: 13 }, { wch: 12 }, { wch: 14 }, { wch: 38 }];

      // Nomes de aba no Excel: máx. 31 caracteres, sem \ / ? * [ ] : e não podem repetir.
      let base = (g.placa || "SEM PLACA").replace(/[\\/?*[\]:]/g, "-").trim().slice(0, 31) || "SEM PLACA";
      let name = base;
      let suffix = 1;
      while (usedNames.has(name)) {
        suffix += 1;
        const tail = ` (${suffix})`;
        name = base.slice(0, 31 - tail.length) + tail;
      }
      usedNames.add(name);
      XLSX.utils.book_append_sheet(wb, ws, name);
    });

    const dateStr = new Date().toISOString().slice(0, 10);
    const carrierSlug = (carrierName || "transportadora").replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 50);
    XLSX.writeFile(wb, `Pendencias_${carrierSlug}_${dateStr}.xlsx`);
  };

  if (groups.length === 0) {
    return (
      <Panel title="Notas pendentes por placa">
        <div style={{ padding: "16px 10px", color: COLORS.textFaint, textAlign: "center", fontSize: 12.5 }}>
          Nenhuma nota pendente — tudo regularizado.
        </div>
      </Panel>
    );
  }

  return (
    <Panel
      title="Notas pendentes por placa"
      subtitle="Clique em uma placa para ver as notas — priorize as marcadas como &quot;Crítico&quot;"
      right={
        <button
          onClick={handleExport}
          style={{
            display: "flex", alignItems: "center", gap: 7, background: COLORS.orange, color: "#fff",
            border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer",
          }}
        >
          <Download size={14} /> Exportar Excel
        </button>
      }
    >
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {groups.map((g) => {
          const isOpen = openPlaca === g.placa;
          return (
            <div key={g.placa} style={{ border: `1px solid ${COLORS.border}`, borderRadius: 8, overflow: "hidden" }}>
              <button
                onClick={() => setOpenPlaca(isOpen ? null : g.placa)}
                style={{
                  width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between",
                  padding: "10px 14px", background: isOpen ? COLORS.orangeSoft : COLORS.panelAlt, border: "none",
                  cursor: "pointer", fontSize: 13,
                }}
              >
                <span style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: 700, color: COLORS.text, fontFamily: "ui-monospace, monospace" }}>
                  {isOpen ? "▾" : "▸"} {g.placa}
                  <span style={{ fontWeight: 500, color: COLORS.textDim, fontFamily: "-apple-system, sans-serif", fontSize: 12 }}>· {g.motorista}</span>
                </span>
                <span style={{ fontSize: 12, fontWeight: 700, color: COLORS.orange, background: "#fff", padding: "3px 10px", borderRadius: 20, border: `1px solid ${COLORS.orange}` }}>
                  {fmtInt(g.total)} nota{g.total > 1 ? "s" : ""} pendente{g.total > 1 ? "s" : ""}
                </span>
              </button>
              {isOpen && (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                    <thead>
                      <tr>
                        {["NF", "Data da rota", "Mês da rota", "Dias em aberto", "Cliente", "Prioridade"].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "7px 12px", color: COLORS.textDim, fontWeight: 600, borderBottom: `1px solid ${COLORS.border}`, background: COLORS.panel }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {g.notes.map((n, i) => {
                        const u = urgencyStyle(n.diasAberto);
                        return (
                          <tr key={i} style={{ background: n.diasAberto >= 1 ? "#fef7f7" : "transparent" }}>
                            <td style={{ padding: "7px 12px", borderBottom: `1px solid ${COLORS.border}`, fontFamily: "ui-monospace, monospace", color: COLORS.textDim }}>{n.nf}</td>
                            <td style={{ padding: "7px 12px", borderBottom: `1px solid ${COLORS.border}`, color: COLORS.text }}>{fmtDate(n.dataRota)}</td>
                            <td style={{ padding: "7px 12px", borderBottom: `1px solid ${COLORS.border}`, color: COLORS.textDim }}>{n.mesRota}</td>
                            <td style={{ padding: "7px 12px", borderBottom: `1px solid ${COLORS.border}`, fontFamily: "ui-monospace, monospace", fontWeight: 700, color: u.color }}>{n.diasAberto}d</td>
                            <td style={{ padding: "7px 12px", borderBottom: `1px solid ${COLORS.border}`, color: COLORS.textDim, maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{n.cliente}</td>
                            <td style={{ padding: "7px 12px", borderBottom: `1px solid ${COLORS.border}` }}>
                              <span style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "3px 10px", borderRadius: 20, fontSize: 11, fontWeight: 700, background: u.bg, color: u.color, border: `1px solid ${u.border}` }}>{u.label}</span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </Panel>
  );
}

export default function PartnerDashboard() {
  const { slug } = useParams();
  const [summary, setSummary] = useState(null);
  const [loading, setLoading] = useState(true);

  const loadStored = useCallback(async () => {
    setLoading(true);
    try {
      const result = await storage.get(STORAGE_KEY);
      if (result && result.value) setSummary(JSON.parse(result.value));
    } catch (e) {
      // ainda sem dados
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { loadStored(); }, [loadStored]);

  const carrierName = useMemo(() => {
    if (!summary) return null;
    return findCarrierBySlug(summary.carriers, slug);
  }, [summary, slug]);

  const carrierData = useMemo(() => {
    if (!summary || !carrierName) return null;
    const stats = summary.carrierStats[carrierName];
    if (!stats) return null;
    const slaBase = stats.noPrazo + stats.foraPrazo + stats.atencao;
    const slaPct = slaBase ? (stats.noPrazo / slaBase) * 100 : 0;
    const pctPendente = stats.total ? (stats.pendente / stats.total) * 100 : 0;
    const notes = (summary.pendingAll || [])
      .filter((p) => p.transportadora === carrierName)
      .sort((a, b) => b.diasAberto - a.diasAberto);
    return { stats, slaPct, pctPendente, notes };
  }, [summary, carrierName]);

  const monthChartData = useMemo(() => {
    if (!summary || !carrierName) return [];
    const data = (summary.carrierMonthPending || {})[carrierName] || {};
    return (summary.monthsAsc || []).map((m) => ({ mes: m, pendentes: data[m] || 0 }));
  }, [summary, carrierName]);

  const lastUpdated = summary?.generatedAt ? new Date(summary.generatedAt).toLocaleDateString("pt-BR") : null;

  return (
    <div style={{
      minHeight: "100vh", background: COLORS.bg, color: COLORS.text,
      fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
      padding: "20px 20px 40px",
    }}>
      <div style={{ maxWidth: 1000, margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 14, marginBottom: 22, paddingBottom: 18, borderBottom: `1px solid ${COLORS.border}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <img src={LOGO_DATA_URI} alt="Logo" style={{ height: 30, width: "auto", objectFit: "contain", flexShrink: 0 }} />
            <div style={{ width: 1, height: 28, background: COLORS.border, flexShrink: 0 }} />
            <div>
              <div style={{ fontSize: 17, fontWeight: 700, letterSpacing: "-0.01em" }}>Painel de Pendências</div>
              <div style={{ fontSize: 12, color: COLORS.textFaint }}>Portal do parceiro transportador</div>
            </div>
          </div>
          {lastUpdated && (
            <div style={{ fontSize: 11.5, color: COLORS.textFaint, display: "flex", alignItems: "center", gap: 6 }}>
              <ShieldCheck size={13} color={COLORS.teal} /> Dados atualizados em {lastUpdated}
            </div>
          )}
        </div>

        {loading && (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 80, color: COLORS.textFaint, gap: 10 }}>
            <Loader2 size={18} className="spin" /> Carregando painel...
          </div>
        )}

        {!loading && !summary && (
          <div style={{ border: `1.5px dashed ${COLORS.borderLight}`, borderRadius: 12, padding: "60px 24px", display: "flex", flexDirection: "column", alignItems: "center", gap: 12, color: COLORS.textDim }}>
            <Info size={26} color={COLORS.orange} />
            <div style={{ fontSize: 15, fontWeight: 600, color: COLORS.text }}>Nenhum dado disponível no momento</div>
            <div style={{ fontSize: 13, textAlign: "center", maxWidth: 420, color: COLORS.textFaint }}>
              Assim que a equipe carregar a planilha mais recente no painel gerencial, os dados aparecem aqui automaticamente.
            </div>
          </div>
        )}

        {!loading && summary && !carrierName && (
          <div style={{ border: `1.5px dashed ${COLORS.borderLight}`, borderRadius: 12, padding: "60px 24px", display: "flex", flexDirection: "column", alignItems: "center", gap: 12, color: COLORS.textDim }}>
            <Info size={26} color={COLORS.orange} />
            <div style={{ fontSize: 15, fontWeight: 600, color: COLORS.text }}>Transportadora não encontrada</div>
            <div style={{ fontSize: 13, textAlign: "center", maxWidth: 400, color: COLORS.textFaint }}>
              O link não corresponde a nenhuma transportadora nos dados carregados. Confira com quem enviou o link.
            </div>
          </div>
        )}

        {!loading && carrierData && (
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            <div style={{ fontSize: 15, fontWeight: 700, color: COLORS.text, display: "flex", alignItems: "center", gap: 8 }}>
              <TruckIcon size={18} color={COLORS.orange} /> {carrierName}
            </div>

            <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
              <KpiCard icon={TruckIcon} label="Total de notas carregadas" accent={COLORS.navy} value={fmtInt(carrierData.stats.total)} sub="No período carregado" />
              <KpiCard icon={PackageX} label="Pendentes de regularização" accent={COLORS.orange} value={fmtInt(carrierData.stats.pendente)} sub={fmtPct(carrierData.pctPendente)} />
              <KpiCard icon={Clock} label="SLA no prazo" accent={COLORS.teal} value={fmtPct(carrierData.slaPct)} sub={`${fmtInt(carrierData.stats.foraPrazo)} fora do prazo`} />
            </div>

            <Panel title="Pendências por mês" subtitle="Situação de todos os meses carregados">
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={monthChartData} margin={{ top: 4, right: 8, left: -14, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
                  <XAxis dataKey="mes" tick={{ fill: COLORS.textFaint, fontSize: 11 }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                  <YAxis tick={{ fill: COLORS.textFaint, fontSize: 11 }} axisLine={false} tickLine={false} />
                  <Tooltip content={<CustomTooltip />} cursor={{ fill: "rgba(255,105,1,0.05)" }} />
                  <Bar dataKey="pendentes" name="Pendentes" fill={COLORS.orange} radius={[4, 4, 0, 0]}>
                    <LabelList dataKey="pendentes" position="top" style={{ fill: COLORS.text, fontSize: 11.5, fontWeight: 700 }} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </Panel>

            <PlacaTables notes={carrierData.notes} carrierName={carrierName} />

            <div style={{ fontSize: 11, color: COLORS.textFaint, textAlign: "center" }}>
              Este painel mostra apenas os dados de {carrierName}. Em caso de dúvidas sobre alguma nota, entre em contato com o time de indenizações.
            </div>
          </div>
        )}

        <style>{`
          @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
          .spin { animation: spin 1s linear infinite; }
        `}</style>
      </div>
    </div>
  );
}
