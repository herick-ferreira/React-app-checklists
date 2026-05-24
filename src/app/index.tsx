// @ts-nocheck
import { useCallback, useEffect, useMemo, useRef, useState } from "react";

const SIDEBAR_BG = "#001D3D";
const CARD_BG = "#F7F8FA";
const APP_BG = "#EDEDED";

const LIST_MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];

// ── Dados de exemplo gerados programaticamente ──────────────────────────
function gerarDados() {
  const lojas = ["Loja A", "Loja B", "Loja C", "Loja D", "Loja E"];
  const topicos = ["Atendimento", "Limpeza", "Produto", "Estoque", "Caixa"];
  const tags = ["Urgente", "Rotina", "Melhoria", "Crítico", "Elogio"];
  const questoes = [
    "O ambiente estava limpo?",
    "Atendimento foi cordial?",
    "Produtos estavam organizados?",
    "Estoque estava adequado?",
    "Fila do caixa foi rápida?",
  ];
  const respostas = ["Sim", "Não", "Parcial"];
  const rows = [];
  for (let ano = 2023; ano <= 2024; ano++) {
    for (let mes = 1; mes <= 12; mes++) {
      lojas.forEach((loja, li) => {
        topicos.forEach((topico, ti) => {
          const tag = tags[(li + ti) % tags.length];
          const questao = questoes[ti];
          const resposta = respostas[Math.floor(Math.random() * 3)];
          const notaPossivel = 10;
          const storeBase = loja === "Loja A" ? 0.92 : loja === "Loja B" ? 0.88 : loja === "Loja C" ? 0.80 : loja === "Loja D" ? 0.75 : 0.70;
          const topicBias = (ti - (topicos.length - 1) / 2) * 0.03;
          const tagBias = (tags.indexOf(tag) - (tags.length - 1) / 2) * 0.02;
          const rawScore = storeBase + topicBias + tagBias + (Math.random() - 0.5) * 0.12;
          const clampedScore = Math.min(0.98, Math.max(0.55, rawScore));
          const notaAtingida = Math.round(notaPossivel * clampedScore * 10) / 10;
          const dia = Math.floor(Math.random() * 28) + 1;
          rows.push({
            Data: new Date(ano, mes - 1, dia),
            Ano: String(ano),
            Mes: String(mes).padStart(2, "0"),
            Loja: loja,
            Topico: topico,
            Tag: tag,
            Questao: questao,
            Resposta: resposta,
            Observacao: Math.random() > 0.7 ? "Verificar com gerência" : "",
            NotaAtingida: Math.max(0, Math.min(notaPossivel, notaAtingida)),
            NotaPossivel: notaPossivel,
          });
        });
      });
    }
  }
  return rows;
}

let FALLBACK_DATA = gerarDados();

async function loadExcelData() {
  try {
    if (typeof window === "undefined") return null;

    if (typeof window.XLSX === "undefined") {
      await new Promise((res, rej) => {
        const s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js';
        s.onload = res;
        s.onerror = rej;
        document.head.appendChild(s);
      });
    }

    const resp = await fetch('/Exemplo.xlsx');
    if (!resp.ok) return null;
    const ab = await resp.arrayBuffer();
    const data = new Uint8Array(ab);
    const workbook = window.XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const raw = window.XLSX.utils.sheet_to_json(sheet, { defval: '' });

    const normalizeMetric = (value) => {
      if (value === null || value === undefined || value === "") return 0;
      if (typeof value === "string") {
        const trimmed = value.trim().replace("%", "").replace(",", ".");
        const parsed = Number(trimmed);
        if (!Number.isFinite(parsed)) return 0;
        if (value.includes("%") || (parsed > 1 && parsed <= 100)) return parsed;
        return parsed;
      }
      if (typeof value === "number") {
        if (value > 1 && value <= 100) return value;
        return value;
      }
      return 0;
    };

    const toScore = (value) => {
      if (value === null || value === undefined || value === '') return null;
      if (typeof value === 'number') {
        if (Number.isNaN(value)) return null;
        return value > 1 && value <= 100 ? value : value;
      }
      if (typeof value === 'string') {
        const parsed = Number(value.trim().replace('%', '').replace(',', '.'));
        if (!Number.isFinite(parsed)) return null;
        return parsed > 1 && parsed <= 100 ? parsed : parsed;
      }
      return null;
    };

    const parseDateCell = (value) => {
      if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return new Date(value.getFullYear(), value.getMonth(), value.getDate());
      }
      if (typeof value === 'number' && Number.isFinite(value)) {
        const d = new Date((value - 25569) * 86400 * 1000);
        if (Number.isNaN(d.getTime())) return null;
        return new Date(d.getFullYear(), d.getMonth(), d.getDate());
      }
      if (typeof value === 'string') {
        const trimmed = value.trim();
        if (!trimmed) return null;
        const dm = trimmed.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})(?:\s|T|$)/);
        if (dm) {
          const day = Number(dm[1]);
          const month = Number(dm[2]);
          let year = Number(dm[3]);
          if (year < 100) year += 2000;
          if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            return new Date(year, month - 1, day);
          }
        }
        const iso = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (iso) {
          return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
        }
        const parsed = new Date(trimmed);
        if (!Number.isNaN(parsed.getTime())) {
          return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
        }
      }
      return null;
    };

    const rows = raw.map(r => {
      const Data = r['Data'] || r['data'] || r['Date'] || r['date'] || r['DATA'];
      const AnoRaw = r['Ano'] || r['ANO'] || r['Year'] || r['year'];
      const MesRaw = r['Mes'] || r['Mês'] || r['MES'] || r['Month'] || r['month'];
      const Loja = r['Loja'] || r['Store'] || r['loja'] || r['LOJA'] || '';
      const Topico = r['Topico'] || r['Tópico'] || r['Topic'] || '';
      const Tag = r['Tag'] || r['tag'] || '';
      const Questao = r['Questao'] || r['Questão'] || r['Question'] || '';
      const Resposta = r['Resposta'] || r['Answer'] || '';
      const Observacao = r['Observacao'] || r['Observação'] || r['Obs'] || '';
      const NotaAtingida = [""].includes(r['NotaAtingida']) ? null : r['Nota Atingida'];
      const NotaPossivel =  [""].includes(r['NotaPossivel']) ? null : r['Nota Possível'];
      const NotaFinal = normalizeMetric(r['Nota Final'] ?? r['NotaFinal'] ?? r['Final']);

      let dateObj = parseDateCell(Data);
      if (!dateObj) {
        const year = Number(AnoRaw);
        const month = Number(MesRaw);
        if (Number.isFinite(year) && Number.isFinite(month) && month >= 1 && month <= 12) {
          dateObj = new Date(year, month - 1, 1);
        }
      }
      if (!dateObj) return null;

      return {
        Data: dateObj,
        Ano: String(dateObj.getFullYear()),
        Mes: String(dateObj.getMonth() + 1).padStart(2, '0'),
        Loja,
        Topico,
        Tag,
        Questao,
        Resposta,
        Observacao,
        NotaAtingida,
        NotaPossivel,
        NotaFinal
      };
    }).filter(Boolean);

    return rows;
  } catch (err) {
    console.warn('Erro carregando Excel', err);
    return null;
  }
}

function WebPortal({ children }) {
  if (typeof document === "undefined") return children;
  const { createPortal } = require("react-dom");
  return createPortal(children, document.body);
}

function colorFor(v) {
  if (v >= 0.85) return "#22C55E";
  if (v >= 0.70) return "#FDBA3B";
  return "#FF4B4B";
}

function fmtPct(v) {
  return (v * 100).toFixed(1).replace(".", ",") + "%";
}

function toPercentValue(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "string") {
    const trimmed = value.trim().replace("%", "").replace(",", ".");
    const parsed = Number(trimmed);
    if (!Number.isFinite(parsed)) return null;
    return parsed > 1 ? parsed / 100 : parsed;
  }
  if (typeof value === "number") {
    return value > 1 ? value / 100 : value;
  }
  return null;
}

function polarToCartesian(cx, cy, r, angleInDegrees) {
  const angleInRadians = (Math.PI / 180) * angleInDegrees;
  return {
    x: cx + r * Math.cos(angleInRadians),
    y: cy + r * Math.sin(angleInRadians),
  };
}

function describeArc(cx, cy, r, startAngle, endAngle) {
  const start = polarToCartesian(cx, cy, r, startAngle);
  const end = polarToCartesian(cx, cy, r, endAngle);
  const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
  return `M ${start.x} ${start.y} A ${r} ${r} 0 ${largeArcFlag} 1 ${end.x} ${end.y}`;
}

function GaugeChart({
  id,
  percent = 0,
  arcsLength = [1],
  colors = ["#ddd"],
  arcWidth = 0.25,
  arcPadding = 0,
  cornerRadius = 0,
  needleColor = "#222",
  needleBaseColor = "#222",
  style,
  meta,
}) {
  const safePercent = Math.min(1, Math.max(0, percent || 0));
  const total = arcsLength.reduce((a, b) => a + b, 0) || 1;
  const sweep = 180;
  const startAngle = 180;
  const padAngle = Math.max(0, arcPadding) * sweep;

  const cx = 100;
  const cy = 100;
  const r = 90;
  const strokeWidth = Math.max(2, r * (arcWidth || 0.25));
  const arcRadius = r - strokeWidth / 2;

  let angle = startAngle;
  const arcs = arcsLength.map((len, i) => {
    const segAngle = (len / total) * sweep;
    const segStart = angle + (i === 0 ? 0 : padAngle / 2);
    const segEnd = angle + segAngle - (i === arcsLength.length - 1 ? 0 : padAngle / 2);
    angle += segAngle;

    if (segEnd <= segStart) return null;
    const d = describeArc(cx, cy, arcRadius, segStart, segEnd);
    return (
      <path
        key={i}
        d={d}
        stroke={colors[i] || colors[colors.length - 1]}
        strokeWidth={strokeWidth}
        fill="none"
        strokeLinecap={cornerRadius > 0 ? "round" : "butt"}
      />
    );
  });

  const needleAngle = startAngle + safePercent * sweep;
  const needleLength = arcRadius * 0.95;
  const needle = polarToCartesian(cx, cy, needleLength, needleAngle);
  const baseRadius = Math.max(4, strokeWidth * 0.15);

  const hasMeta = typeof meta === "number" && meta >= 0 && meta <= 1;
  const metaAngle = startAngle + (meta ?? 0.85) * sweep;
  const metaOuter = polarToCartesian(cx, cy, r + 4, metaAngle);
  const metaInner = polarToCartesian(cx, cy, r - strokeWidth - 4, metaAngle);

  return (
    <div style={style}>
      <svg
        id={id}
        viewBox="0 0 200 120"
        style={{ display: "block", width: "100%", height: 160 }}
        aria-hidden="true"
      >
        <g>
          {arcs}
          {hasMeta && (
            <line
              x1={metaInner.x}
              y1={metaInner.y}
              x2={metaOuter.x}
              y2={metaOuter.y}
              stroke="#222"
              strokeWidth={3}
              strokeLinecap="round"
            />
          )}
          <line x1={cx} y1={cy} x2={needle.x} y2={needle.y} stroke={needleColor} strokeWidth={2.5} />
          <circle cx={cx} cy={cy} r={baseRadius} fill={needleBaseColor} />
        </g>
      </svg>
    </div>
  );
}

function Gauge({ value }) {
  const pct = Math.min(1, Math.max(0, value));
  const color = colorFor(pct);

  let metaValue = 0.85;
  if (pct < 0.75) {
    metaValue = 0.75;
  } else if (pct < 0.85) {
    metaValue = 0.85;
  } else {
    metaValue = 1.0;
  }

  const colors = ["#FF4B4B", "#FDBA3B", "#22C55E"];
  const arcs = [0.70, 0.15, 0.15];

  return (
    <div style={{ position: "relative", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", flex: 1, width: "100%" }}>
      <GaugeChart
        id="gauge-media-meta"
        nrOfLevels={3}
        arcsLength={arcs}
        colors={colors}
        percent={pct}
        arcWidth={0.25}
        arcPadding={0.01}
        cornerRadius={4}
        needleColor="#222222"
        needleBaseColor="#222222"
        hideText={true}
        style={{ width: "100%" }}
        meta={metaValue}
      />
      <div style={{
        textAlign: "center", marginTop: 4, paddingBottom: 4,
        fontWeight: 700, fontSize: 28, color,
        fontFamily: "'Segoe UI', sans-serif"
      }}>
        {fmtPct(pct)}
      </div>
      <div style={{
        textAlign: "center", fontSize: 12, color: "#888", marginBottom: 4,
        fontFamily: "'Segoe UI', sans-serif"
      }}>
        Meta: {fmtPct(metaValue)}
      </div>
    </div>
  );
}

function LineChart({ data, height = 210 }) {
  if (!data.length) return <div style={{ height: 240, display: "flex", alignItems: "center", justifyContent: "center", color: "#aaa", fontSize: 13 }}>Sem dados</div>;
  const W = 560, H = height, PL = 48, PR = 16, PT = 24, PB = 36;
  const vals = data.map(d => d.media);
  const minV = Math.max(0, Math.min(...vals) - 0.05);
  const maxV = Math.min(1, Math.max(...vals) + 0.05);
  const scaleX = (i) => PL + (i / (data.length - 1 || 1)) * (W - PL - PR);
  const scaleY = (v) => PT + ((maxV - v) / (maxV - minV)) * (H - PT - PB);
  const pts = data.map((d, i) => `${scaleX(i)},${scaleY(d.media)}`).join(" ");
  const ticks = [minV, (minV + maxV) / 2, maxV];

  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" height={H} style={{ display: "block" }}>
      {ticks.map((t, i) => {
        const y = scaleY(t);
        return (
          <g key={i}>
            <line x1={PL} y1={y} x2={W - PR} y2={y} stroke="#E5E7EB" strokeWidth="1" strokeDasharray="4,3" />
            <text x={PL - 6} y={y + 4} textAnchor="end" fontSize="10" fill="#888">{fmtPct(t)}</text>
          </g>
        );
      })}
      <polyline points={pts} fill="none" stroke="#A8A8A8" strokeWidth="3" strokeLinejoin="round" />
      {data.map((d, i) => (
        <circle key={i} cx={scaleX(i)} cy={scaleY(d.media)} r="5" fill={colorFor(d.media)} stroke="#fff" strokeWidth="1.5" />
      ))}
      {data.length <= 14 && data.map((d, i) => (
        <text key={i} x={scaleX(i)} y={scaleY(d.media) - 10} textAnchor="middle" fontSize="9" fill={colorFor(d.media)} fontWeight="600">
          {fmtPct(d.media)}
        </text>
      ))}
      {data.map((d, i) => {
        const show = data.length <= 8 || i % Math.ceil(data.length / 8) === 0 || i === data.length - 1;
        if (!show) return null;
        return (
          <text key={i} x={scaleX(i)} y={H - 4} textAnchor="middle" fontSize="9" fill="#888">
            {d.label}
          </text>
        );
      })}
    </svg>
  );
}

function RankingTable({ data, colName }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', minHeight: 220 }}>
      <div style={{ overflowY: "auto", maxHeight: 300, borderRadius: 14, flex: 1, minHeight: 0 }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, fontFamily: "'Segoe UI', sans-serif" }}>
          <thead>
            <tr>
              <th style={{ textAlign: "left", padding: "6px 10px", color: "#222", fontWeight: 700, background: CARD_BG, position: "sticky", top: 0 }}>{colName}</th>
              <th style={{ textAlign: "right", padding: "6px 10px", color: "#222", fontWeight: 700, background: CARD_BG, position: "sticky", top: 0 }}>Média</th>
            </tr>
          </thead>
          <tbody>
            {data.map((row, i) => (
              <tr key={i} style={{ background: i % 2 === 1 ? "#EFEFEF" : CARD_BG }}>
                <td style={{ padding: "6px 10px", color: "#222" }}>{row.name}</td>
                <td style={{ padding: "6px 10px", textAlign: "right", color: colorFor(row.media), fontWeight: 700 }}>
                  {fmtPct(row.media)}
                </td>
              </tr>
            ))}
            {data.length === 0 && (
              <tr><td colSpan={2} style={{ padding: "12px 10px", color: "#aaa", textAlign: "center" }}>Sem dados</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function GeneralTable({ rows }) {
  const cols = ["Data", "Loja", "Tópico", "Tag", "Questão", "Resposta", "Observação"];
  return (
    <div style={{ overflowX: "auto", borderRadius: 14 }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, fontFamily: "'Segoe UI', sans-serif", tableLayout: "fixed", minWidth: 700 }}>
        <thead>
          <tr>
            {cols.map(c => (
              <th key={c} style={{ textAlign: "left", padding: "6px 10px", color: "#222", fontWeight: 700, background: CARD_BG, position: "sticky", top: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 180 }}>{c}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{ background: i % 2 === 1 ? "#EFEFEF" : CARD_BG }}>
              <td style={{ padding: "6px 10px", color: "#333", whiteSpace: "nowrap" }}>{row.Data}</td>
              <td style={{ padding: "6px 10px", color: "#333" }}>{row.Loja}</td>
              <td style={{ padding: "6px 10px", color: "#333", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 120 }}>{row.Topico}</td>
              <td style={{ padding: "6px 10px", color: "#333" }}>{row.Tag}</td>
              <td style={{ padding: "6px 10px", color: "#333", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 200 }}>{row.Questao}</td>
              <td style={{ padding: "6px 10px", color: "#333" }}>{row.Resposta}</td>
              <td style={{ padding: "6px 10px", color: "#333", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 160 }}>{row.Observacao}</td>
            </tr>
          ))}
          {rows.length === 0 && (
            <tr><td colSpan={7} style={{ padding: "16px 10px", color: "#aaa", textAlign: "center" }}>Sem dados</td></tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

function MultiSelect({ label, options, value, onChange }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const toggle = (opt) => {
    if (value.includes(opt)) onChange(value.filter(v => v !== opt));
    else onChange([...value, opt]);
  };

  const displayText = value.length === 0
    ? "Selecione..."
    : value.length <= 2 ? value.join(", ") : `${value.length} selecionados`;

  return (
    <div ref={ref} style={{ position: "relative", marginBottom: 8 }}>
      <div style={{ color: "#fff", fontWeight: 700, fontSize: 12, marginBottom: 4, marginTop: 8 }}>{label}</div>
      <div
        onClick={() => setOpen(o => !o)}
        style={{
          background: "#F2F2F2", borderRadius: 12, border: "1px solid #D0D0D0",
          fontSize: 13, padding: "6px 10px", cursor: "pointer", userSelect: "none",
          display: "flex", justifyContent: "space-between", alignItems: "center", color: "#333"
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{displayText}</span>
        <span style={{ marginLeft: 6, color: "#666", fontSize: 10 }}>{open ? "▲" : "▼"}</span>
      </div>
      {value.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginTop: 4 }}>
          {value.map(v => (
            <span key={v} style={{
              background: "#001d3d", color: "#fff", borderRadius: 8, fontSize: 11,
              padding: "2px 8px", display: "flex", alignItems: "center", gap: 4
            }}>
              {v}
              <span onClick={() => toggle(v)} style={{ cursor: "pointer", fontWeight: 700, opacity: 0.8 }}>×</span>
            </span>
          ))}
        </div>
      )}
      {open && (
        <div style={{
          position: "absolute", top: "100%", left: 0, right: 0, zIndex: 9999,
          background: "#fff", border: "1px solid #D0D0D0", borderRadius: 12,
          maxHeight: 200, overflowY: "auto", boxShadow: "0 8px 24px rgba(0,0,0,0.15)", marginTop: 2
        }}>
          {options.map(opt => (
            <div
              key={opt}
              onClick={() => toggle(opt)}
              style={{
                padding: "7px 12px", fontSize: 13, cursor: "pointer",
                background: value.includes(opt) ? "#e8eef7" : "#fff",
                color: value.includes(opt) ? "#001d3d" : "#333",
                fontWeight: value.includes(opt) ? 700 : 400,
              }}
            >
              {value.includes(opt) ? "✓ " : ""}{opt}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function Card({ title, children }) {
  return (
    <div style={{
      background: CARD_BG, borderRadius: 26, padding: 20,
      boxShadow: "0 4px 15px rgba(0,0,0,0.10)", marginBottom: 15, flex: 1,
      display: "flex", flexDirection: "column"
    }}>
      <div style={{ color: "#222", fontWeight: 700, fontSize: 14, textAlign: "center", marginBottom: 8 }}>{title}</div>
      <div style={{ flex: 1, display: "flex", flexDirection: "column", justifyContent: "center" }}>
        {children}
      </div>
    </div>
  );
}

// ── Main Dashboard ────────────────────────────────────────────────────────────
export default function Dashboard() {
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [dados, setDados] = useState(FALLBACK_DATA);
  const [filtroAno, setFiltroAno] = useState([]);
  const [filtroMes, setFiltroMes] = useState([]);
  const [filtroLoja, setFiltroLoja] = useState([]);
  const [filtroTopico, setFiltroTopico] = useState([]);
  const [filtroTag, setFiltroTag] = useState([]);
  const [rowsPerPage, setRowsPerPage] = useState(20);
  const [pageNum, setPageNum] = useState(1);
  const [isMobile, setIsMobile] = useState(false);

  useEffect(() => {
    let mounted = true;
    loadExcelData().then(d => {
      if (mounted && d && d.length) setDados(d);
    });
    return () => { mounted = false; };
  }, []);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const update = () => setIsMobile(window.innerWidth <= 520);
    update();
    window.addEventListener("resize", update);
    return () => window.removeEventListener("resize", update);
  }, []);

  const anos = useMemo(() => [...new Set(dados.map(d => d.Ano))].sort(), [dados]);
  const meses = useMemo(() => {
    const unique = [...new Set(dados.map(d => parseInt(d.Mes, 10)))];
    unique.sort((a, b) => a - b);
    return unique.map(m => LIST_MESES[m - 1]);
  }, [dados]);
  const lojas = useMemo(() => [...new Set(dados.map(d => d.Loja))].sort(), [dados]);
  const topicos = useMemo(() => [...new Set(dados.map(d => d.Topico))].sort(), [dados]);
  const tags = useMemo(() => [...new Set(dados.map(d => d.Tag))].sort(), [dados]);

  // ── CORREÇÃO: filtroTopico e filtroTag agora estão nas deps de `filtered`
  // e `filtered` é a única fonte de verdade para TODOS os gráficos.
  const filtered = useMemo(() => {
    let d = [...dados];
    if (filtroAno.length) d = d.filter(r => filtroAno.includes(r.Ano));
    if (filtroMes.length) {
      const nums = filtroMes.map(m => LIST_MESES.indexOf(m) + 1);
      d = d.filter(r => nums.includes(parseInt(r.Mes, 10)));
    }
    if (filtroLoja.length) d = d.filter(r => filtroLoja.includes(r.Loja));
    if (filtroTopico.length) d = d.filter(r => filtroTopico.includes(r.Topico));
    if (filtroTag.length) d = d.filter(r => filtroTag.includes(r.Tag));
    return d;
  }, [dados, filtroAno, filtroMes, filtroLoja, filtroTopico, filtroTag]);

  // ── CORREÇÃO: getRowMedia declarada com useCallback antes de todos os useMemos
  // que a utilizam, e incluída como dependência em cada um deles.
  const getRowMedia = useCallback((row) => {
    const finalPercent = toPercentValue(row.NotaFinal);
    // if (finalPercent !== null) return finalPercent;
    const possible = Number(row.NotaPossivel) || 0;
    const attained = Number(row.NotaAtingida) || 0;
    return possible > 0 ? attained / possible : 0;
  }, []); // sem deps externas: só usa os campos do próprio `row`

  // ── Todos os useMemos agora listam `getRowMedia` nas dependências ──────────

  const mediaGeral = useMemo(() => {
    if (!filtered.length) return 0;
    const notaAtingida = filtered.reduce((a, b) => a + Number(b.NotaAtingida) || 0, 0);
    const notaPossivel = filtered.reduce((a, b) => a + Number(b.NotaPossivel) || 0, 0);
    return notaPossivel > 0 ? notaAtingida / notaPossivel : 0;
  }, [filtered]);

  const lineData = useMemo(() => {
    const map = {};
    filtered.forEach(r => {
      const k = r.Mes + "/" + r.Ano;
      const obtainedWeight = Number(r.NotaAtingida) || 0;
      const possibleWeight = Number(r.NotaPossivel) || 0;
     
      if (!map[k]) map[k] = { label: k, totalWeighted: 0, weightSum: 0, ano: r.Ano, mes: r.Mes };

      map[k].totalWeighted += possibleWeight;
      map[k].weightSum += obtainedWeight;
    });
    return Object.values(map)
      .sort((a, b) => a.ano !== b.ano ? a.ano.localeCompare(b.ano) : a.mes.localeCompare(b.mes))
      .map(d => ({ ...d, media: d.weightSum > 0 ? d.weightSum / d.totalWeighted : 0 }));
  }, [filtered]);

  const rankLojas = useMemo(() => {
    const map = {};
    filtered.forEach(r => {
      const obtainedWeight = Number(r.NotaAtingida) || 0;
      const possibleWeight = Number(r.NotaPossivel) || 0;


      if (!map[r.Loja]) map[r.Loja] = { totalWeighted: 0, weightSum: 0 };
      map[r.Loja].totalWeighted += possibleWeight;
      map[r.Loja].weightSum += obtainedWeight;
    });
    return Object.entries(map)
      .map(([name, d]) => ({ name, media: d.weightSum > 0 ? d.weightSum / d.totalWeighted : 0 }))
      .sort((a, b) => b.media - a.media);
  }, [filtered]);

  const rankTopicos = useMemo(() => {
    const map = {};

    filtered.forEach(r => {
      if (!map[r.Topico]) {
        map[r.Topico] = {
          atingida: 0,
          possivel: 0,
        };
      }

      map[r.Topico].atingida += Number(r.NotaAtingida) || 0;
      map[r.Topico].possivel += Number(r.NotaPossivel) || 0;
    });

    return Object.entries(map)
      .map(([name, d]) => ({
        name,
        media:
          d.possivel > 0
            ? d.atingida / d.possivel
            : 0,
      }))
      .sort((a, b) => b.media - a.media);
  }, [filtered]);

  const rankTags = useMemo(() => {
    const map = {};

    filtered.forEach(r => {
      if (!map[r.Tag]) {
        map[r.Tag] = {
          atingida: 0,
          possivel: 0,
        };
      }

      map[r.Tag].atingida += Number(r.NotaAtingida) || 0;
      map[r.Tag].possivel += Number(r.NotaPossivel) || 0;
    });

    return Object.entries(map)
      .map(([name, d]) => ({
        name,
        media:
          d.possivel > 0
            ? d.atingida / d.possivel
            : 0,
      }))
      .sort((a, b) => b.media - a.media);
  }, [filtered]);

  const tabelaBase = useMemo(() => {

    const topicoScore = {};
    const tagScore = {};

    let groups = {};
    filtered.forEach(r => {
      const key = `${r.Loja}|${r.Data}|${r.Topico}`;
      if (!groups[key]) groups[key] = { atingida: 0, possivel: 0 };
      groups[key].atingida += Number(r.NotaAtingida) || 0;
      groups[key].possivel += Number(r.NotaPossivel) || 0;
    });
    Object.entries(groups).forEach(([key, g]) => {
      topicoScore[key] = g.possivel > 0 ? g.atingida / g.possivel : 0;
    });
    


    groups = {};
    filtered.forEach(r => {
      const key = `${r.Loja}|${r.Data}|${r.Tag}`;
      if (!groups[key]) groups[key] = { atingida: 0, possivel: 0 };
      groups[key].atingida += Number(r.NotaAtingida) || 0;
      groups[key].possivel += Number(r.NotaPossivel) || 0;
    });
    Object.entries(groups).forEach(([key, g]) => {
      tagScore[key] = g.possivel > 0 ? g.atingida / g.possivel : 0;
    });
    

    return filtered.map(r => {
      const tkKey = `${r.Loja}|${r.Data}|${r.Topico}`;
      const tgKey = `${r.Loja}|${r.Data}|${r.Tag}`;
      return {
        Data: r.Data.toLocaleDateString("pt-BR"),
        Loja: r.Loja,
        Topico: r.Topico,
        Tag: r.Tag,
        Questao: r.Questao,
        Resposta: r.Resposta,
        Observacao: r.Observacao,
      };
    });
  }, [filtered]);

  const totalRows = tabelaBase.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / rowsPerPage));
  const safePage = Math.min(pageNum, totalPages);
  const tabelaPage = tabelaBase.slice((safePage - 1) * rowsPerPage, safePage * rowsPerPage);

  return (
    <div style={{ background: APP_BG, minHeight: "100vh", fontFamily: "'Segoe UI', sans-serif", display: "flex", flexDirection: "column" }}>
      <WebPortal>
        <>
          <button
            onClick={() => setSidebarOpen(o => !o)}
            style={{
              position: "fixed", top: 12, left: 12, zIndex: 2000,
              background: SIDEBAR_BG, color: "#fff", border: "none",
              borderRadius: 10, padding: "6px 10px", fontSize: 18, lineHeight: 1,
              boxShadow: "0 6px 16px rgba(0,0,0,0.20)", cursor: "pointer"
            }}
          >
            ☰
          </button>

          {sidebarOpen && (
            <div
              onClick={() => setSidebarOpen(false)}
              style={{ position: "fixed", inset: 0, zIndex: 1400, background: "rgba(0,0,0,0.3)" }}
            />
          )}

          <div style={{
            position: "fixed", top: 0, left: sidebarOpen ? 0 : -320, width: 280, height: "100vh",
            background: SIDEBAR_BG, padding: "1.5rem 1rem", zIndex: 1500,
            borderRight: "1px solid rgba(255,255,255,0.08)", boxShadow: "4px 0 16px rgba(0,0,0,0.15)",
            overflowY: "auto", overflowX: "hidden",
            transition: "left 0.25s ease"
          }}>
            <div style={{ color: "#fff", fontWeight: 700, fontSize: 16, marginBottom: 20, textAlign: "center", letterSpacing: "0.04em" }}>
              🔎 Filtros
            </div>
            <MultiSelect label="Ano" options={anos} value={filtroAno} onChange={setFiltroAno} />
            <MultiSelect label="Mês" options={meses} value={filtroMes} onChange={setFiltroMes} />
            <MultiSelect label="Loja" options={lojas} value={filtroLoja} onChange={setFiltroLoja} />
            <MultiSelect label="Tópico" options={topicos} value={filtroTopico} onChange={setFiltroTopico} />
            <MultiSelect label="Tag" options={tags} value={filtroTag} onChange={setFiltroTag} />
            <button
              onClick={() => { setFiltroAno([]); setFiltroMes([]); setFiltroLoja([]); setFiltroTopico([]); setFiltroTag([]); }}
              style={{
                marginTop: 20, width: "100%", background: "rgba(255,255,255,0.12)", color: "#fff",
                border: "1px solid rgba(255,255,255,0.25)", borderRadius: 12, padding: "8px 0",
                fontSize: 13, cursor: "pointer", fontWeight: 600
              }}
            >
              Limpar filtros
            </button>
          </div>
        </>
      </WebPortal>

      <div style={{ flex: 1, overflowY: "auto", overflowX: "hidden", padding: "0 1.5rem 2rem 1.5rem", scrollBehavior: "smooth" }}>
        <h1 style={{ color: "#222", fontWeight: 700, fontSize: 26, textAlign: "center", padding: "1.2rem 0 1rem 0", letterSpacing: "0.01em" }}>
          📊 Dashboard Exemplo
        </h1>

        <div style={{ display: "flex", gap: 15, flexWrap: "wrap", alignItems: "stretch" }}>
          <div style={{ flex: isMobile ? "1 1 100%" : "0 0 calc(40% - 8px)", minWidth: isMobile ? "100%" : 240, display: "flex" }}>
            <Card title="Média / Meta">
              <Gauge value={mediaGeral} />
            </Card>
          </div>
          <div style={{ flex: isMobile ? "1 1 100%" : "0 0 calc(60% - 10px)", minWidth: isMobile ? "100%" : 280, display: "flex" }}>
            <Card title="Média por Ano e Mês">
              <LineChart data={lineData} height={isMobile ? 220 : 210} />
            </Card>
          </div>
        </div>

        <div style={{ display: "flex", gap: 15, marginBottom: 15, height: "auto", flexWrap: "wrap" }}>
          {[
            { title: "Ranking por Loja", data: rankLojas, col: "Loja" },
            { title: "Ranking por Tópico", data: rankTopicos, col: "Tópico" },
            { title: "Ranking por Tag", data: rankTags, col: "Tag" },
          ].map(({ title, data, col }) => (
            <div key={col} style={{ flex: "1 1 calc(33% - 10px)", minWidth: 200, display: 'flex', flexDirection: 'column' }}>
              <Card title={title}>
                <RankingTable data={data} colName={col} />
              </Card>
            </div>
          ))}
        </div>

        <div style={{ background: CARD_BG, borderRadius: 26, padding: 20, boxShadow: "0 4px 15px rgba(0,0,0,0.10)" }}>
          <div style={{ color: "#222", fontWeight: 700, fontSize: 14, textAlign: "center", marginBottom: 12 }}>Tabela Geral</div>

          <div style={{ display: "flex", gap: 16, marginBottom: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
            <div>
              <div style={{ color: "#222", fontWeight: 700, fontSize: 12, marginBottom: 4 }}>Linhas por página</div>
              <select
                value={rowsPerPage}
                onChange={e => { setRowsPerPage(Number(e.target.value)); setPageNum(1); }}
                style={{
                  width: 90, borderRadius: 12, border: "1px solid #D0D0D0", background: "#F2F2F2",
                  fontSize: 13, padding: "5px 8px"
                }}
              >
                {[10, 20, 50, 100].map(n => <option key={n} value={n}>{n}</option>)}
              </select>
            </div>
            <div>
              <div style={{ color: "#222", fontWeight: 700, fontSize: 12, marginBottom: 4 }}>Página</div>
              <input
                type="number" min={1} max={totalPages} value={safePage}
                onChange={e => setPageNum(Math.max(1, Math.min(totalPages, Number(e.target.value))))}
                style={{
                  width: 70, borderRadius: 12, border: "1px solid #D0D0D0", background: "#F2F2F2",
                  fontSize: 13, padding: "5px 8px"
                }}
              />
            </div>
            <div style={{ color: "#555", fontSize: 12 }}>
              Página {safePage} de {totalPages} | {totalRows} linhas
            </div>
          </div>

          <GeneralTable rows={tabelaPage} />
        </div>
      </div>
    </div>
  );
}
