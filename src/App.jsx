import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  PencilRuler,
  Printer,
  ArrowLeft,
  Plus,
  FileSpreadsheet,
  AlertTriangle,
  Trash2,
  FileText,
} from "lucide-react";

/* =========================================================================
   Croqui de Tubos · Amostra (Colapso)
   - Tela inicial: importar (xls) OU inserir manual
   - Cada croqui = 1 página A4 paisagem
   - Amostra centralizada na "posição do colapso" (zero = extremidade esquerda)
   ========================================================================= */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root{
  --bg:#0a0e17; --panel:#111726; --panel2:#0d131f;
  --line:#1d2738; --line2:#2a3850;
  --ink:#0f1115; --paper:#ffffff;
  --txt:#e7eef9; --muted:#8197b8; --muted2:#5e7088;
  --accent:#38bdf8; --accent-d:#0ea5e9;
  --signal:#f0563f;
}
*{box-sizing:border-box}
.ct-root{
  font-family:'IBM Plex Sans',sans-serif;
  background:var(--bg);
  color:var(--txt);
  min-height:100vh;
  background-image:
    linear-gradient(rgba(56,189,248,.045) 1px,transparent 1px),
    linear-gradient(90deg,rgba(56,189,248,.045) 1px,transparent 1px);
  background-size:26px 26px;
}
.mono{font-family:'IBM Plex Mono',monospace}
.ct-wrap{max-width:1180px;margin:0 auto;padding:28px 20px 80px}

/* header */
.ct-head{display:flex;align-items:center;justify-content:space-between;gap:16px;
  border-bottom:1px solid var(--line);padding-bottom:18px;margin-bottom:34px;flex-wrap:wrap}
.ct-brand{display:flex;align-items:center;gap:14px}
.ct-mark{width:42px;height:42px;border:1px solid var(--line2);border-radius:9px;
  display:flex;align-items:center;justify-content:center;background:var(--panel);
  color:var(--accent);box-shadow:0 0 0 4px rgba(56,189,248,.05)}
.ct-title{font-size:18px;font-weight:700;letter-spacing:.2px;line-height:1.1}
.ct-sub{font-size:12px;color:var(--muted);letter-spacing:.5px;text-transform:uppercase}
.ct-cred{display:flex;flex-direction:column;align-items:flex-start;line-height:1.3}
.ct-cred-top{font-size:10.5px;color:var(--muted)}
.ct-cred-name{font-size:10px;font-weight:600;letter-spacing:.4px;color:var(--accent);text-transform:uppercase;margin-top:2px}

/* home cards */
.ct-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:720px){.ct-grid{grid-template-columns:1fr}}
.ct-card{position:relative;border:1px solid var(--line);background:linear-gradient(180deg,var(--panel),var(--panel2));
  border-radius:14px;padding:26px;cursor:pointer;transition:.18s;overflow:hidden}
.ct-card:hover{border-color:var(--accent-d);transform:translateY(-2px);
  box-shadow:0 18px 40px -24px rgba(56,189,248,.5)}
.ct-card::after{content:"";position:absolute;inset:0;background:
  radial-gradient(420px 120px at 80% -10%,rgba(56,189,248,.12),transparent 70%);pointer-events:none}
.ct-ic{width:46px;height:46px;border-radius:10px;background:rgba(56,189,248,.1);
  color:var(--accent);display:flex;align-items:center;justify-content:center;margin-bottom:18px}
.ct-card h3{margin:0 0 6px;font-size:17px;font-weight:600}
.ct-card p{margin:0;color:var(--muted);font-size:13.5px;line-height:1.5}
.ct-tag{display:inline-block;margin-top:16px;font-size:11px;letter-spacing:1px;
  text-transform:uppercase;color:var(--accent)}

.ct-lead{font-size:13px;color:var(--muted);margin:0 0 24px;max-width:640px;line-height:1.6}

/* form */
.ct-panel{border:1px solid var(--line);background:var(--panel);border-radius:14px;padding:26px}
.ct-frm{display:grid;grid-template-columns:1fr 1fr;gap:16px 20px}
@media(max-width:640px){.ct-frm{grid-template-columns:1fr}}
.fld{display:flex;flex-direction:column;gap:7px}
.fld label{font-size:12px;color:var(--muted);letter-spacing:.4px;text-transform:uppercase;font-weight:500}
.fld .hint{font-size:11px;color:var(--muted2,#5e7088);text-transform:none;letter-spacing:0}
.inp{background:var(--panel2);border:1px solid var(--line2);color:var(--txt);
  border-radius:9px;padding:11px 13px;font-size:15px;font-family:'IBM Plex Mono',monospace;
  outline:none;transition:.15s;width:100%}
.inp:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(56,189,248,.15)}
select.inp{appearance:none;-webkit-appearance:none;cursor:pointer;padding-right:34px;
  background-image:linear-gradient(45deg,transparent 50%,var(--muted) 50%),
    linear-gradient(135deg,var(--muted) 50%,transparent 50%);
  background-position:calc(100% - 18px) calc(50% - 1px),calc(100% - 13px) calc(50% - 1px);
  background-size:5px 5px,5px 5px;background-repeat:no-repeat}
.unit{position:absolute;right:12px;top:50%;transform:translateY(-50%);font-size:12px;color:var(--muted)}
.fld .ip{position:relative}

.pos-block{display:grid;grid-template-columns:1fr 1fr;gap:16px 20px;margin-top:18px;
  border-top:1px solid var(--line);padding-top:18px}
@media(max-width:640px){.pos-block{grid-template-columns:1fr}}
.seg{display:inline-flex;border:1px solid var(--line2);border-radius:9px;overflow:hidden;width:max-content}
.seg button{background:var(--panel2);color:var(--muted);border:none;padding:10px 18px;
  font-family:inherit;font-size:14px;font-weight:600;cursor:pointer;transition:.15s}
.seg button+button{border-left:1px solid var(--line2)}
.seg button.on{background:var(--accent);color:#04121d}
.seg button:not(.on):hover{color:var(--txt)}

.btn{display:inline-flex;align-items:center;gap:9px;border:1px solid var(--line2);
  background:var(--panel2);color:var(--txt);padding:11px 18px;border-radius:9px;
  font-size:14px;font-weight:600;cursor:pointer;transition:.15s;font-family:inherit}
.btn:hover{border-color:var(--accent-d)}
.btn.primary{background:var(--accent);border-color:var(--accent);color:#04121d}
.btn.primary:hover{background:var(--accent-d)}
.btn.ghost{background:transparent}
.ct-actions{display:flex;gap:12px;flex-wrap:wrap;margin-top:24px;align-items:center}

.back{display:inline-flex;align-items:center;gap:8px;color:var(--muted);font-size:13px;
  cursor:pointer;margin-bottom:22px;background:none;border:none;font-family:inherit}
.back:hover{color:var(--accent)}

/* dropzone */
.drop{border:1.5px dashed var(--line2);border-radius:14px;padding:48px 24px;text-align:center;
  background:var(--panel2);cursor:pointer;transition:.15s}
.drop:hover{border-color:var(--accent);background:rgba(56,189,248,.04)}
.drop .di{color:var(--accent);margin-bottom:14px}
.drop h4{margin:0 0 6px;font-size:16px}
.drop p{margin:0;color:var(--muted);font-size:13px}

.note{display:flex;gap:11px;align-items:flex-start;border:1px solid var(--line2);
  background:rgba(240,86,63,.06);border-color:rgba(240,86,63,.3);
  border-radius:10px;padding:13px 15px;margin-top:18px;font-size:12.5px;color:#f4b6ac;line-height:1.5}
.note.info{background:rgba(56,189,248,.06);border-color:rgba(56,189,248,.3);color:#9fd8f5}
.note svg{flex:none;margin-top:1px}

/* result toolbar */
.ct-bar{display:flex;align-items:center;justify-content:space-between;gap:14px;
  margin-bottom:24px;flex-wrap:wrap}
.ct-count{font-size:13px;color:var(--muted)}
.ct-count b{color:var(--txt)}

/* A4 pages */
.pages{display:flex;flex-direction:column;gap:30px;align-items:center}
.page{width:100%;max-width:1120px;aspect-ratio:297/210;background:var(--paper);
  border-radius:6px;overflow:hidden;box-shadow:0 24px 60px -30px rgba(0,0,0,.8);position:relative}
.page svg{display:block;width:100%;height:100%}
.page-tag{position:absolute;top:3px;right:6px;font-size:8px;color:#9aa6b8;
  letter-spacing:1px;z-index:2}

/* rodapé com link do manual */
.ct-foot{margin-top:48px;padding-top:18px;border-top:1px solid var(--line);
  display:flex;justify-content:center}
.ct-foot a{display:inline-flex;align-items:center;gap:8px;color:var(--muted2);
  font-size:14px;text-decoration:none;transition:.15s;letter-spacing:.3px}
.ct-foot a:hover{color:var(--accent)}

@media print{
  @page{size:A4 landscape;margin:0}
  .ct-root{background:#fff !important;background-image:none !important}
  .no-print{display:none !important}
  .ct-wrap{padding:0;max-width:none}
  .pages{gap:0}
  .page{width:297mm;height:210mm;max-width:none;aspect-ratio:auto;border-radius:0;
    box-shadow:none;break-after:page;page-break-after:always}
  .page:last-child{break-after:auto}
}
`;

const nf = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 2 });
const fmt = (n) => (isFinite(n) ? nf.format(n) : "—");

// arredonda para a dezena superior mais próxima (915 -> 920, 900 -> 900)
const roundUpTen = (v) => Math.ceil(v / 10) * 10;

// arredonda para a dezena mais próxima (5043 -> 5040, 5046 -> 5050)
const roundNearestTen = (v) => Math.round(v / 10) * 10;

// tamanho fixo da amostra de colapso conforme OD e pressão spec (Psi)
// retorna null quando fora dos critérios
function collapseSampleLen(od, psi) {
  if (od > 355.6) return 5400; // OD > 355,6 — qualquer pressão
  if (psi > 11600 && od >= 273.1 && od <= 473.1) return 5400; // 273,1–473,1 e psi > 11600
  if (psi <= 11600 && od >= 298.4 && od <= 355.6) return 3600; // 298,4–355,6 e psi ≤ 11600
  if (od >= 114.3 && od <= 273.1) return 2800; // 114,3–273,1 — qualquer pressão
  return null;
}

// Planeja o lado (esq/dir) de cada colapso, exigindo espaço para
// tensão residual (3xOD↑dezena) + tração (380mm) e evitando sobreposição
// com amostras já posicionadas. Ambos os lados livres -> aleatório.
function planCollapses(od, L, collapses) {
  const residualLen = roundUpTen(3 * od);
  const tracaoLen = 380;
  const cl = (v) => Math.max(0, Math.min(L, v));
  const occ = [];
  const hit = (a, b) =>
    occ.some(([x, y]) => Math.min(a, b) < y && Math.max(a, b) > x);
  const groups = collapses.map((c) => ({
    colStart: cl(c.pos - c.len / 2),
    colEnd: cl(c.pos + c.len / 2),
  }));
  groups.forEach((g) => occ.push([g.colStart, g.colEnd]));
  return groups.map((g) => {
    const right = { side: "right", span: [g.colEnd, g.colEnd + residualLen + tracaoLen] };
    right.ok = right.span[1] <= L && !hit(right.span[0], right.span[1]);
    const left = { side: "left", span: [g.colStart - residualLen - tracaoLen, g.colStart] };
    left.ok = left.span[0] >= 0 && !hit(left.span[0], left.span[1]);
    let chosen;
    if (right.ok && left.ok) chosen = Math.random() < 0.5 ? right : left;
    else if (right.ok) chosen = right;
    else if (left.ok) chosen = left;
    else chosen = L - g.colEnd >= g.colStart ? right : left;
    occ.push(chosen.span);
    return chosen.side;
  });
}

// Layout manual: todas as amostras de colapso têm o mesmo comprimento (colLen).
// Apenas o 1º colapso tem posição (fixa = valor informado; aleatória = sorteada).
// Os demais entram EM SEQUÊNCIA, contíguos às amostras anteriores, no mesmo lado.
function buildManualCollapses({ od, L, n, colLen, posMode, firstPos }) {
  const residualLen = roundUpTen(3 * od);
  const tracaoLen = 380;
  const groupLen = colLen + residualLen + tracaoLen; // colapso + tensão residual + tração
  const chainLen = n * groupLen; // do início do 1º colapso até o fim da última tração

  let side;
  let pos1 = firstPos;

  if (posMode === "aleatoria") {
    side = Math.random() < 0.5 ? "right" : "left";
    if (chainLen <= L) {
      if (side === "right") {
        const c1Start = Math.random() * (L - chainLen);
        pos1 = Math.round(c1Start + colLen / 2);
      } else {
        const c1End = chainLen + Math.random() * (L - chainLen);
        pos1 = Math.round(c1End - colLen / 2);
      }
    } else {
      pos1 = side === "right" ? colLen / 2 : L - colLen / 2;
    }
  } else {
    const c1Start = firstPos - colLen / 2;
    const c1End = firstPos + colLen / 2;
    const rightFits = c1Start >= 0 && c1Start + chainLen <= L;
    const leftFits = c1End <= L && c1End - chainLen >= 0;
    if (rightFits && leftFits) side = Math.random() < 0.5 ? "right" : "left";
    else if (rightFits) side = "right";
    else if (leftFits) side = "left";
    else side = L - c1Start >= c1End ? "right" : "left"; // melhor esforço
  }

  pos1 = roundNearestTen(pos1);

  const out = [];
  if (side === "right") {
    // 1ª amostra: mantém o comprimento total; se ultrapassar o zero, desloca
    // para a direita (referência de centro fica descentralizada, sem corte)
    let start0 = pos1 - colLen / 2;
    if (start0 < 0) start0 = 0;
    if (colLen > L) start0 = 0;
    let cursor = start0;
    for (let k = 0; k < n; k++) {
      const start = cursor;
      const ref = k === 0 ? pos1 : start + colLen / 2;
      out.push({ start, len: colLen, ref, side });
      cursor = start + groupLen;
    }
  } else {
    let end0 = pos1 + colLen / 2;
    if (end0 > L) end0 = L;
    if (colLen > L) end0 = L;
    let cursor = end0;
    for (let k = 0; k < n; k++) {
      const end = cursor;
      const start = end - colLen;
      const ref = k === 0 ? pos1 : start + colLen / 2;
      out.push({ start, len: colLen, ref, side });
      cursor = start - residualLen - tracaoLen;
    }
  }
  return out;
}

/* ---------------------------------------------------------------- SVG croqui */
function Croqui({ s, index, total }) {
  // viewBox em unidades de desenho
  const VB_W = 1120,
    VB_H = 740;
  const mL = 150,
    mR = 70;
  const drawW = VB_W - mL - mR;
  const cy = 322;
  const tubeH = 112;
  const top = cy - tubeH / 2;
  const bot = cy + tubeH / 2;

  const L = Math.max(s.length, 1);
  const sx = drawW / L;
  const X = (mm) => mL + mm * sx;

  const idIn = Math.max(s.od - 2 * s.wt, 0);
  const wallVis = Math.min(Math.max((s.wt / s.od) * tubeH, 7), 24);
  const endRx = 9; // elipse da extremidade (sugestão de tubo circular)

  const ink = "#0f1115";
  const sig = "#e1442c"; // colapso
  const sigRes = "#c8841c"; // tensão residual
  const sigTra = "#2f8f6b"; // tração

  const cl = (v) => Math.max(0, Math.min(L, v));
  const residualLen = roundUpTen(3 * s.od); // tensão residual = 3 x OD (dezena superior)
  const tracaoLen = 380; // tração = fixo 380 mm

  // lista de colapsos (compat. com croqui antigo de colapso único)
  const rawCollapses =
    s.collapses && s.collapses.length
      ? s.collapses
      : [{ pos: s.collapsePos, len: s.sampleLength, side: s.side }];

  // geometria de cada grupo: [colapso][tensão residual][tração]
  const groups = rawCollapses.map((c, i) => {
    const len = c.len;
    // início real da amostra de colapso: usa 'start' quando disponível;
    // senão (formato antigo) centraliza em 'pos'
    let start = c.start != null ? c.start : c.pos - len / 2;
    let ref = c.ref != null ? c.ref : c.pos != null ? c.pos : start + len / 2;
    // manter comprimento total deslocando se ultrapassar as bordas (sem corte)
    if (len <= L) {
      if (start < 0) start = 0;
      else if (start + len > L) start = L - len;
    } else {
      start = 0;
    }
    const colStart = Math.max(0, start);
    const colEnd = Math.min(L, start + len);
    const side = c.side === "left" ? "left" : "right";
    let resStart, resEnd, traStart, traEnd;
    if (side === "right") {
      resStart = colEnd;
      resEnd = colEnd + residualLen;
      traStart = resEnd;
      traEnd = resEnd + tracaoLen;
    } else {
      resEnd = colStart;
      resStart = colStart - residualLen;
      traEnd = resStart;
      traStart = resStart - tracaoLen;
    }
    const overflow =
      len > L ||
      Math.min(resStart, traStart) < 0 ||
      Math.max(resEnd, traEnd) > L;
    // extensão total do grupo (colapso + tensão residual + tração) p/ o colchete
    const gMin = Math.max(0, Math.min(colStart, resStart, traStart));
    const gMax = Math.min(L, Math.max(colEnd, resEnd, traEnd));
    const posicao = c.posicao || s.posicao || "—";
    const tipo = c.tipo || s.tipoColapso || "INTERNO (sem acompanhamento)";
    return { i, ref, len, side, posicao, tipo, gMin, gMax, colStart, colEnd,
      resStart, resEnd, traStart, traEnd, overflow };
  });
  const anyWarn = groups.some((g) => g.overflow);

  const produto =
    s.produto ||
    `${fmt(s.od)} x ${fmt(s.wt)}${s.grau ? " - " + s.grau : ""}`;
  const colLenVal = (groups[0] && groups[0].len) || 0;

  const Band = ({ a, b, color, pattern }) => (
    <rect x={X(cl(a))} y={top} width={Math.max(X(cl(b)) - X(cl(a)), 0)}
      height={tubeH} fill={`url(#${pattern})`} stroke={color} strokeWidth="1.4" />
  );

  // rótulo do nome do segmento, acima da amostra (escalonado por nível)
  const Tag = ({ a, b, color, text, level }) => {
    const cx = (X(cl(a)) + X(cl(b))) / 2;
    const ly = top - 6 - level * 24;
    const tw = text.length * 9.5 + 16;
    return (
      <g>
        <line x1={cx} y1={top} x2={cx} y2={ly} stroke={color} strokeWidth="0.7"
          strokeDasharray="2 2" />
        <circle cx={cx} cy={top} r="2" fill={color} />
        <rect x={cx - tw / 2} y={ly - 17} width={tw} height="22" rx="3"
          fill="#fff" stroke={color} strokeWidth="0.9" />
        <text x={cx} y={ly - 2} textAnchor="middle" fontSize="18" fontWeight="600"
          fontFamily="'IBM Plex Sans',sans-serif" fill={color}>{text}</text>
      </g>
    );
  };

  // colchete que agrupa colapso + tensão residual + tração de uma posição
  const PosBracket = ({ a, b, label }) => {
    const x1 = X(cl(a));
    const x2 = X(cl(b));
    if (x2 - x1 < 2) return null;
    const yB = top - 86; // linha do colchete
    const drop = 9; // pernas descendo em direção ao tubo
    return (
      <g>
        <path
          d={`M ${x1} ${yB + drop} L ${x1} ${yB} L ${x2} ${yB} L ${x2} ${yB + drop}`}
          fill="none" stroke="#5a6a82" strokeWidth="1.4" />
        <text x={(x1 + x2) / 2} y={yB - 7} textAnchor="middle" fontSize="14.5"
          fontWeight="700" fontFamily="'IBM Plex Sans',sans-serif" fill="#5a6a82">
          {label}
        </text>
      </g>
    );
  };

  const Dim = ({ x1, x2, y, label, color = ink }) => {
    const mid = (x1 + x2) / 2;
    return (
      <g>
        <line x1={x1} y1={y} x2={x2} y2={y} stroke={color} strokeWidth="1"
          markerStart="url(#arr)" markerEnd="url(#arr)" />
        <rect x={mid - 78} y={y - 14} width="156" height="20" fill="#fff" opacity="0.92" />
        <text x={mid} y={y + 1} textAnchor="middle" fontSize="18"
          fontFamily="'IBM Plex Mono',monospace" fill={color}>{label}</text>
      </g>
    );
  };

  return (
    <svg viewBox={`0 0 ${VB_W} ${VB_H}`} preserveAspectRatio="xMidYMid meet">
      <defs>
        <marker id="arr" markerWidth="10" markerHeight="10" refX="5" refY="5"
          orient="auto-start-reverse">
          <path d="M9,2 L1,5 L9,8" fill="none" stroke={ink} strokeWidth="1.1" />
        </marker>
        {[
          ["hatch", sig],
          ["hatchRes", sigRes],
          ["hatchTra", sigTra],
        ].map(([id, c]) => (
          <pattern key={id} id={id} width="9" height="9" patternUnits="userSpaceOnUse"
            patternTransform="rotate(45)">
            <rect width="9" height="9" fill={c} opacity="0.10" />
            <line x1="0" y1="0" x2="0" y2="9" stroke={c} strokeWidth="1.4" opacity="0.55" />
          </pattern>
        ))}
      </defs>

      {/* borda da prancha */}
      <rect x="14" y="14" width={VB_W - 28} height={VB_H - 28} fill="none"
        stroke="#c9d2df" strokeWidth="1.2" rx="3" />

      {/* cabeçalho */}
      <text x="34" y="48" fontSize="20" fontWeight="700" fill={ink}
        fontFamily="'IBM Plex Sans',sans-serif">CROQUI AMOSTRA DE COLAPSO</text>
      <line x1="34" y1="64" x2={VB_W - 34} y2="64" stroke="#dde3ec" strokeWidth="1" />

      {/* seção de identificação no canto superior esquerdo */}
      {[
        ["PRODUTO", produto],
        ["PEDIDO/ITEM", s.pedidoItem || "—"],
        ["ORDEM DE PRODUÇÃO", s.ordemProducao || "—"],
        ["IPPN", s.ippn || "—"],
        ["PRESSÃO SPEC", `${fmt(s.psi)} psi`],
      ].map((r, i) => (
        <text key={i} x="34" y={84 + i * 18} fontSize="12.5"
          fontFamily="'IBM Plex Sans',sans-serif">
          <tspan fill="#5a6a82" fontWeight="600">{r[0]}: </tspan>
          <tspan fill={ink}>{r[1]}</tspan>
        </text>
      ))}

      {/* ---- TUBO (vista lateral) ---- */}
      {/* corpo */}
      <rect x={X(0)} y={top} width={X(L) - X(0)} height={tubeH} fill="#f3f6fa"
        stroke={ink} strokeWidth="1.6" />
      {/* paredes (faixas) */}
      <rect x={X(0)} y={top} width={X(L) - X(0)} height={wallVis} fill="#e2e8f1" />
      <rect x={X(0)} y={bot - wallVis} width={X(L) - X(0)} height={wallVis} fill="#e2e8f1" />
      <line x1={X(0)} y1={top + wallVis} x2={X(L)} y2={top + wallVis} stroke={ink} strokeWidth="0.8" />
      <line x1={X(0)} y1={bot - wallVis} x2={X(L)} y2={bot - wallVis} stroke={ink} strokeWidth="0.8" />
      {/* extremidades elípticas (sugestão circular) */}
      <ellipse cx={X(L)} cy={cy} rx={endRx} ry={tubeH / 2} fill="#eef2f7" stroke={ink} strokeWidth="1.6" />
      <ellipse cx={X(L)} cy={cy} rx={endRx * (idIn / s.od)} ry={(tubeH - 2 * wallVis) / 2}
        fill="#f3f6fa" stroke={ink} strokeWidth="0.8" />
      <path d={`M ${X(0)} ${top} A ${endRx} ${tubeH / 2} 0 0 0 ${X(0)} ${bot}`}
        fill="none" stroke={ink} strokeWidth="0.8" opacity="0.6" />

      {/* ---- AMOSTRAS por colapso: colapso + tensão residual + tração ---- */}
      {groups.map((g) => (
        <g key={`b${g.i}`}>
          <Band a={g.colStart} b={g.colEnd} color={sig} pattern="hatch" />
          <Band a={g.resStart} b={g.resEnd} color={sigRes} pattern="hatchRes" />
          <Band a={g.traStart} b={g.traEnd} color={sigTra} pattern="hatchTra" />
        </g>
      ))}

      {/* nome do segmento acima de cada amostra (item 8) */}
      {groups.map((g) => (
        <g key={`t${g.i}`}>
          <Tag a={g.colStart} b={g.colEnd} color={sig} level={2}
            text={groups.length > 1 ? `COLAPSO ${g.i + 1}` : "COLAPSO"} />
          <Tag a={g.resStart} b={g.resEnd} color={sigRes} level={1} text="TENSÃO RESIDUAL" />
          <Tag a={g.traStart} b={g.traEnd} color={sigTra} level={0} text="TRAÇÃO" />
        </g>
      ))}

      {/* colchete por posição (agrupa colapso + tensão residual + tração) */}
      {groups.map((g) => (
        <PosBracket key={`pb${g.i}`} a={g.gMin} b={g.gMax} label={`Posição ${g.posicao}`} />
      ))}

      {/* linha de centro por colapso */}
      {groups.map((g) => {
        const gx = X(cl(g.ref));
        return (
          <g key={`c${g.i}`}>
            <line x1={gx} y1={top} x2={gx} y2={bot + 22} stroke={sig}
              strokeWidth="1" strokeDasharray="10 4 2 4" />
            <circle cx={gx} cy={cy} r="4.5" fill="none" stroke={sig} strokeWidth="1.4" />
            <line x1={gx - 8} y1={cy} x2={gx + 8} y2={cy} stroke={sig} strokeWidth="1.4" />
            <line x1={gx} y1={cy - 8} x2={gx} y2={cy + 8} stroke={sig} strokeWidth="1.4" />
          </g>
        );
      })}

      {/* ---- cotas de posição (ordinal): início de cada amostra a partir do zero ---- */}
      <text x="44" y={bot + 34} fontSize="16" fontWeight="700" fontFamily="'IBM Plex Sans',sans-serif"
        fill="#5a6a82">LADO PÉ</text>
      <line x1={X(0)} y1={bot} x2={X(0)} y2={bot + 22} stroke="#5a6a82" strokeWidth="0.9" />
      <text x={X(0)} y={bot + 34} textAnchor="middle" fontSize="19"
        fontFamily="'IBM Plex Mono',monospace" fill="#5a6a82">0</text>
      {(() => {
        const starts = [];
        groups.forEach((g) => {
          starts.push({ x: g.colStart, c: sig });
          starts.push({ x: g.resStart, c: sigRes });
          starts.push({ x: g.traStart, c: sigTra });
        });
        starts.sort((a, b) => a.x - b.x);
        return starts.map((p, i) => {
          const xx = X(cl(p.x));
          const ly = bot + 34 + (i % 2) * 32;
          return (
            <g key={`o${i}`}>
              <line x1={xx} y1={bot} x2={xx} y2={ly - 15} stroke={p.c}
                strokeWidth="0.7" strokeDasharray="2 2" />
              <text x={xx} y={ly} textAnchor="middle" fontSize="18"
                fontFamily="'IBM Plex Mono',monospace" fill={p.c}>{fmt(Math.max(0, p.x))}</text>
            </g>
          );
        });
      })()}

      {/* cota de comprimento total */}
      <line x1={X(0)} y1={bot + 92} x2={X(0)} y2={bot + 138} stroke="#9aa6b8" strokeWidth="0.7" />
      <line x1={X(L)} y1={bot} x2={X(L)} y2={bot + 138} stroke="#9aa6b8" strokeWidth="0.7" />
      <Dim x1={X(0)} x2={X(L)} y={bot + 124} label={`L = ${fmt(s.length)} mm`} />

      {/* aviso de extrapolação */}
      {anyWarn && (
        <g>
          <rect x="34" y={VB_H - 64} width="450" height="30" rx="5"
            fill="#fdecea" stroke={sig} strokeWidth="1" />
          <text x="48" y={VB_H - 44} fontSize="11.5" fill={sig}
            fontFamily="'IBM Plex Mono',monospace">
            ⚠ amostras/posição fora dos limites do tubo — representação recortada
          </text>
        </g>
      )}

      {/* ---- legenda (tabela) ---- */}
      {(() => {
        const INT = "INTERNO (sem acompanhamento)";
        const Ng = groups.length || 1;
        const rowsMeta = [
          { name: "COLAPSO", color: sig, pat: "hatch", size: colLenVal, cod: "AM_COLPS", tipoOf: (g) => (g ? g.tipo : INT) || INT },
          { name: "TENSÃO RESIDUAL", color: sigRes, pat: "hatchRes", size: residualLen, cod: "AM_TENSA", tipoOf: () => INT },
          { name: "TRAÇÃO", color: sigTra, pat: "hatchTra", size: tracaoLen, cod: "AM_TRACC", tipoOf: () => INT },
        ];
        const bx = 34, by = VB_H - 214;
        const cSw = bx + 12, cAm = bx + 36, cTam = bx + 190, cCod = bx + 290;
        const gx0 = bx + 380;
        const bwMax = VB_W - 34 - bx;
        const gcw = Math.max(150, Math.floor((bwMax - (gx0 - bx) - 12) / Ng));
        const bw = (gx0 - bx) + Ng * gcw + 12;
        const tipoFs = Ng >= 4 ? 9 : Ng === 3 ? 10.5 : Ng === 2 ? 12.5 : 13.5;
        const headY = by + 58; // linha dos cabeçalhos de coluna
        return (
          <g fontFamily="'IBM Plex Mono',monospace">
            <rect x={bx} y={by} width={bw} height="150" fill="#fff" stroke={ink} strokeWidth="1.2" />
            <text x={bx + 12} y={by + 22} fontSize="16.5" fontWeight="700" fill="#5a6a82"
              fontFamily="'IBM Plex Sans',sans-serif">LEGENDA</text>

            {/* agrupador POSIÇÃO sobre as colunas de tipo-amostra */}
            <rect x={gx0} y={by + 30} width={Ng * gcw} height="17" fill="#dce6f1" />
            <text x={gx0 + (Ng * gcw) / 2} y={by + 43} textAnchor="middle" fontSize="13"
              fontWeight="700" fill="#3a4a63" fontFamily="'IBM Plex Sans',sans-serif">POSIÇÃO</text>

            {/* cabeçalhos das colunas fixas */}
            {[["AMOSTRA", cAm], ["TAMANHO", cTam], ["COD. MES", cCod]].map(([t, x]) => (
              <text key={t} x={x} y={headY} fontSize="13.5" fontWeight="700" fill="#5a6a82"
                fontFamily="'IBM Plex Sans',sans-serif">{t}</text>
            ))}

            {/* cabeçalho por posição (Pé/Meio/Ponta) + legenda "TIPO-AMOSTRA" */}
            {groups.map((g, j) => {
              const gx = gx0 + j * gcw;
              return (
                <g key={`gh${j}`}>
                  {j > 0 && (
                    <line x1={gx} y1={by + 30} x2={gx} y2={by + 142} stroke="#e3e8ef" strokeWidth="1" />
                  )}
                  <text x={gx + 8} y={headY} fontSize="13.5" fontWeight="700" fill="#3a4a63"
                    fontFamily="'IBM Plex Sans',sans-serif">{String(g.posicao).toUpperCase()}</text>
                  <text x={gx + 8} y={headY + 14} fontSize="10.5" fontWeight="600" fill="#8190a6"
                    fontFamily="'IBM Plex Sans',sans-serif">TIPO-AMOSTRA</text>
                </g>
              );
            })}

            <line x1={bx + 12} y1={by + 78} x2={bx + bw - 12} y2={by + 78} stroke="#dde3ec" strokeWidth="1" />

            {/* linhas: COLAPSO / TENSÃO RESIDUAL / TRAÇÃO */}
            {rowsMeta.map((r, i) => {
              const ry = by + 98 + i * 20;
              return (
                <g key={i}>
                  <rect x={cSw} y={ry - 12} width="15" height="15"
                    fill={`url(#${r.pat})`} stroke={r.color} strokeWidth="1" />
                  <text x={cAm} y={ry} fontSize="13" fill={r.color}>{r.name}</text>
                  <text x={cTam} y={ry} fontSize="13" fill={ink}>{`${fmt(r.size)}mm`}</text>
                  <text x={cCod} y={ry} fontSize="13" fill={ink}>{r.cod}</text>
                  {groups.map((g, j) => {
                    const tipoVal = r.tipoOf(g);
                    const isExterno = String(tipoVal).toUpperCase().includes("EXTERNO");
                    return (
                      <text key={`c${i}_${j}`} x={gx0 + j * gcw + 8} y={ry} fontSize={tipoFs}
                        fill={isExterno ? sig : ink} fontWeight={isExterno ? "700" : "400"}>
                        {tipoVal}
                      </text>
                    );
                  })}
                </g>
              );
            })}
          </g>
        );
      })()}

      <text x="34" y={VB_H - 20} fontSize="15.1" fontWeight="600" fill={ink}
        fontFamily="'IBM Plex Sans',sans-serif">
        Sucatear o tubo com o Código 2431 – Amostra de Colapso
      </text>

      {/* paginação — canto inferior direito, junto à linha de margem */}
      <text x={VB_W - 22} y={VB_H - 22} textAnchor="end" fontSize="17" fill="#9aa6b8"
        fontFamily="'IBM Plex Mono',monospace">
        PÁG. {index + 1}/{total || 1}
      </text>
    </svg>
  );
}

/* ---------------------------------------------------------------- App */
export default function App() {
  const [screen, setScreen] = useState("home"); // home | manual | import | result
  const [sketches, setSketches] = useState([]);
  const [importErr, setImportErr] = useState("");
  const fileRef = useRef(null);

  const blank = {
    od: "",
    wt: "",
    grau: "",
    length: "",
    psi: "",
    pedidoItem: "",
    ordemProducao: "",
    count: "1",
    posMode: "fixa", // fixa | aleatoria
    firstPos: "",
    ippn: "",
    // posição (Pé|Meio|Ponta) e tipo-amostra por amostra de colapso
    positions: ["Meio"],
    tipos: ["EXTERNO (com acompanhamento TPI)"],
  };
  const [form, setForm] = useState(blank);
  const [err, setErr] = useState("");

  // nº de amostras de colapso (1..20)
  const nCount = Math.min(Math.max(parseInt(form.count || "1", 10) || 1, 1), 20);

  const printAll = () => window.print();

  const toNum = (v) => {
    const n = parseFloat(String(v).replace(",", "."));
    return isNaN(n) ? null : n;
  };

  const setCount = (v) => {
    const digits = v.replace(/[^\d]/g, "");
    setForm((f) => {
      const n = Math.min(Math.max(parseInt(digits || "1", 10) || 1, 1), 20);
      const grow = (arr, fill) => {
        const a = (arr || []).slice(0, n);
        while (a.length < n) a.push(fill);
        return a;
      };
      return {
        ...f,
        count: digits,
        positions: grow(f.positions, "Meio"),
        tipos: grow(f.tipos, "EXTERNO (com acompanhamento TPI)"),
      };
    });
  };

  const submitManual = () => {
    const od = toNum(form.od),
      wt = toNum(form.wt),
      length = toNum(form.length),
      psi = toNum(form.psi);
    const n = Math.min(Math.max(parseInt(form.count || "1", 10) || 1, 1), 20);
    if ([od, wt, length, psi].some((v) => v === null)) {
      setErr("Preencha OD, WT, comprimento do tubo e pressão spec.");
      return;
    }
    if (od <= 0 || wt <= 0 || length <= 0 || psi <= 0) {
      setErr("OD, WT, comprimento e pressão devem ser maiores que zero.");
      return;
    }
    if (2 * wt >= od) {
      setErr("2×WT não pode ser maior ou igual ao OD.");
      return;
    }
    if (!form.grau.trim() || !form.pedidoItem.trim() || !form.ordemProducao.trim() || !form.ippn.trim()) {
      setErr("Grau do aço, Pedido/Item, Ordem de Produção e IPPN são obrigatórios.");
      return;
    }
    const colLen = collapseSampleLen(od, psi);
    if (colLen === null) {
      setErr("OD/pressão fora dos critérios de tamanho da amostra de colapso (item 6).");
      return;
    }
    let firstPos = 0;
    if (form.posMode === "fixa") {
      firstPos = toNum(form.firstPos);
      if (firstPos === null) {
        setErr("Informe a posição do 1º colapso (ou troque para aleatória).");
        return;
      }
    }
    setErr("");
    const collapses = buildManualCollapses({
      od, L: length, n, colLen, posMode: form.posMode, firstPos,
    });
    // posição (Pé/Meio/Ponta) e tipo-amostra por colapso
    const positions = (form.positions || []).slice(0, n);
    const tipos = (form.tipos || []).slice(0, n);
    while (positions.length < n) positions.push("Meio");
    while (tipos.length < n) tipos.push("EXTERNO (com acompanhamento TPI)");
    const collapsesMeta = collapses.map((c, k) => ({
      ...c,
      posicao: positions[k],
      tipo: tipos[k],
    }));
    const produto = `${fmt(od)} x ${fmt(wt)} - ${form.grau.trim()}`;
    const sk = {
      id: Date.now(),
      od, wt, length, psi,
      grau: form.grau.trim(),
      produto,
      pedidoItem: form.pedidoItem.trim(),
      ordemProducao: form.ordemProducao.trim(),
      posicao: positions[0],
      tipoColapso: tipos[0],
      ippn: form.ippn.trim(),
      collapses: collapsesMeta,
    };
    setSketches((p) => [...p, sk]);
    setForm(blank);
    setScreen("result");
  };

  // normaliza cabeçalho: minúsculas, sem acentos, espaços colapsados
  const norm = (s) =>
    String(s)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();

  const colNum = (nrow, pred) => {
    const k = Object.keys(nrow).find((h) => pred(h));
    return k !== undefined ? toNum(nrow[k]) : null;
  };
  const colTxt = (nrow, pred) => {
    const k = Object.keys(nrow).find((h) => pred(h));
    return k !== undefined && String(nrow[k]).trim() !== "" ? String(nrow[k]).trim() : "";
  };

  const handleFile = async (file) => {
    if (!file) return;
    setImportErr("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!rows.length) {
        setImportErr("Planilha vazia ou sem cabeçalho reconhecível.");
        return;
      }
      // 1 - lê cada linha
      const records = rows.map((row) => {
        const nrow = {};
        Object.keys(row).forEach((k) => (nrow[norm(k)] = row[k]));
        return {
          od: colNum(nrow, (h) => h === "od"),
          wt: colNum(nrow, (h) => h === "wt"),
          length: colNum(nrow, (h) => h.includes("comprimento")),
          psi: colNum(nrow, (h) => h.includes("psi")) || 0,
          grau: colTxt(nrow, (h) => h.includes("grau")),
          pedidoItem: colTxt(nrow, (h) => h.includes("pedido") && h.includes("item")),
          ordemProducao: colTxt(nrow, (h) => h.includes("ordem")),
          posNum: colNum(nrow, (h) => h.includes("high")),
          posTxtRaw: colTxt(nrow, (h) => h === "posicao"),
          tipoRaw: colTxt(nrow, (h) => h.includes("tipo-amostra") || h.includes("tipo amostra")),
          ippn: colTxt(nrow, (h) => h.includes("ippn")),
        };
      });

      // 2 - agrupa por IPPN + Ordem de Produção (mesmo tubo -> mesmo croqui)
      const groupsMap = new Map();
      const orderKeys = [];
      records.forEach((r, i) => {
        const hasKey = norm(r.ippn) !== "" && norm(r.ordemProducao) !== "";
        const key = hasKey ? `${norm(r.ippn)}||${norm(r.ordemProducao)}` : `__single_${i}`;
        if (!groupsMap.has(key)) {
          groupsMap.set(key, []);
          orderKeys.push(key);
        }
        groupsMap.get(key).push(r);
      });

      // 3 - monta um croqui por grupo
      const parsed = [];
      let skipped = 0;
      orderKeys.forEach((key, gi) => {
        const grp = groupsMap.get(key);
        const f = grp[0]; // a 1ª linha define o tubo e a posição do 1º colapso
        if (
          ![f.od, f.wt, f.length].every((v) => typeof v === "number" && v > 0) ||
          typeof f.posNum !== "number"
        ) {
          skipped += grp.length;
          return;
        }
        const colLen = collapseSampleLen(f.od, f.psi);
        if (colLen === null) {
          skipped += grp.length;
          return; // OD/pressão fora dos critérios de tamanho
        }

        // posição (Pé/Meio/Ponta) e tipo-amostra de cada colapso (uma por linha do grupo)
        const metas = grp.map((r) => {
          const pt = norm(r.posTxtRaw);
          const posicao = pt.startsWith("pe") ? "Pé" : pt.startsWith("pont") ? "Ponta" : "Meio";
          const tipo = norm(r.tipoRaw).includes("externo")
            ? "EXTERNO (com acompanhamento TPI)"
            : "INTERNO (sem acompanhamento)";
          return { posicao, tipo };
        });

        // nº de colapsos = nº de linhas do grupo; a posição High Collapse vale só
        // para o 1º colapso; os demais entram em sequência.
        const n = grp.length;
        const firstPos = roundNearestTen(f.posNum);
        const collapses = buildManualCollapses({
          od: f.od, L: f.length, n, colLen, posMode: "fixa", firstPos,
        }).map((c, k) => ({ ...c, posicao: metas[k].posicao, tipo: metas[k].tipo }));

        parsed.push({
          id: Date.now() + gi,
          od: f.od, wt: f.wt, length: f.length, psi: f.psi, grau: f.grau,
          produto: `${fmt(f.od)} x ${fmt(f.wt)}${f.grau ? " - " + f.grau : ""}`,
          pedidoItem: f.pedidoItem || "—",
          ordemProducao: f.ordemProducao || "—",
          posicao: metas[0].posicao,
          tipoColapso: metas[0].tipo,
          ippn: f.ippn || `Linha ${gi + 1}`,
          collapses,
        });
      });
      if (!parsed.length) {
        setImportErr(
          "Nenhuma linha válida. Confira as colunas: OD, WT, Comprimento do tubo MES, Pressão Spec (Psi), Posição High Collapse(mm), Posição, IPPN, Tipo-amostra."
        );
        return;
      }
      if (skipped > 0)
        setImportErr(`${parsed.length} página(s) gerada(s); ${skipped} linha(s) ignorada(s) por dados ausentes ou fora dos critérios.`);
      setSketches(parsed);
      setScreen("result");
    } catch (e) {
      setImportErr("Não foi possível ler o arquivo. Use .xlsx, .xls ou .csv.");
    }
  };

  return (
    <div className="ct-root">
      <style dangerouslySetInnerHTML={{ __html: CSS }} />
      <div className="ct-wrap">
        {/* HEADER */}
        <div className="ct-head no-print">
          <div className="ct-brand">
            <div className="ct-mark"><PencilRuler size={22} /></div>
            <div>
              <div className="ct-title">CONTROLE DA QUALIDADE</div>
              <div className="ct-sub">Croqui para corte de amostras · colapso</div>
            </div>
          </div>
          <div className="ct-cred">
            <span className="ct-cred-top">Desenvolvido por</span>
            <span className="ct-cred-name">Emerson Santos</span>
          </div>
        </div>

        {/* HOME */}
        {screen === "home" && (
          <>
            <p className="ct-lead">
              Gere o croqui de um tubo circular (vista horizontal) com a seção de amostra
              posicionada pelo seu centro a partir do ponto zero (extremidade esquerda).
              Escolha como deseja fornecer os dados.
            </p>
            <div className="ct-grid">
              <div className="ct-card" onClick={() => { setImportErr(""); setScreen("import"); }}>
                <div className="ct-ic"><FileSpreadsheet size={24} /></div>
                <h3>Importar dados</h3>
                <p>Envie uma planilha (.xlsx / .xls / .csv). Cada linha gera um croqui em sua própria página A4 paisagem.</p>
                <span className="ct-tag mono">PLANILHA →</span>
              </div>
              <div className="ct-card" onClick={() => { setErr(""); setScreen("manual"); }}>
                <div className="ct-ic"><PencilRuler size={24} /></div>
                <h3>Inserir manualmente</h3>
                <p>Informe OD, WT, comprimento, posição do colapso e o comprimento da amostra para gerar um croqui.</p>
                <span className="ct-tag mono">FORMULÁRIO →</span>
              </div>
            </div>
            {sketches.length > 0 && (
              <div className="ct-actions">
                <button className="btn" onClick={() => setScreen("result")}>
                  Ver croquis gerados ({sketches.length})
                </button>
              </div>
            )}
          </>
        )}

        {/* MANUAL */}
        {screen === "manual" && (
          <>
            <button className="back no-print" onClick={() => setScreen("home")}>
              <ArrowLeft size={16} /> Voltar
            </button>
            <div className="ct-panel">
              <div className="ct-frm">
                <Field label="OD — diâmetro externo" unit="mm" value={form.od}
                  onChange={(v) => setForm({ ...form, od: v })} />
                <Field label="WT — espessura de parede" unit="mm" value={form.wt}
                  onChange={(v) => setForm({ ...form, wt: v })} />
                <Field label="Grau do aço" text value={form.grau}
                  onChange={(v) => setForm({ ...form, grau: v })} />
                <Field label="Comprimento do tubo" unit="mm" value={form.length}
                  onChange={(v) => setForm({ ...form, length: v })} />
                <Field label="PRESSÃO DE COLAPSO ESPECIFICADA" unit="psi" value={form.psi}
                  onChange={(v) => setForm({ ...form, psi: v })} />
                <Field label="Pedido/Item" text value={form.pedidoItem}
                  onChange={(v) => setForm({ ...form, pedidoItem: v })} />
                <Field label="Ordem de Produção" text value={form.ordemProducao}
                  onChange={(v) => setForm({ ...form, ordemProducao: v })} />
                <Field label="Nº de amostras de colapso por tubo"
                  value={form.count} onChange={setCount} />
              </div>

              <div className="pos-block">
                <div className="fld">
                  <label>Posição do colapso</label>
                  <div className="seg">
                    {[["fixa", "Fixa"], ["aleatoria", "Aleatória"]].map(([k, t]) => (
                      <button key={k} type="button"
                        className={form.posMode === k ? "on" : ""}
                        onClick={() => setForm({ ...form, posMode: k })}>{t}</button>
                    ))}
                  </div>
                  {form.posMode === "aleatoria" && (
                    <span className="hint">
                      posição sorteada (arredondada à dezena); as demais seguem em sequência.
                    </span>
                  )}
                </div>
                {form.posMode === "fixa" && (
                  <Field label="Posição High Collapse (1º colapso)" unit="mm"
                    value={form.firstPos}
                    onChange={(v) => setForm({ ...form, firstPos: v })} />
                )}
              </div>

              <div className="ct-frm" style={{ marginTop: 16 }}>
                <Field label="IPPN" text value={form.ippn}
                  onChange={(v) => setForm({ ...form, ippn: v })} />
              </div>

              {/* Posição (Pé/Meio/Ponta) e Tipo-amostra por amostra de colapso */}
              <div className="pos-block" style={{ gridTemplateColumns: "1fr" }}>
                <label style={{ fontSize: 12, color: "var(--muted)", letterSpacing: ".4px",
                  textTransform: "uppercase", fontWeight: 500 }}>
                  Posição e tipo-amostra por colapso
                </label>
              </div>
              {Array.from({ length: nCount }).map((_, k) => (
                <div className="ct-frm" key={`amostra-${k}`} style={{ marginTop: 12 }}>
                  <Field
                    label={nCount > 1 ? `Posição — colapso ${k + 1}` : "Posição"}
                    options={["Pé", "Meio", "Ponta"]}
                    value={(form.positions && form.positions[k]) || "Meio"}
                    onChange={(v) =>
                      setForm((f) => {
                        const a = (f.positions || []).slice();
                        a[k] = v;
                        return { ...f, positions: a };
                      })
                    } />
                  <Field
                    label={nCount > 1 ? `Tipo-amostra — colapso ${k + 1}` : "Tipo-amostra"}
                    options={[
                      "EXTERNO (com acompanhamento TPI)",
                      "INTERNO (sem acompanhamento)",
                    ]}
                    value={(form.tipos && form.tipos[k]) || "EXTERNO (com acompanhamento TPI)"}
                    onChange={(v) =>
                      setForm((f) => {
                        const a = (f.tipos || []).slice();
                        a[k] = v;
                        return { ...f, tipos: a };
                      })
                    } />
                </div>
              ))}

              {err && (
                <div className="note">
                  <AlertTriangle size={16} /> <span>{err}</span>
                </div>
              )}
              <div className="ct-actions">
                <button className="btn primary" onClick={submitManual}>
                  <Plus size={16} /> Gerar croqui
                </button>
              </div>
            </div>
          </>
        )}

        {/* IMPORT */}
        {screen === "import" && (
          <>
            <button className="back no-print" onClick={() => setScreen("home")}>
              <ArrowLeft size={16} /> Voltar
            </button>
            <div className="ct-panel">
              <div className="drop" onClick={() => fileRef.current?.click()}>
                <div className="di"><Upload size={32} /></div>
                <h4>Selecionar planilha</h4>
                <p>.xlsx · .xls · .csv — cada linha vira um croqui (1 página A4)</p>
              </div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv"
                style={{ display: "none" }}
                onChange={(e) => handleFile(e.target.files[0])} />
              {importErr && (
                <div className="note"><AlertTriangle size={16} /> <span>{importErr}</span></div>
              )}
              <div className="note info">
                <FileSpreadsheet size={16} />
                <span>
                  Cada linha gera uma página (1 colapso por tubo, vinculado ao <b>IPPN</b>).
                  Linhas com o <b>mesmo IPPN e a mesma Ordem de Produção</b> são agrupadas
                  num único croqui (vários colapsos no mesmo tubo); nesse caso a
                  <b> Posição High Collapse</b> vale só para o 1º colapso e os demais entram em sequência.
                  Colunas esperadas: <b>OD</b>, <b>WT</b>, <b>Grau do aço</b>,
                  <b> Comprimento do tubo MES</b>, <b>Pressão Spec (Psi)</b>, <b>Pedido/Item</b>,
                  <b> Ordem de Produção</b>, <b>Posição High Collapse(mm)</b>, <b>Posição</b>,
                  <b> IPPN</b> e <b>Tipo-amostra</b>. O tamanho do colapso é definido por OD + pressão.
                </span>
              </div>
            </div>
          </>
        )}

        {/* RESULT */}
        {screen === "result" && (
          <>
            <div className="ct-bar no-print">
              <button className="back" style={{ margin: 0 }} onClick={() => setScreen("home")}>
                <ArrowLeft size={16} /> Início
              </button>
              <div className="ct-count">
                <b>{sketches.length}</b> croqui(s) · <b>{sketches.length}</b> página(s) A4 paisagem
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button className="btn ghost" onClick={() => { setSketches([]); setScreen("home"); }}>
                  <Trash2 size={16} /> Limpar
                </button>
                <button className="btn primary" onClick={printAll}>
                  <Printer size={16} /> Imprimir / PDF
                </button>
              </div>
            </div>
            {sketches.length === 0 ? (
              <div className="ct-panel">Nenhum croqui gerado ainda.</div>
            ) : (
              <div className="pages">
                {sketches.map((s, i) => (
                  <div className="page" key={s.id}>
                    <Croqui s={s} index={i} total={sketches.length} />
                  </div>
                ))}
              </div>
            )}
          </>
        )}

        {/* RODAPÉ — link do manual (não sai na impressão) */}
        <div className="ct-foot no-print">
          <a href="/manual.pdf" target="_blank" rel="noreferrer">
            <FileText size={24} /> Manual do usuário
          </a>
        </div>
      </div>
    </div>
  );
}

function Field({ label, unit, hint, value, onChange, text, options }) {
  return (
    <div className="fld">
      <label>
        {label} {hint && <span className="hint">· {hint}</span>}
      </label>
      {options ? (
        <select className="inp" value={value} onChange={(e) => onChange(e.target.value)}>
          {options.map((o) => (
            <option key={o} value={o}>{o}</option>
          ))}
        </select>
      ) : (
        <div className="ip">
          <input
            className="inp"
            inputMode={text ? "text" : "decimal"}
            value={value}
            placeholder={text ? "" : "0"}
            onChange={(e) => onChange(e.target.value)}
            style={unit ? { paddingRight: 44 } : undefined}
          />
          {unit && <span className="unit mono">{unit}</span>}
        </div>
      )}
    </div>
  );
}
