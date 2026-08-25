import React, { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTES ───────────────────────────────────────────────────────────────
const ORDEM_MES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho",
                   "Agosto","Setembro","Outubro","Novembro","Dezembro"];
const MESES_CURTOS = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];

const STATUS_LIST = ["Concluído","Pendente Cliente","Pendente GPS","N/A","Negócio Perdido"];
const STATUS_COR = {
  "Concluído":        { dot:"#16A34A", text:"#15803D", bg:"#DCFCE7" },
  "Pendente Cliente": { dot:"#D97706", text:"#B45309", bg:"#FEF9C3" },
  "Pendente GPS":     { dot:"#2563EB", text:"#1D4ED8", bg:"#DBEAFE" },
  "N/A":              { dot:"#9CA3AF", text:"#6B7280", bg:"#F3F4F6" },
  "Negócio Perdido":  { dot:"#DC2626", text:"#B91C1C", bg:"#FEE2E2" },
};
const TIPO_LIST   = ["Reajuste","DT","Renovação","Negócio Perdido"];
const TIPO_LIST_P = ["Reajuste","Renovação","Up Selling","Defesa de Território","Alteração de Escopo","BID/Cotação","Outros"];
const TIPO_COR_P  = {
  "Reajuste":"#7C3AED","Renovação":"#0891B2","Up Selling":"#059669",
  "Defesa de Território":"#D97706","Alteração de Escopo":"#0369A1",
  "BID/Cotação":"#DC2626","Outros":"#6B7280"
};

const EQUIPE = ["Mariana","Wilder","Giovanni","Carla","Darlan"];
const COR_PESSOA = {
  Mariana:"#7C3AED", Wilder:"#0369A1", Giovanni:"#059669",
  Carla:"#D97706",   Darlan:"#DC2626"
};
const COR_PESSOA_BG = {
  Mariana:"#F5F3FF", Wilder:"#EFF6FF", Giovanni:"#ECFDF5",
  Carla:"#FFFBEB",   Darlan:"#FEF2F2"
};
const SIN_COR = { Verde:"#16A34A", Amarelo:"#D97706", Vermelho:"#DC2626" };
const SIN_BG  = { Verde:"#DCFCE7", Amarelo:"#FEF9C3", Vermelho:"#FEE2E2" };

const DARK  = "#1a2332";
const DARK2 = "#243044";
const MENU_BG = "#0F1F35";

// ─── FORMATAÇÃO ───────────────────────────────────────────────────────────────
const brl = v => (!v && v !== 0) ? "—"
  : new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL",maximumFractionDigits:0}).format(v);
const pct = v => (v != null && v !== "" && !isNaN(v) && parseFloat(v) !== 0)
  ? `${(parseFloat(v)*100).toFixed(2)}%` : "—";
const num = v => { const n = parseFloat(String(v).replace(/[R$\s.]/g,"").replace(",",".")); return isNaN(n) ? 0 : n; };

// ─── FETCH ────────────────────────────────────────────────────────────────────
const fetchCSV = async (gid) => {
  const res = await fetch(`/api/sheet?gid=${gid}`);
  if (!res.ok) throw new Error(`Erro HTTP ${res.status}`);
  const text = await res.text();
  if (!text || text.length < 10) throw new Error("Planilha vazia");

  const parseCSVLine = line => {
    const result = []; let cur = "", inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"' && line[i+1] === '"') { cur += '"'; i++; }
      else if (ch === '"') { inQ = !inQ; }
      else if (ch === ',' && !inQ) { result.push(cur.trim()); cur = ""; }
      else { cur += ch; }
    }
    result.push(cur.trim());
    return result;
  };

  const lines = text.split("\n").filter(l => l.trim());
  if (!lines.length) throw new Error("Planilha vazia");
  const headers = parseCSVLine(lines[0]);
  return lines.slice(1).map(line => {
    const vals = parseCSVLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (vals[i] ?? "").replace(/^"|"$/g, "").trim(); });
    return obj;
  }).filter(r => Object.values(r).some(v => v));
};

// ─── PROCESSA GERENCIAL ───────────────────────────────────────────────────────
// Ordem dos meses para calcular se mês vigência é futuro/atual/passado
const MESES_MAP = {
  "janeiro":1,"fevereiro":2,"março":3,"abril":4,"maio":5,"junho":6,
  "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12,
  "jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,
  "jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12,
};
const parseMes = s => {
  if (!s) return null;
  const clean = s.toString().toLowerCase().replace(/\.$/,"").trim();
  return MESES_MAP[clean] || null;
};
const parseData = s => {
  if (!s) return null;
  const str = s.toString().trim();
  if (!str || str==="-" || str==="NaT") return null;
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
};
const diasEntre = (d1, d2) => {
  if (!d1 || !d2) return null;
  return Math.round(Math.abs(d2 - d1) / (1000*60*60*24));
};
const fmtData = d => {
  if (!d) return "—";
  return d.toLocaleDateString("pt-BR", {day:"2-digit",month:"2-digit",year:"numeric"});
};
const pctFmt = v => {
  if (!v || v === "" || v === "0" || parseFloat(v) === 0) return "—";
  const n = parseFloat(v);
  if (isNaN(n)) return "—";
  // Valores já em decimal (0.05) ou em percentual (5.29)
  const val = n > 1 ? n : n * 100;
  return `${val.toFixed(2)}%`;
};

const calcStatusTemporal = (statusReal, mesVig, inicioNeg, dataAprov) => {
  const agora = new Date();
  const mesAtual = agora.getMonth() + 1;
  const anoAtual = agora.getFullYear();
  const mesNum = parseMes(mesVig);

  // Tem data de aprovação → sempre Concluído
  if (dataAprov) return { status: statusReal || "Concluído", concluido: true };

  // Sem aprovação
  if (inicioNeg) {
    // Tem início de negociação
    if (mesNum && mesNum > mesAtual) {
      return { status: "Em Negociação Antecipada", concluido: false };
    }
    return { status: "Em Negociação", concluido: false };
  }

  // Sem início
  if (mesNum && mesNum > mesAtual) {
    return { status: "Aguardando Reajuste", concluido: false };
  }
  return { status: statusReal || "N/A", concluido: false };
};

const processGerencial = raw => {
  const keys = raw.length ? Object.keys(raw[0]) : [];
  const find = (...kws) => keys.find(k => kws.some(kw => k.toUpperCase().includes(kw.toUpperCase()))) || null;
  const get  = (r, ...kws) => { const k = find(...kws); return k ? (r[k] ?? "") : ""; };
  return raw.map((r, i) => {
    const grupo  = get(r,"Grupo Cliente").toString().trim();
    const cr     = get(r,"CR","CONTRATO").toString().trim();
    if (!grupo || !cr) return null;

    const statusReal = get(r,"Status Real","Status").toString().trim() || "N/A";
    const mesVig     = get(r,"Mês Vigência","MES").toString().trim();
    const inicioNeg  = parseData(get(r,"Inicio Negociação","Início Negociação","INICIO NEG"));
    const dataAprov  = parseData(get(r,"Data de Aprovação","DATA APROVAÇÃO","DATA APROV"));
    const devido     = get(r,"Devido","DEVIDO").toString().trim();
    const aplicado   = get(r,"Aplicado","APLICADO").toString().trim();

    const { status } = calcStatusTemporal(statusReal, mesVig, inicioNeg, dataAprov);
    const dias = diasEntre(inicioNeg, dataAprov || (dataAprov ? null : (inicioNeg ? new Date() : null)));

    return {
      _id: i, grupo, cr,
      descr:     get(r,"DESCRI").toString().trim(),
      tipo:      get(r,"Tipo de Negócio","Tipo").toString().trim() || "Reajuste",
      status,
      statusReal,
      mes:       mesVig,
      farmer:    get(r,"Responsável Farmer","FARMER").toString().trim(),
      info:      get(r,"Informações").toString().trim(),
      devido,    aplicado,
      inicioNeg, dataAprov,
      diasNeg:   dias,
      _edited: false,
    };
  }).filter(Boolean);
};

// ─── PROCESSA PRODUTIVIDADE ───────────────────────────────────────────────────
const matchPessoa = s => {
  const u = (s||"").toUpperCase();
  if (u.includes("MARIANA")) return "Mariana";
  if (u.includes("WILDER"))  return "Wilder";
  if (u.includes("GIOVANNI"))return "Giovanni";
  if (u.includes("CARLA"))   return "Carla";
  if (u.includes("DARLAN"))  return "Darlan";
  return null;
};
const matchSin = v => {
  const s = (v||"").toUpperCase();
  if (s.includes("VERDE"))   return "Verde";
  if (s.includes("AMARELO")) return "Amarelo";
  if (s.includes("VERM"))    return "Vermelho";
  return null;
};
const matchCatP = (ativ, tipo) => {
  const c = ((ativ||"")+" "+(tipo||"")).toUpperCase();
  if (c.includes("REAJUSTE")||c.includes("NOTIFICAÇ")||c.includes("CARTA DE REAJUSTE")) return "Reajuste";
  if (c.includes("RENOVAÇ")) return "Renovação";
  if (c.includes("UP-SELLING")||c.includes("UP SELLING")||c.includes("UPSELLING")) return "Up Selling";
  if (c.includes("DEFESA")) return "Defesa de Território";
  if (c.includes("ALTERAÇ")&&c.includes("ESCOPO")) return "Alteração de Escopo";
  if (c.includes("BID")||c.includes("COTAÇ")) return "BID/Cotação";
  return "Outros";
};

const processProdutividade = raw => {
  const keys = raw.length ? Object.keys(raw[0]) : [];
  const find = (...kws) => keys.find(k => kws.some(kw => k.toUpperCase().includes(kw.toUpperCase()))) || null;
  const get  = (r, ...kws) => { const k = find(...kws); return k ? (r[k] ?? "") : ""; };
  return raw.map(r => {
    const respF = get(r,"Responsável Farmer","RESP. FARMER").toString().trim();
    const respH = get(r,"Responsável Hunter","RESP. HUNTER").toString().trim();
    const pessoa = matchPessoa(respF) || matchPessoa(respH);
    if (!pessoa) return null;
    const ativ   = get(r,"ATIVIDADE","ATIVIDADE1").toString().toUpperCase().trim();
    const tipo   = get(r,"Tipo de Negócio","TIPO DE PROPOSTA").toString().trim();
    const status = get(r,"Status","STATUS").toString().trim();
    const valAtual  = num(get(r,"Valor Contrato Atual"));
    const valPleito = num(get(r,"Valor Final / Pleito","COM REAJUSTE","VALOR PROPOSTA"));
    const grupo  = get(r,"Grupo Cliente","GRUPO CLIENTE").toString().trim();
    const nProp  = get(r,"Nº Proposta","N PROPOSTA").toString().trim();
    const mes    = get(r,"Mês","MES").toString().trim();
    const obs    = get(r,"OBS.","OBS").toString().trim();
    return {
      pessoa, grupo, nProp, mes, obs, ativ, tipo, status,
      valAtual, valPleito, diferenca: valPleito - valAtual,
      pctPleito: num(get(r,"% Reajuste Pleito","PLEITO (%)")),
      pctAceito: num(get(r,"% Reajuste Aceito","APROVADO PELO CLIENTE (%)")),
      valAprov:  num(get(r,"APROVADO PELO CLIENTE (R$)")),
      isRevisao: (get(r,"Nº da Revisão","Revisão")||"0").toString() !== "0",
      sinalizacao: matchSin(get(r,"Semáforo","Sinalização")),
      categoria: matchCatP(ativ, tipo),
      escopo: get(r,"Escopo Atuação","ESCOPO").toString().trim(),
    };
  }).filter(Boolean);
};

// ─── EXPORTAR EXCEL ───────────────────────────────────────────────────────────
const exportExcel = data => {
  const rows = data.map(r => ({
    "Grupo Cliente": r.grupo, "CR": r.cr, "Descrição CR": r.descr,
    "Tipo de Negócio": r.tipo, "Mês Vigência": r.mes,
    "Responsável Farmer": r.farmer, "Status Real": r.status,
    "Informações": r.info, "Editado": r._edited ? "Sim" : "",
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Gerencial");
  XLSX.writeFile(wb, "Gerencial_Atualizado.xlsx");
};

// ─── INLINE SELECT ────────────────────────────────────────────────────────────
const InlineSelect = ({ value, options, colorMap, onChange }) => {
  const [open, setOpen] = useState(false);
  const [pos,  setPos]  = useState({ top:0, left:0 });
  const triggerRef = useRef(null);
  const c = colorMap?.[value] || { dot:"#9CA3AF", text:"#6B7280", bg:"#F3F4F6" };

  const handleOpen = e => {
    e.stopPropagation();
    if (!open && triggerRef.current) {
      const r = triggerRef.current.getBoundingClientRect();
      const spaceBelow = window.innerHeight - r.bottom;
      const menuH = options.length * 42;
      const openUp = spaceBelow < menuH + 8;
      setPos({ top: openUp ? r.top - menuH - 4 : r.bottom + 4, left: r.left });
    }
    setOpen(o => !o);
  };

  return (
    <div style={{ position:"relative", display:"inline-block" }}>
      <div ref={triggerRef} onClick={handleOpen}
           style={{ display:"inline-flex", alignItems:"center", gap:6, cursor:"pointer",
                    padding:"3px 10px", borderRadius:99, background:c.bg,
                    border:`1.5px solid ${open ? c.dot : "transparent"}`,
                    transition:"border .15s", userSelect:"none" }}>
        {colorMap && <div style={{ width:7, height:7, borderRadius:99, background:c.dot, flexShrink:0 }}/>}
        <span style={{ fontSize:12, fontWeight:600, color:c.text, whiteSpace:"nowrap" }}>{value}</span>
        <span style={{ fontSize:9, color:c.text, opacity:.6 }}>▼</span>
      </div>
      {open && (
        <>
          <div onClick={()=>setOpen(false)} style={{ position:"fixed", inset:0, zIndex:500 }}/>
          <div style={{ position:"fixed", top:pos.top, left:pos.left, zIndex:501,
                        background:"#fff", borderRadius:10, boxShadow:"0 8px 24px rgba(0,0,0,.18)",
                        border:"1px solid #E2E8F0", minWidth:180, overflow:"hidden" }}>
            {options.map(opt => {
              const oc = colorMap?.[opt] || { dot:"#9CA3AF", text:"#374151" };
              return (
                <div key={opt} onClick={e=>{ e.stopPropagation(); onChange(opt); setOpen(false); }}
                     style={{ display:"flex", alignItems:"center", gap:8, padding:"9px 14px",
                              cursor:"pointer", background: opt===value ? "#F8FAFC" : "#fff" }}
                     onMouseOver={e=>e.currentTarget.style.background="#F1F5F9"}
                     onMouseOut={e=>e.currentTarget.style.background=opt===value?"#F8FAFC":"#fff"}>
                  {colorMap && <div style={{ width:8, height:8, borderRadius:99, background:oc.dot }}/>}
                  <span style={{ fontSize:13, color:oc.text, fontWeight:opt===value?600:400 }}>{opt}</span>
                  {opt===value && <span style={{ marginLeft:"auto", fontSize:11, color:oc.dot }}>✓</span>}
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
};

// ─── EXPORT BTN ───────────────────────────────────────────────────────────────
const ExportBtn = ({ nEditados, onExport }) => {
  const [expanded, setExpanded] = useState(false);
  return (
    <div style={{ position:"fixed", bottom:24, right:24, zIndex:300,
                  display:"flex", flexDirection:"column", alignItems:"flex-end", gap:8 }}>
      {expanded && (
        <div style={{ background:"#fff", borderRadius:12, padding:"14px 18px",
                      boxShadow:"0 8px 24px rgba(0,0,0,.18)", border:"1px solid #E2E8F0",
                      display:"flex", flexDirection:"column", gap:10, minWidth:200 }}>
          <div style={{ fontSize:12, color:"#64748B" }}>
            <span style={{ fontWeight:700, color:"#F59E0B" }}>{nEditados}</span> alteraç{nEditados===1?"ão":"ões"}
          </div>
          <button onClick={onExport} style={{ padding:"9px 16px", borderRadius:8, background:"#16A34A",
            border:"none", color:"#fff", fontSize:13, fontWeight:700, cursor:"pointer" }}>
            ⬇ Exportar Excel
          </button>
          <button onClick={()=>setExpanded(false)} style={{ padding:"6px", borderRadius:8,
            background:"#F1F5F9", border:"none", color:"#64748B", fontSize:12, cursor:"pointer" }}>
            Minimizar
          </button>
        </div>
      )}
      <button onClick={()=>setExpanded(o=>!o)}
        style={{ width:48, height:48, borderRadius:99, background:expanded?"#64748B":"#16A34A",
                 border:"none", color:"#fff", fontSize:20, cursor:"pointer",
                 boxShadow:"0 4px 16px rgba(0,0,0,.25)", display:"flex",
                 alignItems:"center", justifyContent:"center", position:"relative" }}>
        {expanded ? "✕" : "⬇"}
        {!expanded && (
          <span style={{ position:"absolute", top:-4, right:-4, background:"#F59E0B", color:"#fff",
                         fontSize:10, fontWeight:700, minWidth:18, height:18, borderRadius:99,
                         display:"flex", alignItems:"center", justifyContent:"center", padding:"0 4px",
                         border:"2px solid #fff" }}>{nEditados}</span>
        )}
      </button>
    </div>
  );
};

// ─── PILL ─────────────────────────────────────────────────────────────────────
const Pill = ({ label, active, onClick, color }) => (
  <button onClick={onClick} style={{
    padding:"4px 12px", borderRadius:99, border:"none", cursor:"pointer",
    fontFamily:"inherit", fontSize:11, fontWeight:500, transition:"all .15s",
    background: active ? (color||DARK) : "#F1F4F8",
    color:      active ? "#fff" : "#64748B",
  }}>{label}</button>
);

// ─── SEÇÃO GERENCIAL ──────────────────────────────────────────────────────────
function SecaoGerencial({ data, setData }) {
  const [fMeses,  setFMeses]  = useState(new Set());
  const [fFarmer, setFFarmer] = useState("Todos");
  const [fStatus, setFStatus] = useState(new Set());
  const [fTipo,   setFTipo]   = useState("Todos");
  const [busca,   setBusca]   = useState("");
  const [tooltip, setTooltip] = useState({ visible:false, text:"", x:0, y:0 });

  const editRow = (id, field, value) =>
    setData(prev => prev.map(r => r._id===id ? {...r,[field]:value,_edited:true} : r));

  const mesesDisp  = ORDEM_MES.filter(m => data.some(r => r.mes===m));
  const todosMeses = fMeses.size===0;
  const toggleMes  = m => setFMeses(prev => { const n=new Set(prev); n.has(m)?n.delete(m):n.add(m); return n; });
  const todosStat  = fStatus.size===0;
  const farmersDisp = ["Todos",...Array.from(new Set(data.map(r=>r.farmer).filter(Boolean))).sort()];

  const filtrado = data.filter(r => {
    if (!todosMeses && !fMeses.has(r.mes)) return false;
    if (fFarmer !== "Todos" && r.farmer !== fFarmer) return false;
    if (!todosStat && !fStatus.has(r.status)) return false;
    if (fTipo !== "Todos" && r.tipo !== fTipo) return false;
    if (busca && ![r.grupo,r.cr,r.descr].some(v=>v.toUpperCase().includes(busca.toUpperCase()))) return false;
    return true;
  });

  const cardBase = todosMeses ? data : data.filter(r=>fMeses.has(r.mes));
  const buildCard = rows => ({
    total: rows.length,
    por: STATUS_LIST.map(s=>({ s, n:rows.filter(r=>r.status===s).length }))
  });
  const cardGeral = buildCard(cardBase);
  const cardsTipo = TIPO_LIST.map(t=>({ tipo:t, ...buildCard(cardBase.filter(r=>r.tipo===t)) })).filter(c=>c.total>0);

  const grupos = (() => {
    const by = {};
    filtrado.forEach(r => { if(!by[r.grupo]) by[r.grupo]=[]; by[r.grupo].push(r); });
    return Object.entries(by).map(([g,crs])=>({g,crs})).sort((a,b)=>a.g.localeCompare(b.g));
  })();

  const nEditados = data.filter(r=>r._edited).length;
  const showTip = (e,t) => { if(t) setTooltip({visible:true,text:t,x:e.clientX,y:e.clientY}); };
  const moveTip = e => { if(tooltip.visible) setTooltip(p=>({...p,x:e.clientX,y:e.clientY})); };
  const hideTip = () => setTooltip(p=>({...p,visible:false}));

  return (
    <div>
      {tooltip.visible && tooltip.text && (
        <div style={{ position:"fixed", zIndex:9999, pointerEvents:"none",
          left:tooltip.x+14, top:tooltip.y-10, background:DARK, color:"#F8FAFC",
          padding:"8px 12px", borderRadius:8, fontSize:12, maxWidth:320, lineHeight:1.5,
          boxShadow:"0 4px 16px rgba(0,0,0,.3)", whiteSpace:"pre-wrap" }}>
          {tooltip.text}
        </div>
      )}

      {/* FILTROS */}
      <div style={{ background:"#fff", borderRadius:12, padding:"12px 18px", marginBottom:20,
                    display:"flex", gap:12, flexWrap:"wrap", alignItems:"center",
                    boxShadow:"0 1px 3px rgba(0,0,0,.06)" }}>
        <input value={busca} onChange={e=>setBusca(e.target.value)}
          placeholder="Buscar cliente, CR ou descrição..."
          style={{ padding:"6px 12px", borderRadius:8, border:"1px solid #E2E8F0",
                   fontSize:12, color:DARK, outline:"none", width:220, fontFamily:"inherit" }}/>
        <div style={{ width:1, height:22, background:"#E2E8F0" }}/>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Mês</span>
          <Pill label="Todos" active={todosMeses} onClick={()=>setFMeses(new Set())}/>
          {mesesDisp.map(m=><Pill key={m} label={m} active={fMeses.has(m)} onClick={()=>toggleMes(m)}/>)}
        </div>
        <div style={{ width:1, height:22, background:"#E2E8F0" }}/>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Farmer</span>
          {farmersDisp.map(f=><Pill key={f} label={f} active={fFarmer===f} onClick={()=>setFFarmer(f)}/>)}
        </div>
        <div style={{ width:1, height:22, background:"#E2E8F0" }}/>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Status</span>
          <Pill label="Todos" active={todosStat} onClick={()=>setFStatus(new Set())}/>
          {STATUS_LIST.map(s=><Pill key={s} label={s} active={fStatus.has(s)}
            onClick={()=>setFStatus(prev=>{ const n=new Set(prev); n.has(s)?n.delete(s):n.add(s); return n; })}/>)}
        </div>
        <div style={{ width:1, height:22, background:"#E2E8F0" }}/>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Tipo</span>
          {["Todos",...TIPO_LIST].map(t=><Pill key={t} label={t} active={fTipo===t} onClick={()=>setFTipo(t)}/>)}
        </div>
        <span style={{ marginLeft:"auto", fontSize:12, color:"#94A3B8" }}>
          {filtrado.length} CRs · {grupos.length} grupos
        </span>
      </div>

      {/* CARDS */}
      <div style={{ display:"grid", gridTemplateColumns:`repeat(${1+cardsTipo.length},1fr)`,
                    gap:14, marginBottom:22 }}>
        {/* Card Geral */}
        <div style={{ background:DARK, borderRadius:14, padding:"20px 22px",
                      boxShadow:"0 4px 14px rgba(26,35,50,.25)" }}>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:14 }}>
            <div>
              <div style={{ fontSize:11, color:"rgba(255,255,255,.5)", marginBottom:4 }}>Total Geral</div>
              <div style={{ fontSize:36, fontWeight:800, color:"#fff", lineHeight:1 }}>{cardGeral.total}</div>
              <div style={{ fontSize:11, color:"rgba(255,255,255,.4)", marginTop:4 }}>
                {cardGeral.total>0 ? Math.round((cardGeral.por.find(p=>p.s==="Concluído")?.n||0)/cardGeral.total*100) : 0}% concluído
              </div>
            </div>
            <div style={{ fontSize:28, opacity:.3 }}>📋</div>
          </div>
          {cardGeral.por.map(({s,n}) => {
            const c = STATUS_COR[s];
            const p = cardGeral.total>0 ? Math.round(n/cardGeral.total*100) : 0;
            return (
              <div key={s} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:7 }}>
                <div style={{ width:8, height:8, borderRadius:99, background:c.dot, flexShrink:0 }}/>
                <span style={{ fontSize:12, color:"rgba(255,255,255,.6)", flex:1 }}>{s}</span>
                <span style={{ fontSize:13, fontWeight:700, color:"#fff", minWidth:22, textAlign:"right" }}>{n}</span>
                <span style={{ fontSize:11, color:"rgba(255,255,255,.35)", minWidth:36, textAlign:"right" }}>({p}%)</span>
              </div>
            );
          })}
        </div>
        {/* Cards por Tipo */}
        {cardsTipo.map(({tipo,total,por}) => {
          const conc = por.find(p=>p.s==="Concluído")?.n||0;
          return (
            <div key={tipo} style={{ background:"#fff", borderRadius:14, padding:"20px 22px",
                                     borderLeft:`4px solid ${DARK}`, boxShadow:"0 1px 4px rgba(0,0,0,.07)" }}>
              <div style={{ display:"flex", justifyContent:"space-between", marginBottom:14 }}>
                <div>
                  <div style={{ fontSize:11, color:"#94A3B8", marginBottom:4 }}>{tipo}</div>
                  <div style={{ fontSize:36, fontWeight:800, color:DARK, lineHeight:1 }}>{total}</div>
                  <div style={{ fontSize:11, color:"#94A3B8", marginTop:4 }}>
                    {total>0?Math.round(conc/total*100):0}% concluído
                  </div>
                </div>
                <div style={{ background:"#F1F4F8", borderRadius:8, padding:"6px 10px",
                               fontSize:12, fontWeight:700, color:DARK }}>{tipo}</div>
              </div>
              {por.map(({s,n}) => {
                const c = STATUS_COR[s];
                const p = total>0 ? Math.round(n/total*100) : 0;
                return (
                  <div key={s} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:7 }}>
                    <div style={{ width:8, height:8, borderRadius:99, background:c.dot, flexShrink:0 }}/>
                    <span style={{ fontSize:12, color:"#64748B", flex:1 }}>{s}</span>
                    <span style={{ fontSize:13, fontWeight:700, color:DARK, minWidth:22, textAlign:"right" }}>{n}</span>
                    <span style={{ fontSize:11, color:"#94A3B8", minWidth:36, textAlign:"right" }}>({p}%)</span>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>

      {/* GRUPOS */}
      <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
        {grupos.length===0 && (
          <div style={{ background:"#fff", borderRadius:14, padding:"48px 0",
                        textAlign:"center", color:"#94A3B8", fontSize:14 }}>
            Nenhum CR encontrado.
          </div>
        )}
        {grupos.map(({g,crs}) => {
          const nEdit = crs.filter(r=>r._edited).length;
          return (
            <div key={g} style={{ background:"#fff", borderRadius:14, overflow:"hidden",
                                   boxShadow:"0 1px 4px rgba(0,0,0,.07)" }}>
              <div style={{ background:DARK2, padding:"12px 18px", display:"flex",
                             alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:10 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <span style={{ fontSize:14, fontWeight:700, color:"#fff" }}>{g}</span>
                  <span style={{ fontSize:11, color:"rgba(255,255,255,.4)" }}>{crs.length} CR{crs.length!==1?"s":""}</span>
                  {nEdit>0 && (
                    <span style={{ fontSize:10, background:"#F59E0B", color:"#fff",
                                   padding:"1px 7px", borderRadius:99, fontWeight:600 }}>
                      {nEdit} editado{nEdit!==1?"s":""}
                    </span>
                  )}
                </div>
                <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
                  {STATUS_LIST.map(s => {
                    const n = crs.filter(c=>c.status===s).length;
                    if (!n) return null;
                    const sc = STATUS_COR[s];
                    return (
                      <span key={s} style={{ fontSize:11, fontWeight:600, padding:"2px 10px",
                                             borderRadius:99, background:sc.bg, color:sc.text }}>
                        {n} {s}
                      </span>
                    );
                  })}
                </div>
              </div>
              <table style={{ width:"100%", borderCollapse:"collapse" }}>
                <thead>
                  <tr style={{ background:"#F8FAFC" }}>
                    {["CR","Descrição","Mês","Farmer","Devido","Aplicado","Tempo Neg.","Início","Aprovação","Tipo","Status"].map((h,i) => (
                      <th key={h} style={{ textAlign:i>2?"center":"left", padding:"8px 16px",
                                          fontSize:11, color:"#94A3B8", fontWeight:600,
                                          borderBottom:"1px solid #E2E8F0", whiteSpace:"nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {crs.map((cr,ci) => (
                    <tr key={cr._id}
                        style={{ background:cr._edited?"#FFFBEB":ci%2===0?"#fff":"#FAFBFC",
                                 borderTop:"1px solid #F1F4F8" }}
                        onMouseOver={e=>{ if(!cr._edited) e.currentTarget.style.background="#EFF6FF"; }}
                        onMouseOut={e=>{ e.currentTarget.style.background=cr._edited?"#FFFBEB":ci%2===0?"#fff":"#FAFBFC"; }}>
                      <td style={{ padding:"9px 16px", fontSize:12, color:"#64748B", fontFamily:"monospace" }}>
                        {cr._edited && <span style={{ color:"#F59E0B", marginRight:4 }}>●</span>}
                        {cr.cr}
                      </td>
                      <td style={{ padding:"9px 16px", fontSize:12, color:DARK, maxWidth:340,
                                   cursor:cr.info?"help":"default" }}
                          onMouseEnter={e=>showTip(e,cr.info)}
                          onMouseMove={moveTip} onMouseLeave={hideTip}>
                        <span style={{ display:"block", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                          {cr.descr}
                        </span>
                      </td>
                      <td style={{ padding:"9px 16px", fontSize:12, color:"#64748B", textAlign:"center", whiteSpace:"nowrap" }}>{cr.mes}</td>
                      <td style={{ padding:"9px 16px", fontSize:12, color:"#64748B", textAlign:"center", whiteSpace:"nowrap" }}>{cr.farmer||"—"}</td>
                      <td style={{ padding:"9px 16px", fontSize:12, textAlign:"center", whiteSpace:"nowrap",
                                   color:cr.devido&&cr.devido!=="—"?"#7C3AED":"#94A3B8", fontWeight:cr.devido&&cr.devido!=="—"?600:400 }}>
                        {pctFmt(cr.devido)}
                      </td>
                      <td style={{ padding:"9px 16px", fontSize:12, textAlign:"center", whiteSpace:"nowrap",
                                   color:cr.aplicado&&cr.aplicado!=="—"?"#059669":"#94A3B8", fontWeight:cr.aplicado&&cr.aplicado!=="—"?600:400 }}>
                        {pctFmt(cr.aplicado)}
                      </td>
                      <td style={{ padding:"9px 16px", fontSize:12, textAlign:"center", whiteSpace:"nowrap" }}>
                        {cr.diasNeg != null ? (
                          <span style={{ fontWeight:600,
                            color: cr.dataAprov ? "#16A34A" : "#D97706",
                            background: cr.dataAprov ? "#DCFCE7" : "#FEF9C3",
                            padding:"2px 8px", borderRadius:99, fontSize:11 }}>
                            {cr.diasNeg}d {cr.dataAprov ? "✓" : "⏳"}
                          </span>
                        ) : "—"}
                      </td>
                      <td style={{ padding:"9px 16px", fontSize:11, color:"#64748B", textAlign:"center", whiteSpace:"nowrap" }}>
                        {cr.inicioNeg ? fmtData(cr.inicioNeg) : "—"}
                      </td>
                      <td style={{ padding:"9px 16px", fontSize:11, textAlign:"center", whiteSpace:"nowrap",
                                   color:cr.dataAprov?"#16A34A":"#94A3B8" }}>
                        {cr.dataAprov ? fmtData(cr.dataAprov) : "—"}
                      </td>
                      <td style={{ padding:"6px 16px", textAlign:"center" }}>
                        <InlineSelect value={cr.tipo} options={TIPO_LIST} onChange={v=>editRow(cr._id,"tipo",v)}/>
                      </td>
                      <td style={{ padding:"6px 16px", textAlign:"center" }}>
                        <InlineSelect value={cr.status} options={STATUS_LIST} colorMap={STATUS_COR} onChange={v=>editRow(cr._id,"status",v)}/>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          );
        })}
      </div>
      {nEditados>0 && <ExportBtn nEditados={nEditados} onExport={()=>exportExcel(data)}/>}
    </div>
  );
}

// ─── SEÇÃO PRODUTIVIDADE — VISÃO GERAL ────────────────────────────────────────
function ProdVisaoGeral({ data }) {
  const [fMes, setFMes] = useState("Todos");
  const mesesDisp = ["Todos",...ORDEM_MES.filter(m=>data.some(r=>r.mes.toLowerCase().startsWith(m.substring(0,3).toLowerCase())))];

  const base = fMes==="Todos" ? data : data.filter(r=>r.mes.toLowerCase().startsWith(fMes.substring(0,3).toLowerCase()));
  const somaVal = arr => arr.reduce((s,r)=>s+r.valPleito,0);
  const somaAprov = arr => arr.reduce((s,r)=>s+r.valAprov,0);
  const nAprov = arr => arr.filter(r=>r.status.toUpperCase().includes("APROVADO")).length;

  const totalAprov = somaAprov(base);
  const totalPleito = somaVal(base);
  const txAprov = base.length>0 ? Math.round(nAprov(base)/base.filter(r=>!r.isRevisao).length*100) : 0;

  return (
    <div>
      {/* Filtro mês */}
      <div style={{ background:"#fff", borderRadius:12, padding:"10px 16px", marginBottom:20,
                    display:"flex", gap:6, alignItems:"center", flexWrap:"wrap",
                    boxShadow:"0 1px 3px rgba(0,0,0,.06)" }}>
        <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8", marginRight:4 }}>Mês</span>
        {mesesDisp.map(m => <Pill key={m} label={m} active={fMes===m} onClick={()=>setFMes(m)}/>)}
      </div>

      {/* KPIs gerais */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:12, marginBottom:24 }}>
        {[
          { l:"Propostas", v:base.filter(r=>!r.isRevisao).length, icon:"📋" },
          { l:"Valor em negociação", v:brl(totalPleito), icon:"🎯" },
          { l:"Valor aprovado", v:brl(totalAprov), icon:"✅", hl:"#16A34A" },
          { l:"Taxa de aprovação", v:`${txAprov}%`, icon:"📈", hl: txAprov>50?"#16A34A":"#D97706" },
        ].map(k => (
          <div key={k.l} style={{ background:"#fff", borderRadius:12, padding:"16px 18px",
                                   boxShadow:"0 1px 4px rgba(0,0,0,.07)" }}>
            <div style={{ fontSize:20, marginBottom:8 }}>{k.icon}</div>
            <div style={{ fontSize:11, color:"#94A3B8", marginBottom:4 }}>{k.l}</div>
            <div style={{ fontSize:20, fontWeight:700, color:k.hl||DARK }}>{k.v}</div>
          </div>
        ))}
      </div>

      {/* Cards por pessoa */}
      <div style={{ fontSize:11, fontWeight:600, letterSpacing:".08em", textTransform:"uppercase",
                    color:"#94A3B8", marginBottom:14 }}>Time</div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:16, marginBottom:24 }}>
        {EQUIPE.filter(p=>base.some(r=>r.pessoa===p)).map(p => {
          const rows = base.filter(r=>r.pessoa===p);
          const proprias = rows.filter(r=>!r.isRevisao);
          const aprovados = rows.filter(r=>r.status.toUpperCase().includes("APROVADO"));
          const valAprovP = somaAprov(aprovados);
          const valPleitoP = somaVal(rows);
          const porCat = TIPO_LIST_P.map(cat=>({ cat, n:rows.filter(r=>r.categoria===cat).length })).filter(x=>x.n>0);
          return (
            <div key={p} style={{ background:"#fff", borderRadius:14, padding:20,
                                   borderTop:`3px solid ${COR_PESSOA[p]}`,
                                   boxShadow:"0 1px 4px rgba(0,0,0,.07)" }}>
              <div style={{ display:"flex", justifyContent:"space-between", marginBottom:16 }}>
                <div>
                  <div style={{ fontSize:11, color:"#94A3B8", marginBottom:3 }}>Analista</div>
                  <div style={{ fontSize:18, fontWeight:700, color:COR_PESSOA[p] }}>{p}</div>
                  <div style={{ fontSize:11, color:"#94A3B8", marginTop:2 }}>
                    {proprias.length} proposta{proprias.length!==1?"s":""} · {rows.filter(r=>r.isRevisao).length} revisõe{rows.filter(r=>r.isRevisao).length!==1?"s":""}
                  </div>
                </div>
                <div style={{ background:"#DCFCE7", borderRadius:8, padding:"4px 10px", textAlign:"center" }}>
                  <div style={{ fontSize:18, fontWeight:700, color:"#16A34A" }}>{aprovados.length}</div>
                  <div style={{ fontSize:10, color:"#16A34A" }}>aprovados</div>
                </div>
              </div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:14 }}>
                {[
                  { l:"Valor pleito", v:brl(valPleitoP) },
                  { l:"Valor aprovado", v:brl(valAprovP), hl:true },
                ].map(k => (
                  <div key={k.l} style={{ background:"#F8FAFC", borderRadius:8, padding:"8px 10px" }}>
                    <div style={{ fontSize:10, color:"#94A3B8", marginBottom:2 }}>{k.l}</div>
                    <div style={{ fontSize:12, fontWeight:600, color:k.hl?"#16A34A":DARK }}>{k.v}</div>
                  </div>
                ))}
              </div>
              <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                {porCat.map(({cat,n}) => (
                  <span key={cat} style={{ fontSize:10, fontWeight:600, padding:"2px 8px", borderRadius:99,
                                           background:`${TIPO_COR_P[cat]}18`, color:TIPO_COR_P[cat],
                                           border:`1px solid ${TIPO_COR_P[cat]}30` }}>
                    {cat} ({n})
                  </span>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── SEÇÃO PRODUTIVIDADE — POR CLIENTE ───────────────────────────────────────
function ProdPorCliente({ data }) {
  const [fSin, setFSin] = useState("Todos");
  const [fPessoa, setFPessoa] = useState("Todos");

  const base = data.filter(r => {
    if (fPessoa!=="Todos" && r.pessoa!==fPessoa) return false;
    return true;
  });

  const grupos = (() => {
    const by = {};
    base.forEach(r => {
      if (!by[r.grupo]) by[r.grupo] = { rows:[], sin:null };
      by[r.grupo].rows.push(r);
      if (!by[r.grupo].sin && r.sinalizacao) by[r.grupo].sin = r.sinalizacao;
    });
    return Object.entries(by)
      .map(([g,{rows,sin}]) => ({ g, rows, sin }))
      .filter(x => fSin==="Todos" || x.sin===fSin)
      .sort((a,b)=>a.g.localeCompare(b.g));
  })();

  return (
    <div>
      <div style={{ background:"#fff", borderRadius:12, padding:"10px 16px", marginBottom:20,
                    display:"flex", gap:12, alignItems:"center", flexWrap:"wrap",
                    boxShadow:"0 1px 3px rgba(0,0,0,.06)" }}>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Semáforo</span>
          {["Todos","Verde","Amarelo","Vermelho"].map(s => (
            <Pill key={s} label={s} active={fSin===s}
                  color={s==="Todos"?DARK:SIN_COR[s]}
                  onClick={()=>setFSin(s)}/>
          ))}
        </div>
        <div style={{ width:1, height:22, background:"#E2E8F0" }}/>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Responsável</span>
          {["Todos",...EQUIPE].map(p => (
            <Pill key={p} label={p} active={fPessoa===p}
                  color={COR_PESSOA[p]||DARK} onClick={()=>setFPessoa(p)}/>
          ))}
        </div>
        <span style={{ marginLeft:"auto", fontSize:12, color:"#94A3B8" }}>{grupos.length} clientes</span>
      </div>
      <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
        {grupos.map(({g,rows,sin}) => {
          const aprov = rows.filter(r=>r.status.toUpperCase().includes("APROVADO")).length;
          const valAprov = rows.reduce((s,r)=>s+r.valAprov,0);
          const valPleito = rows.reduce((s,r)=>s+r.valPleito,0);
          return (
            <div key={g} style={{ background:"#fff", borderRadius:12, overflow:"hidden",
                                   boxShadow:"0 1px 4px rgba(0,0,0,.07)",
                                   borderLeft:`4px solid ${sin?SIN_COR[sin]:"#E2E8F0"}` }}>
              <div style={{ padding:"12px 18px", display:"flex", alignItems:"center",
                             justifyContent:"space-between", flexWrap:"wrap", gap:10 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <span style={{ fontSize:14, fontWeight:700, color:DARK }}>{g}</span>
                  {sin && (
                    <span style={{ fontSize:11, fontWeight:600, padding:"2px 8px", borderRadius:99,
                                   background:SIN_BG[sin], color:SIN_COR[sin] }}>
                      {sin}
                    </span>
                  )}
                  <span style={{ fontSize:11, color:"#94A3B8" }}>{rows.length} interaç{rows.length!==1?"ões":"ão"}</span>
                </div>
                <div style={{ display:"flex", gap:20 }}>
                  {[
                    { l:"Valor pleito", v:brl(valPleito) },
                    { l:"Valor aprovado", v:brl(valAprov), c:"#16A34A" },
                    { l:"Aprovações", v:`${aprov}/${rows.length}` },
                  ].map(k => (
                    <div key={k.l} style={{ textAlign:"right" }}>
                      <div style={{ fontSize:10, color:"#94A3B8", marginBottom:2 }}>{k.l}</div>
                      <div style={{ fontSize:13, fontWeight:600, color:k.c||DARK }}>{k.v}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          );
        })}
        {!grupos.length && (
          <div style={{ background:"#fff", borderRadius:12, padding:"48px 0",
                        textAlign:"center", color:"#94A3B8", fontSize:14 }}>
            Nenhum cliente encontrado.
          </div>
        )}
      </div>
    </div>
  );
}

// ─── SEÇÃO PRODUTIVIDADE — HISTÓRICO ─────────────────────────────────────────
function ProdHistorico({ data }) {
  const [fPessoa, setFPessoa] = useState("Todos");
  const [fCat,    setFCat]    = useState("Todos");
  const [fStat,   setFStat]   = useState("Todos");
  const [busca,   setBusca]   = useState("");

  const filtrado = data.filter(r => {
    if (fPessoa!=="Todos" && r.pessoa!==fPessoa) return false;
    if (fCat!=="Todos" && r.categoria!==fCat) return false;
    if (fStat==="Aprovado" && !r.status.toUpperCase().includes("APROVADO")) return false;
    if (fStat==="Em Negociação" && !r.status.toUpperCase().includes("NEGOCI")) return false;
    if (fStat==="Recusado" && !r.status.toUpperCase().includes("RECUS")) return false;
    if (busca && !r.grupo.toUpperCase().includes(busca.toUpperCase()) &&
                 !r.nProp.includes(busca) && !r.ativ.includes(busca.toUpperCase())) return false;
    return true;
  });

  return (
    <div>
      <div style={{ background:"#fff", borderRadius:12, padding:"12px 16px", marginBottom:16,
                    display:"flex", gap:12, flexWrap:"wrap", alignItems:"center",
                    boxShadow:"0 1px 3px rgba(0,0,0,.06)" }}>
        <input value={busca} onChange={e=>setBusca(e.target.value)}
          placeholder="Buscar cliente ou proposta..."
          style={{ padding:"6px 12px", borderRadius:8, border:"1px solid #E2E8F0",
                   fontSize:12, outline:"none", width:200, fontFamily:"inherit" }}/>
        <div style={{ width:1, height:22, background:"#E2E8F0" }}/>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Responsável</span>
          {["Todos",...EQUIPE].map(p => (
            <Pill key={p} label={p} active={fPessoa===p} color={COR_PESSOA[p]||DARK} onClick={()=>setFPessoa(p)}/>
          ))}
        </div>
        <div style={{ width:1, height:22, background:"#E2E8F0" }}/>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Status</span>
          {["Todos","Aprovado","Em Negociação","Recusado"].map(s => (
            <Pill key={s} label={s} active={fStat===s} onClick={()=>setFStat(s)}/>
          ))}
        </div>
        <div style={{ width:1, height:22, background:"#E2E8F0" }}/>
        <div style={{ display:"flex", gap:5, alignItems:"center", flexWrap:"wrap" }}>
          <span style={{ fontSize:11, fontWeight:600, color:"#94A3B8" }}>Tipo</span>
          {["Todos",...TIPO_LIST_P].map(t => (
            <Pill key={t} label={t} active={fCat===t} color={TIPO_COR_P[t]||DARK} onClick={()=>setFCat(t)}/>
          ))}
        </div>
        <span style={{ marginLeft:"auto", fontSize:12, color:"#94A3B8" }}>{filtrado.length} registros</span>
      </div>
      <div style={{ background:"#fff", borderRadius:12, overflow:"hidden",
                    boxShadow:"0 1px 4px rgba(0,0,0,.07)" }}>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
            <thead>
              <tr style={{ background:"#F8FAFC" }}>
                {["Responsável","Mês","Cliente","Proposta","Tipo","Escopo","Val. Pleito","Val. Aprovado","% Aceito","Status","Rev."].map(h => (
                  <th key={h} style={{ textAlign:"left", padding:"9px 12px", fontSize:11,
                                       color:"#94A3B8", fontWeight:600, whiteSpace:"nowrap",
                                       borderBottom:"1px solid #E2E8F0" }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtrado.map((r,i) => {
                const isAprov = r.status.toUpperCase().includes("APROVADO");
                const isRecus = r.status.toUpperCase().includes("RECUS");
                const sc = isAprov ? {text:"#16A34A",bg:"#DCFCE7"} : isRecus ? {text:"#DC2626",bg:"#FEE2E2"} : {text:"#D97706",bg:"#FEF9C3"};
                return (
                  <tr key={i} style={{ background:i%2===0?"#fff":"#F8FAFC",
                                       borderTop:"1px solid #F1F4F8" }}>
                    <td style={{ padding:"8px 12px", borderLeft:`3px solid ${COR_PESSOA[r.pessoa]||"#E2E8F0"}`,
                                  color:COR_PESSOA[r.pessoa], fontWeight:600, whiteSpace:"nowrap" }}>{r.pessoa}</td>
                    <td style={{ padding:"8px 12px", color:"#94A3B8", whiteSpace:"nowrap" }}>{r.mes}</td>
                    <td style={{ padding:"8px 12px", maxWidth:120, overflow:"hidden", textOverflow:"ellipsis",
                                  whiteSpace:"nowrap", color:DARK }}>{r.grupo}</td>
                    <td style={{ padding:"8px 12px", color:"#94A3B8", whiteSpace:"nowrap", fontFamily:"monospace" }}>{r.nProp}</td>
                    <td style={{ padding:"8px 12px", whiteSpace:"nowrap" }}>
                      <span style={{ fontSize:11, fontWeight:600, padding:"2px 8px", borderRadius:99,
                                     background:`${TIPO_COR_P[r.categoria]||"#6B7280"}18`,
                                     color:TIPO_COR_P[r.categoria]||"#6B7280" }}>{r.categoria}</span>
                    </td>
                    <td style={{ padding:"8px 12px", whiteSpace:"nowrap", color:"#94A3B8", fontSize:11 }}>
                      {r.escopo||"—"}
                    </td>
                    <td style={{ padding:"8px 12px", whiteSpace:"nowrap", fontVariantNumeric:"tabular-nums" }}>
                      {r.valPleito>0?brl(r.valPleito):"—"}
                    </td>
                    <td style={{ padding:"8px 12px", color:"#16A34A", whiteSpace:"nowrap", fontVariantNumeric:"tabular-nums" }}>
                      {r.valAprov>0?brl(r.valAprov):"—"}
                    </td>
                    <td style={{ padding:"8px 12px", color:"#2563EB", whiteSpace:"nowrap" }}>
                      {r.pctAceito>0?pct(r.pctAceito):"—"}
                    </td>
                    <td style={{ padding:"8px 12px" }}>
                      <span style={{ fontSize:11, fontWeight:500, padding:"2px 8px", borderRadius:99,
                                     background:sc.bg, color:sc.text, whiteSpace:"nowrap" }}>{r.status}</span>
                    </td>
                    <td style={{ padding:"8px 12px", textAlign:"center",
                                  color:r.isRevisao?"#DC2626":"#94A3B8", fontSize:11 }}>
                      {r.isRevisao?"Rev":"—"}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {!filtrado.length && (
            <div style={{ textAlign:"center", padding:"40px 0", color:"#94A3B8", fontSize:14 }}>
              Nenhum registro encontrado.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── APP PRINCIPAL ────────────────────────────────────────────────────────────
export default function App() {
  const [gData,  setGData]  = useState([]);
  const [pData,  setPData]  = useState([]);
  const [loading, setLoading] = useState(true);
  const [erro,    setErro]   = useState(null);
  const [ultimaAt, setUltimaAt] = useState(null);

  // Menu: secao + subpagina
  const [secao,   setSecao]   = useState("gerencial");
  const [subPag,  setSubPag]  = useState("visao");
  const [menuOpen, setMenuOpen] = useState(true);

  const buscarDados = useCallback(async () => {
    setLoading(true); setErro(null);
    try {
      const [rawG, rawP] = await Promise.all([
        fetchCSV("2073814116"),
        fetchCSV("1622380363"),
      ]);
      setGData(processGerencial(rawG));
      setPData(processProdutividade(rawP));
      setUltimaAt(new Date());
    } catch(e) {
      setErro(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { buscarDados(); }, [buscarDados]);

  const navItems = [
    { id:"gerencial", label:"Gerencial", subs:[
      { id:"visao", label:"Visão Geral" },
    ]},
    { id:"produtividade", label:"Produtividade", subs:[
      { id:"visao",    label:"Visão Geral" },
      { id:"clientes", label:"Por Cliente" },
      { id:"historico",label:"Histórico" },
    ]},
  ];

  const MENU_W = menuOpen ? 220 : 64;

  return (
    <div style={{ display:"flex", minHeight:"100vh", background:"#F1F4F8",
                  fontFamily:"system-ui,-apple-system,sans-serif", color:DARK }}>

      {/* ── MENU LATERAL ─────────────────────────────────────────────────── */}
      <div style={{ width:MENU_W, minHeight:"100vh", background:MENU_BG,
                    display:"flex", flexDirection:"column", transition:"width .2s",
                    position:"fixed", top:0, left:0, zIndex:100, overflow:"hidden" }}>

        {/* Logo */}
        <div style={{ padding: menuOpen?"20px 20px 16px":"20px 0 16px",
                      display:"flex", alignItems:"center",
                      justifyContent: menuOpen?"space-between":"center",
                      borderBottom:"1px solid rgba(255,255,255,.08)" }}>
          {menuOpen && (
            <div style={{ lineHeight:1 }}>
              <div style={{ fontSize:10, fontWeight:400, color:"rgba(255,255,255,.4)",
                            letterSpacing:".12em" }}>GRUPO</div>
              <div style={{ fontSize:26, fontWeight:800, color:"#fff", letterSpacing:"-.01em" }}>GPS</div>
            </div>
          )}
          <button onClick={()=>setMenuOpen(o=>!o)} style={{
            width:32, height:32, borderRadius:8, border:"none",
            background:"rgba(255,255,255,.08)", color:"rgba(255,255,255,.6)",
            cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center",
            fontSize:14, flexShrink:0
          }}>
            {menuOpen ? "◀" : "▶"}
          </button>
        </div>

        {/* Nav */}
        <nav style={{ flex:1, padding:"12px 0", overflowY:"auto" }}>
          {navItems.map(({ id, label, subs }) => (
            <div key={id}>
              {/* Cabeçalho da seção */}
              <div onClick={()=>{ setSecao(id); setSubPag(subs[0].id); }}
                   style={{ padding: menuOpen?"10px 20px":"10px 0", display:"flex",
                             alignItems:"center", gap:10, cursor:"pointer",
                             justifyContent: menuOpen?"flex-start":"center",
                             background: secao===id?"rgba(255,255,255,.08)":"transparent",
                             borderLeft: secao===id?"3px solid #60A5FA":"3px solid transparent",
                             transition:"all .15s" }}>
                {!menuOpen && (
                  <span style={{ fontSize:12, fontWeight:700, color: secao===id?"#60A5FA":"rgba(255,255,255,.5)",
                                  width:32, textAlign:"center" }}>
                    {label.charAt(0)}
                  </span>
                )}
                {menuOpen && (
                  <span style={{ fontSize:13, fontWeight:600,
                                  color: secao===id?"#fff":"rgba(255,255,255,.6)" }}>
                    {label}
                  </span>
                )}
              </div>
              {/* Sub-itens */}
              {menuOpen && secao===id && subs.length>1 && subs.map(sub => (
                <div key={sub.id} onClick={()=>setSubPag(sub.id)}
                     style={{ padding:"7px 20px 7px 48px", cursor:"pointer", fontSize:12,
                               color: subPag===sub.id?"#60A5FA":"rgba(255,255,255,.45)",
                               fontWeight: subPag===sub.id?600:400,
                               background: subPag===sub.id?"rgba(96,165,250,.08)":"transparent",
                               transition:"all .15s" }}>
                  {sub.label}
                </div>
              ))}
            </div>
          ))}
        </nav>

        {/* Última atualização */}
        {menuOpen && ultimaAt && (
          <div style={{ padding:"12px 20px", borderTop:"1px solid rgba(255,255,255,.08)",
                        fontSize:10, color:"rgba(255,255,255,.3)" }}>
            Atualizado às {ultimaAt.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}
          </div>
        )}
      </div>

      {/* ── CONTEÚDO ─────────────────────────────────────────────────────── */}
      <div style={{ flex:1, marginLeft:MENU_W, transition:"margin-left .2s", display:"flex", flexDirection:"column" }}>

        {/* Header */}
        <div style={{ background:DARK, padding:"0 28px", display:"flex", alignItems:"center",
                      justifyContent:"space-between", minHeight:60, gap:16, flexShrink:0 }}>
          <div>
            <div style={{ fontSize:15, fontWeight:700, color:"#fff" }}>
              {navItems.find(n=>n.id===secao)?.label} —{" "}
              {navItems.find(n=>n.id===secao)?.subs.find(s=>s.id===subPag)?.label}
            </div>
            <div style={{ fontSize:11, color:"rgba(255,255,255,.4)" }}>Grupo GPS · Controle Comercial</div>
          </div>
          <button onClick={buscarDados} disabled={loading} style={{
            padding:"7px 16px", borderRadius:8, background:"rgba(255,255,255,.1)",
            border:"1px solid rgba(255,255,255,.2)", color:"#fff", fontSize:12,
            fontWeight:600, cursor:loading?"not-allowed":"pointer",
            display:"flex", alignItems:"center", gap:6, opacity:loading?.7:1
          }}>
            <span style={{ display:"inline-block" }}>{loading?"⏳":"⟳"}</span>
            {loading?"Atualizando...":"Atualizar"}
          </button>
        </div>

        {/* Conteúdo principal */}
        <div style={{ flex:1, padding:"24px 28px", overflowY:"auto" }}>
          {loading && (
            <div style={{ background:"#fff", borderRadius:14, padding:"80px 0",
                          textAlign:"center", boxShadow:"0 1px 4px rgba(0,0,0,.07)" }}>
              <div style={{ fontSize:36, marginBottom:12 }}>⏳</div>
              <div style={{ fontSize:15, fontWeight:600, color:DARK, marginBottom:6 }}>Carregando dados...</div>
              <div style={{ fontSize:13, color:"#94A3B8" }}>Buscando planilha do Google Sheets</div>
            </div>
          )}
          {erro && !loading && (
            <div style={{ background:"#FEF2F2", borderRadius:14, padding:"40px",
                          textAlign:"center", border:"1px solid #FECACA" }}>
              <div style={{ fontSize:36, marginBottom:12 }}>⚠️</div>
              <div style={{ fontSize:15, fontWeight:600, color:"#DC2626", marginBottom:8 }}>
                Não foi possível carregar os dados
              </div>
              <div style={{ fontSize:13, color:"#EF4444", marginBottom:20 }}>{erro}</div>
              <button onClick={buscarDados} style={{ padding:"10px 24px", borderRadius:8,
                background:"#DC2626", border:"none", color:"#fff", fontSize:13,
                fontWeight:600, cursor:"pointer" }}>Tentar novamente</button>
            </div>
          )}
          {!loading && !erro && secao==="gerencial" && (
            <SecaoGerencial data={gData} setData={setGData}/>
          )}
          {!loading && !erro && secao==="produtividade" && subPag==="visao" && (
            <ProdVisaoGeral data={pData}/>
          )}
          {!loading && !erro && secao==="produtividade" && subPag==="clientes" && (
            <ProdPorCliente data={pData}/>
          )}
          {!loading && !erro && secao==="produtividade" && subPag==="historico" && (
            <ProdHistorico data={pData}/>
          )}
        </div>
      </div>
    </div>
  );
}
