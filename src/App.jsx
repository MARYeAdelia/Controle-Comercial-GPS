import React, { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const ORDEM_MES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho",
                   "Agosto","Setembro","Outubro","Novembro","Dezembro"];

const STATUS_LIST = ["Concluído","Pendente Cliente","Pendente GPS","N/A","Negócio Perdido"];
const TIPO_LIST   = ["Reajuste","DT","Renovação","Negócio Perdido"];

const STATUS_COR = {
  "Concluído":        { dot:"#16A34A", text:"#15803D", bg:"#DCFCE7" },
  "Pendente Cliente": { dot:"#D97706", text:"#B45309", bg:"#FEF9C3" },
  "Pendente GPS":     { dot:"#2563EB", text:"#1D4ED8", bg:"#DBEAFE" },
  "N/A":              { dot:"#9CA3AF", text:"#6B7280", bg:"#F3F4F6" },
  "Negócio Perdido":  { dot:"#DC2626", text:"#B91C1C", bg:"#FEE2E2" },
};

const DARK  = "#1a2332";
const DARK2 = "#243044";

// ─── PARSE EXCEL ─────────────────────────────────────────────────────────────
const parseExcel = file => new Promise((res, rej) => {
  const rd = new FileReader();
  rd.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }));
    } catch(err) { rej(err); }
  };
  rd.onerror = rej;
  rd.readAsArrayBuffer(file);
});

const processRows = raw => {
  if (!raw?.length) return [];
  const keys = Object.keys(raw[0]);
  const find = (...kws) => keys.find(k => kws.some(kw => k.toUpperCase().includes(kw.toUpperCase()))) || null;
  const get  = (r, ...kws) => { const k = find(...kws); return k ? (r[k] ?? "") : ""; };
  return raw.map((r, i) => {
    const grupo  = (get(r,"Grupo Cliente")||"").toString().trim();
    const cr     = (get(r,"CR")||"").toString().trim();
    if (!grupo || !cr) return null;
    return {
      _id:    i,
      grupo,
      cr,
      descr:  (get(r,"DESCRI")||"").toString().trim(),
      tipo:   (get(r,"Tipo de Negócio","Tipo")||"").toString().trim() || "Reajuste",
      status: (get(r,"Status Real","Status")||"").toString().trim() || "N/A",
      mes:    (get(r,"Mês Vigência","MES")||"").toString().trim(),
      farmer: (get(r,"Responsável Farmer","FARMER")||"").toString().trim(),
      info:   (get(r,"Informações")||"").toString().trim(),
      _edited: false,
    };
  }).filter(Boolean);
};

// ─── BUSCA GOOGLE SHEETS ─────────────────────────────────────────────────────

const fetchSheetData = async () => {
  // Chama nossa função serverless no Vercel — sem bloqueio CORS
  const res = await fetch("/api/sheet?sheet=Gerencial");
  if (!res.ok) throw new Error(`Erro HTTP ${res.status}`);
  const text = await res.text();
  if (!text) throw new Error("Planilha vazia");

  const parseCSVLine = line => {
    const result = [];
    let cur = "", inQ = false;
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
    headers.forEach((h, i) => { obj[h] = (vals[i] ?? "").replace(/^"|"$/g, ""); });
    return obj;
  }).filter(r => Object.values(r).some(v => v));
};

// ─── EXPORTAR EXCEL ──────────────────────────────────────────────────────────
const exportExcel = (data) => {
  const rows = data.map(r => ({
    "Grupo Cliente":       r.grupo,
    "CR":                  r.cr,
    "Descrição CR":        r.descr,
    "Tipo de Negócio":     r.tipo,
    "Mês Vigência":        r.mes,
    "Responsável Farmer":  r.farmer,
    "Status Real":         r.status,
    "Informações":         r.info,
    "Editado":             r._edited ? "Sim" : "",
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Dados");
  XLSX.writeFile(wb, `Reajustes_Contratos_2026_Atualizado.xlsx`);
};

// ─── PILL BUTTON ─────────────────────────────────────────────────────────────
const Pill = ({ label, active, onClick }) => (
  <button onClick={onClick} style={{
    padding:"4px 12px", borderRadius:99, border:"none", cursor:"pointer",
    fontFamily:"inherit", fontSize:11, fontWeight:500, transition:"all .15s",
    background: active ? DARK : "#F1F4F8",
    color:      active ? "#fff" : "#64748B",
  }}>{label}</button>
);

// ─── DROPDOWN INLINE ─────────────────────────────────────────────────────────
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
      const menuHeight = options.length * 42;
      // Se não há espaço suficiente embaixo, abre para cima
      const openUp = spaceBelow < menuHeight + 8;
      setPos({
        top: openUp ? r.top - menuHeight - 4 : r.bottom + 4,
        left: r.left,
        openUp,
      });
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
          <div onClick={()=>setOpen(false)}
               style={{ position:"fixed", inset:0, zIndex:500 }}/>
          <div style={{ position:"fixed", top:pos.top, left:pos.left, zIndex:501,
                        background:"#fff", borderRadius:10, boxShadow:"0 8px 24px rgba(0,0,0,.18)",
                        border:"1px solid #E2E8F0", minWidth:180, overflow:"hidden" }}>
            {options.map(opt => {
              const oc = colorMap?.[opt] || { dot:"#9CA3AF", text:"#374151", bg:"#F3F4F6" };
              return (
                <div key={opt}
                     onClick={e=>{ e.stopPropagation(); onChange(opt); setOpen(false); }}
                     style={{ display:"flex", alignItems:"center", gap:8, padding:"9px 14px",
                              cursor:"pointer", background: opt===value ? "#F8FAFC" : "#fff",
                              transition:"background .1s" }}
                     onMouseOver={e=>e.currentTarget.style.background="#F1F5F9"}
                     onMouseOut={e=>e.currentTarget.style.background=opt===value?"#F8FAFC":"#fff"}>
                  {colorMap && <div style={{ width:8, height:8, borderRadius:99, background:oc.dot }}/>}
                  <span style={{ fontSize:13, color:oc.text, fontWeight: opt===value?600:400 }}>{opt}</span>
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


// ─── BOTÃO EXPORTAR MINIMIZÁVEL ──────────────────────────────────────────────
const ExportBtn = ({ nEditados, onExport }) => {
  const [expanded, setExpanded] = useState(false);
  return (
    <div
      style={{
        position:"fixed", bottom:24, right:24, zIndex:300,
        display:"flex", flexDirection:"column", alignItems:"flex-end", gap:8,
      }}
    >
      {/* Painel expandido */}
      {expanded && (
        <div style={{
          background:"#fff", borderRadius:12, padding:"14px 18px",
          boxShadow:"0 8px 24px rgba(0,0,0,.18)", border:"1px solid #E2E8F0",
          display:"flex", flexDirection:"column", gap:10, minWidth:200,
          animation:"fadeIn .15s ease"
        }}>
          <div style={{fontSize:12,color:"#64748B"}}>
            <span style={{fontWeight:700,color:"#F59E0B"}}>{nEditados}</span> alteraç{nEditados===1?"ão":"ões"} pendente{nEditados!==1?"s":""}
          </div>
          <button onClick={onExport} style={{
            padding:"9px 16px", borderRadius:8, background:"#16A34A",
            border:"none", color:"#fff", fontSize:13, fontWeight:700,
            cursor:"pointer", display:"flex", alignItems:"center", gap:8,
            boxShadow:"0 2px 8px rgba(22,163,74,.35)"
          }}>
            ⬇ Exportar Excel
          </button>
          <button onClick={()=>setExpanded(false)} style={{
            padding:"6px", borderRadius:8, background:"#F1F5F9",
            border:"none", color:"#64748B", fontSize:12, cursor:"pointer"
          }}>
            Minimizar
          </button>
        </div>
      )}

      {/* Botão flutuante minimizado */}
      <button
        onClick={()=>setExpanded(o=>!o)}
        title={expanded ? "Minimizar" : `${nEditados} alteraç${nEditados===1?"ão":"ões"} — clique para exportar`}
        style={{
          width:48, height:48, borderRadius:99,
          background: expanded ? "#64748B" : "#16A34A",
          border:"none", color:"#fff", fontSize:20, cursor:"pointer",
          boxShadow:"0 4px 16px rgba(0,0,0,.25)",
          display:"flex", alignItems:"center", justifyContent:"center",
          position:"relative", transition:"background .2s"
        }}
      >
        {expanded ? "✕" : "⬇"}
        {/* Badge com contagem */}
        {!expanded && (
          <span style={{
            position:"absolute", top:-4, right:-4,
            background:"#F59E0B", color:"#fff",
            fontSize:10, fontWeight:700, minWidth:18, height:18,
            borderRadius:99, display:"flex", alignItems:"center",
            justifyContent:"center", padding:"0 4px",
            border:"2px solid #fff"
          }}>
            {nEditados}
          </span>
        )}
      </button>
    </div>
  );
};

// ─── APP ─────────────────────────────────────────────────────────────────────
export default function PainelGPS() {
  const [data,    setData]    = useState([]);
  const [arquivo, setArquivo] = useState(null);
  const [loading, setLoading] = useState(false);
  const [drag,    setDrag]    = useState(false);
  const [ultimaAtualizacao, setUltimaAtualizacao] = useState(null);
  const [erroFetch, setErroFetch] = useState(null);
  const [tooltip, setTooltip] = useState({ visible:false, text:"", x:0, y:0 });

  // Filtros globais
  const [fMeses,  setFMeses]  = useState(new Set());
  const [fFarmer, setFFarmer] = useState("Todos");
  const [fStatus, setFStatus] = useState(new Set()); // vazio = todos
  const [fTipo,   setFTipo]   = useState("Todos");
  const [busca,   setBusca]   = useState("");


  const carregar = useCallback(async file => {
    setLoading(true);
    try {
      const rows = processRows(await parseExcel(file));
      setData(rows);
      setArquivo(file.name);
    } catch(e) { console.error(e); }
    finally { setLoading(false); }
  }, []);

  // ── Busca automática ao abrir ────────────────────────────────────────────
  const buscarPlanilha = useCallback(async () => {
    setLoading(true); setErroFetch(null);
    try {
      const rows = processRows(await fetchSheetData());
      setData(rows);
      setArquivo("Controle Comercial - Farmer.xlsx");
      setUltimaAtualizacao(new Date());
    } catch(e) {
      console.error(e);
      setErroFetch("Não foi possível carregar a planilha. Verifique se ela está pública no Google Sheets.");
    }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { buscarPlanilha(); }, [buscarPlanilha]);

  // ── Edição inline ────────────────────────────────────────────────────────
  const editRow = (id, field, value) => {
    setData(prev => prev.map(r =>
      r._id === id ? { ...r, [field]: value, _edited: true } : r
    ));
  };

  // ── Dados ────────────────────────────────────────────────────────────────
  const mesesDisp  = ORDEM_MES.filter(m => data.some(r => r.mes === m));
  const todosMeses = fMeses.size === 0;
  const toggleMes  = m => setFMeses(prev => { const n=new Set(prev); n.has(m)?n.delete(m):n.add(m); return n; });
  const farmersDisp = ["Todos", ...Array.from(new Set(data.map(r=>r.farmer).filter(Boolean))).sort()];

  // Filtro global (mês + farmer + busca)
  const filtradoGlobal = data.filter(r => {
    if (!todosMeses && !fMeses.has(r.mes)) return false;
    if (fFarmer !== "Todos" && r.farmer !== fFarmer) return false;
    if (fStatus.size > 0 && !fStatus.has(r.status)) return false;
    if (fTipo   !== "Todos" && r.tipo   !== fTipo)   return false;
    if (busca && ![r.grupo,r.cr,r.descr].some(v=>v.toUpperCase().includes(busca.toUpperCase()))) return false;
    return true;
  });

  // Grupos
  const grupos = (() => {
    const by = {};
    filtradoGlobal.forEach(r => { if(!by[r.grupo]) by[r.grupo]=[]; by[r.grupo].push(r); });
    return Object.entries(by).map(([g,crs])=>({g,crs})).sort((a,b)=>a.g.localeCompare(b.g));
  })();

  // Cards — usa dados sem filtro de tipo/status
  const cardBase = filtradoGlobal;
  const buildCard = rows => ({
    total: rows.length,
    por: STATUS_LIST.map(s => ({ s, n: rows.filter(r=>r.status===s).length }))
  });
  const cardGeral  = buildCard(cardBase);
  const cardsTipo  = TIPO_LIST.map(t => ({ tipo:t, ...buildCard(cardBase.filter(r=>r.tipo===t)) })).filter(c=>c.total>0);

  const nEditados = data.filter(r=>r._edited).length;

  // Tooltip
  const showTip = (e,t) => { if(t) setTooltip({visible:true,text:t,x:e.clientX,y:e.clientY}); };
  const moveTip = e => { if(tooltip.visible) setTooltip(p=>({...p,x:e.clientX,y:e.clientY})); };
  const hideTip = () => setTooltip(p=>({...p,visible:false}));

  return (
    <div style={{minHeight:"100vh",background:"#F1F4F8",fontFamily:"system-ui,-apple-system,sans-serif",color:DARK}}>

      {/* TOOLTIP */}
      {tooltip.visible && tooltip.text && (
        <div style={{position:"fixed",zIndex:9999,pointerEvents:"none",
          left:tooltip.x+14,top:tooltip.y-10,background:DARK,color:"#F8FAFC",
          padding:"8px 12px",borderRadius:8,fontSize:12,maxWidth:320,lineHeight:1.5,
          boxShadow:"0 4px 16px rgba(0,0,0,.3)",whiteSpace:"pre-wrap",wordBreak:"break-word"}}>
          {tooltip.text}
        </div>
      )}

      {/* HEADER */}
      <div style={{background:DARK,padding:"0 28px",display:"flex",alignItems:"center",
                   justifyContent:"space-between",gap:16,minHeight:70}}>
        <div style={{display:"flex",alignItems:"center",gap:20}}>
          <div style={{display:"flex",flexDirection:"column",alignItems:"flex-start",lineHeight:1}}>
            <span style={{fontSize:11,fontWeight:400,color:"rgba(255,255,255,.5)",
                          letterSpacing:".12em",textTransform:"uppercase"}}>GRUPO</span>
            <span style={{fontSize:32,fontWeight:800,color:"#fff",letterSpacing:"-.01em",lineHeight:1}}>GPS</span>
          </div>
          <div style={{width:1,height:40,background:"rgba(255,255,255,.2)"}}/>
          <div>
            <div style={{fontSize:18,fontWeight:700,color:"#fff",letterSpacing:"-.01em"}}>
              Controle de Reajustes 2026
            </div>
            <div style={{fontSize:11,color:"rgba(255,255,255,.5)"}}>
              Visão gerencial consolidada — Grupo GPS
            </div>
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {nEditados > 0 && (
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:11,color:"rgba(255,255,255,.5)"}}>
                {nEditados} alteraç{nEditados===1?"ão":"ões"}
              </span>
              <button onClick={()=>exportExcel(data)} style={{
                padding:"8px 18px",borderRadius:8,background:"#16A34A",border:"none",
                color:"#fff",fontSize:12,fontWeight:600,cursor:"pointer",
                boxShadow:"0 2px 8px rgba(22,163,74,.4)"}}>
                ⬇ Exportar Excel
              </button>
            </div>
          )}
          {ultimaAtualizacao && (
            <span style={{fontSize:11,color:"rgba(255,255,255,.45)"}}>
              Atualizado às {ultimaAtualizacao.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}
            </span>
          )}
          {erroFetch && (
            <span style={{fontSize:11,color:"#FCA5A5",background:"rgba(220,38,38,.15)",
                          padding:"4px 10px",borderRadius:8,maxWidth:260}}>
              ⚠ {erroFetch}
            </span>
          )}
          <button onClick={buscarPlanilha} disabled={loading} style={{
            padding:"8px 16px", borderRadius:8, background:"rgba(255,255,255,.12)",
            border:"1px solid rgba(255,255,255,.2)", color:"#fff",
            fontSize:12, fontWeight:600, cursor:loading?"not-allowed":"pointer",
            display:"flex", alignItems:"center", gap:6, opacity:loading?0.6:1
          }}>
            <span style={{display:"inline-block", animation:loading?"spin 1s linear infinite":"none"}}>⟳</span>
            {loading ? "Atualizando..." : "Atualizar"}
          </button>
        </div>
      </div>

      <div style={{padding:"20px 28px"}}>

        {/* ESTADO DE CARREGAMENTO INICIAL */}
        {!arquivo && loading && (
          <div style={{background:"#fff",borderRadius:14,padding:"80px 40px",textAlign:"center",
                       boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
            <div style={{fontSize:36,marginBottom:16}}>⟳</div>
            <div style={{fontSize:16,fontWeight:600,color:DARK,marginBottom:8}}>
              Carregando planilha...
            </div>
            <div style={{fontSize:13,color:"#64748B"}}>
              Buscando dados do Google Sheets
            </div>
          </div>
        )}
        {/* ERRO DE CARREGAMENTO */}
        {!arquivo && !loading && erroFetch && (
          <div style={{background:"#FEF2F2",borderRadius:14,padding:"40px",textAlign:"center",
                       border:"1px solid #FECACA"}}>
            <div style={{fontSize:36,marginBottom:12}}>⚠️</div>
            <div style={{fontSize:15,fontWeight:600,color:"#DC2626",marginBottom:8}}>
              Não foi possível carregar a planilha
            </div>
            <div style={{fontSize:13,color:"#EF4444",marginBottom:20}}>{erroFetch}</div>
            <button onClick={buscarPlanilha} style={{
              padding:"10px 24px",borderRadius:8,background:"#DC2626",border:"none",
              color:"#fff",fontSize:13,fontWeight:600,cursor:"pointer"
            }}>Tentar novamente</button>
          </div>
        )}

        {arquivo && loading && (
          <div style={{background:"#fff",borderRadius:14,padding:"40px 0",textAlign:"center",
                       marginBottom:20,boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
            <div style={{fontSize:13,color:"#2563EB"}}>⟳ Atualizando dados...</div>
          </div>
        )}

        {arquivo && !loading && (<>

          {/* FILTROS GLOBAIS */}
          <div style={{background:"#fff",borderRadius:12,padding:"12px 18px",marginBottom:20,
                       display:"flex",gap:12,flexWrap:"wrap",alignItems:"center",
                       boxShadow:"0 1px 3px rgba(0,0,0,.06)"}}>
            <input value={busca} onChange={e=>setBusca(e.target.value)}
              placeholder="Buscar cliente, CR ou descrição..."
              style={{padding:"6px 12px",borderRadius:8,border:"1px solid #E2E8F0",
                      fontSize:12,color:DARK,outline:"none",width:220,fontFamily:"inherit"}}/>
            <div style={{width:1,height:22,background:"#E2E8F0"}}/>
            <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:11,fontWeight:600,color:"#94A3B8"}}>Mês</span>
              <Pill label="Todos" active={todosMeses} onClick={()=>setFMeses(new Set())}/>
              {mesesDisp.map(m=><Pill key={m} label={m} active={fMeses.has(m)} onClick={()=>toggleMes(m)}/>)}
            </div>
            <div style={{width:1,height:22,background:"#E2E8F0"}}/>
            <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:11,fontWeight:600,color:"#94A3B8"}}>Farmer</span>
              {farmersDisp.map(f=>(
                <Pill key={f} label={f} active={fFarmer===f} onClick={()=>setFFarmer(f)}/>
              ))}
            </div>
            <div style={{width:1,height:22,background:"#E2E8F0"}}/>
            <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:11,fontWeight:600,color:"#94A3B8"}}>Status</span>
              <Pill label="Todos" active={fStatus.size===0} onClick={()=>setFStatus(new Set())}/>
              {STATUS_LIST.map(s=>(
                <Pill key={s} label={s} active={fStatus.has(s)}
                  onClick={()=>setFStatus(prev=>{const n=new Set(prev);n.has(s)?n.delete(s):n.add(s);return n;})}/>
              ))}
            </div>
            <div style={{width:1,height:22,background:"#E2E8F0"}}/>
            <div style={{display:"flex",gap:5,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:11,fontWeight:600,color:"#94A3B8"}}>Tipo</span>
              {["Todos",...TIPO_LIST].map(t=>(
                <Pill key={t} label={t} active={fTipo===t} onClick={()=>setFTipo(t)}/>
              ))}
            </div>
            <span style={{marginLeft:"auto",fontSize:12,color:"#94A3B8",whiteSpace:"nowrap"}}>
              {filtradoGlobal.length} CRs · {grupos.length} grupos
            </span>
          </div>

          {/* CARDS */}
          <div style={{display:"grid",gridTemplateColumns:`repeat(${1+cardsTipo.length},1fr)`,
                       gap:14,marginBottom:22}}>
            {/* Card Geral */}
            {(()=>{
              const conc=cardGeral.por.find(p=>p.s==="Concluído")?.n||0;
              const pctC=cardGeral.total>0?Math.round(conc/cardGeral.total*100):0;
              return (
                <div style={{background:DARK,borderRadius:14,padding:"20px 22px",
                             boxShadow:"0 4px 14px rgba(26,35,50,.25)"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
                    <div>
                      <div style={{fontSize:11,color:"rgba(255,255,255,.5)",marginBottom:4,fontWeight:500}}>Total Geral</div>
                      <div style={{fontSize:36,fontWeight:800,color:"#fff",lineHeight:1}}>{cardGeral.total}</div>
                      <div style={{fontSize:11,color:"rgba(255,255,255,.4)",marginTop:4}}>{pctC}% concluído</div>
                    </div>
                    <div style={{fontSize:28,opacity:.3}}>📋</div>
                  </div>
                  {cardGeral.por.map(({s,n})=>{
                    const c=STATUS_COR[s];
                    const p=cardGeral.total>0?Math.round(n/cardGeral.total*100):0;
                    return (
                      <div key={s} style={{display:"flex",alignItems:"center",gap:8,marginBottom:7}}>
                        <div style={{width:8,height:8,borderRadius:99,background:c.dot,flexShrink:0}}/>
                        <span style={{fontSize:12,color:"rgba(255,255,255,.6)",flex:1}}>{s}</span>
                        <span style={{fontSize:13,fontWeight:700,color:"#fff",minWidth:22,textAlign:"right"}}>{n}</span>
                        <span style={{fontSize:11,color:"rgba(255,255,255,.35)",minWidth:36,textAlign:"right"}}>({p}%)</span>
                      </div>
                    );
                  })}
                </div>
              );
            })()}
            {/* Cards por tipo */}
            {cardsTipo.map(({tipo,total,por})=>{
              const conc=por.find(p=>p.s==="Concluído")?.n||0;
              const pctC=total>0?Math.round(conc/total*100):0;
              return (
                <div key={tipo} style={{background:"#fff",borderRadius:14,padding:"20px 22px",
                                        borderLeft:`4px solid ${DARK}`,boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
                    <div>
                      <div style={{fontSize:11,color:"#94A3B8",marginBottom:4,fontWeight:500}}>{tipo}</div>
                      <div style={{fontSize:36,fontWeight:800,color:DARK,lineHeight:1}}>{total}</div>
                      <div style={{fontSize:11,color:"#94A3B8",marginTop:4}}>{pctC}% concluído</div>
                    </div>
                    <div style={{background:"#F1F4F8",borderRadius:8,padding:"6px 10px",
                                 fontSize:12,fontWeight:700,color:DARK}}>{tipo}</div>
                  </div>
                  {por.map(({s,n})=>{
                    const c=STATUS_COR[s];
                    const p=total>0?Math.round(n/total*100):0;
                    return (
                      <div key={s} style={{display:"flex",alignItems:"center",gap:8,marginBottom:7}}>
                        <div style={{width:8,height:8,borderRadius:99,background:c.dot,flexShrink:0}}/>
                        <span style={{fontSize:12,color:"#64748B",flex:1}}>{s}</span>
                        <span style={{fontSize:13,fontWeight:700,color:DARK,minWidth:22,textAlign:"right"}}>{n}</span>
                        <span style={{fontSize:11,color:"#94A3B8",minWidth:36,textAlign:"right"}}>({p}%)</span>
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>

          {/* GRUPOS */}
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            {grupos.length===0 && (
              <div style={{background:"#fff",borderRadius:14,padding:"48px 0",
                           textAlign:"center",color:"#94A3B8",fontSize:14}}>
                Nenhum CR encontrado.
              </div>
            )}
            {grupos.map(({g, crs})=>{
              const crsFilt = crs;
              const nEdit = crs.filter(r=>r._edited).length;

              return (
                <div key={g} style={{background:"#fff",borderRadius:14,overflow:"hidden",
                                     boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>

                  {/* CABEÇALHO DO GRUPO */}
                  <div style={{background:DARK2,padding:"12px 18px",display:"flex",
                               alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
                    <div style={{display:"flex",alignItems:"center",gap:10}}>
                      <span style={{fontSize:14,fontWeight:700,color:"#fff"}}>{g}</span>
                      <span style={{fontSize:11,color:"rgba(255,255,255,.4)"}}>
                        {crs.length} CR{crs.length!==1?"s":""}
                      </span>
                      {nEdit>0 && (
                        <span style={{fontSize:10,background:"#F59E0B",color:"#fff",
                                      padding:"1px 7px",borderRadius:99,fontWeight:600}}>
                          {nEdit} editado{nEdit!==1?"s":""}
                        </span>
                      )}
                    </div>
                    <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                      {STATUS_LIST.map(s=>{
                        const n=crs.filter(c=>c.status===s).length;
                        if(!n) return null;
                        const sc=STATUS_COR[s];
                        return (
                          <span key={s} style={{fontSize:11,fontWeight:600,padding:"2px 10px",
                                                borderRadius:99,background:sc.bg,color:sc.text}}>
                            {n} {s}
                          </span>
                        );
                      })}
                    </div>
                  </div>

                  {/* TABELA DO GRUPO */}
                  {crsFilt.length===0 ? (
                    <div style={{padding:"20px",textAlign:"center",color:"#94A3B8",fontSize:12}}>
                      Nenhum CR com os filtros selecionados.
                    </div>
                  ) : (
                    <table style={{width:"100%",borderCollapse:"collapse"}}>
                      <thead>
                        <tr style={{background:"#F8FAFC"}}>
                          {["CR","Descrição","Mês","Farmer","Tipo","Status"].map((h,i)=>(
                            <th key={h} style={{textAlign:i>2?"center":"left",padding:"8px 16px",
                                               fontSize:11,color:"#94A3B8",fontWeight:600,
                                               borderBottom:"1px solid #E2E8F0",whiteSpace:"nowrap"}}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {crsFilt.map((cr,ci)=>(
                          <tr key={cr._id}
                              style={{background:cr._edited?"#FFFBEB":ci%2===0?"#fff":"#FAFBFC",
                                      borderTop:"1px solid #F1F4F8",transition:"background .1s"}}
                              onMouseOver={e=>{if(!cr._edited)e.currentTarget.style.background="#EFF6FF";}}
                              onMouseOut={e=>{e.currentTarget.style.background=cr._edited?"#FFFBEB":ci%2===0?"#fff":"#FAFBFC";}}>
                            <td style={{padding:"9px 16px",fontSize:12,color:"#64748B",
                                        fontFamily:"monospace",whiteSpace:"nowrap"}}>
                              {cr._edited && <span style={{color:"#F59E0B",marginRight:4}}>●</span>}
                              {cr.cr}
                            </td>
                            <td style={{padding:"9px 16px",fontSize:12,color:DARK,maxWidth:340,
                                        cursor:cr.info?"help":"default"}}
                                onMouseEnter={e=>showTip(e,cr.info)}
                                onMouseMove={moveTip}
                                onMouseLeave={hideTip}>
                              <span style={{display:"block",overflow:"hidden",
                                           textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                                {cr.descr}
                              </span>
                            </td>
                            <td style={{padding:"9px 16px",fontSize:12,color:"#64748B",
                                        textAlign:"center",whiteSpace:"nowrap"}}>{cr.mes}</td>
                            <td style={{padding:"9px 16px",fontSize:12,color:"#64748B",
                                        textAlign:"center",whiteSpace:"nowrap"}}>{cr.farmer||"—"}</td>
                            {/* TIPO — editável */}
                            <td style={{padding:"6px 16px",textAlign:"center"}}>
                              <InlineSelect
                                value={cr.tipo}
                                options={TIPO_LIST}
                                onChange={v=>editRow(cr._id,"tipo",v)}
                              />
                            </td>
                            {/* STATUS — editável */}
                            <td style={{padding:"6px 16px",textAlign:"center"}}>
                              <InlineSelect
                                value={cr.status}
                                options={STATUS_LIST}
                                colorMap={STATUS_COR}
                                onChange={v=>editRow(cr._id,"status",v)}
                              />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              );
            })}
          </div>

          {/* BOTÃO EXPORTAR MINIMIZÁVEL */}
          {nEditados>0 && <ExportBtn nEditados={nEditados} onExport={()=>exportExcel(data)}/>}

        </>)}
      </div>
    </div>
  );
}
