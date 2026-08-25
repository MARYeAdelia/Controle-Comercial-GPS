import React, { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── CONSTANTES ───────────────────────────────────────────────────────────────
const ORDEM_MES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho",
                   "Agosto","Setembro","Outubro","Novembro","Dezembro"];

const STATUS_G  = ["Concluído","Pendente Cliente","Pendente GPS","N/A","Negócio Perdido",
                   "Em Negociação","Em Negociação Antecipada","Aguardando Reajuste"];
const STATUS_COR = {
  "Concluído":               { dot:"#16A34A", text:"#15803D", bg:"#DCFCE7" },
  "Pendente Cliente":        { dot:"#D97706", text:"#B45309", bg:"#FEF9C3" },
  "Pendente GPS":            { dot:"#2563EB", text:"#1D4ED8", bg:"#DBEAFE" },
  "N/A":                     { dot:"#9CA3AF", text:"#6B7280", bg:"#F3F4F6" },
  "Negócio Perdido":         { dot:"#DC2626", text:"#B91C1C", bg:"#FEE2E2" },
  "Em Negociação":           { dot:"#D97706", text:"#B45309", bg:"#FEF9C3" },
  "Em Negociação Antecipada":{ dot:"#7C3AED", text:"#6D28D9", bg:"#F5F3FF" },
  "Aguardando Reajuste":     { dot:"#9CA3AF", text:"#6B7280", bg:"#F3F4F6" },
};
const TIPO_G = ["Reajuste","DT","Renovação","Negócio Perdido"];
const TIPO_P = ["Reajuste","Renovação","Up Selling","Defesa de Território","Alteração de Escopo","BID/Cotação","Outros"];
const TIPO_COR_P = {
  "Reajuste":"#7C3AED","Renovação":"#0891B2","Up Selling":"#059669",
  "Defesa de Território":"#D97706","Alteração de Escopo":"#0369A1",
  "BID/Cotação":"#DC2626","Outros":"#6B7280"
};
const EQUIPE = ["Mariana","Wilder","Giovanni","Carla","Darlan"];
const COR_P  = { Mariana:"#7C3AED",Wilder:"#0369A1",Giovanni:"#059669",Carla:"#D97706",Darlan:"#DC2626" };
const SIN_COR = { Verde:"#16A34A",Amarelo:"#D97706",Vermelho:"#DC2626" };
const SIN_BG  = { Verde:"#DCFCE7",Amarelo:"#FEF9C3",Vermelho:"#FEE2E2" };
const DARK  = "#1a2332";
const DARK2 = "#243044";
const MENU_BG = "#0F1F35";
const MENU_W_OPEN  = 220;
const MENU_W_CLOSE = 48;

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const brl = v => (!v&&v!==0)?"—"
  : new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL",maximumFractionDigits:0}).format(v);
const num = v => { const n=parseFloat(String(v).replace(/[R$\s.]/g,"").replace(",",".")); return isNaN(n)?0:n; };
const pctFmt = v => {
  if (!v||v===""||v==="0") return "—";
  const n = parseFloat(v);
  if (isNaN(n)||n===0) return "—";
  const val = n>1?n:n*100;
  return `${val.toFixed(2)}%`;
};
const fmtData = d => {
  if (!d) return "—";
  return d.toLocaleDateString("pt-BR",{day:"2-digit",month:"2-digit",year:"numeric"});
};
const parseData = s => {
  if (!s) return null;
  const str = s.toString().trim();
  if (!str||str==="-"||str==="NaT") return null;
  const d = new Date(str); return isNaN(d.getTime())?null:d;
};
const diasEntre = (d1,d2) => {
  if (!d1||!d2) return null;
  return Math.round(Math.abs(d2-d1)/(1000*60*60*24));
};
const MESES_MAP = {
  "janeiro":1,"fevereiro":2,"março":3,"abril":4,"maio":5,"junho":6,
  "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12,
  "jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,
  "jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12,
};
const parseMesNum = s => {
  if (!s) return null;
  return MESES_MAP[s.toString().toLowerCase().replace(/\.$/,"").trim()] || null;
};
const isPessoa = s => EQUIPE.some(p => s.toUpperCase().includes(p.toUpperCase()));

// ─── FETCH ────────────────────────────────────────────────────────────────────
const fetchCSV = async gid => {
  const res = await fetch(`/api/sheet?gid=${gid}`);
  if (!res.ok) throw new Error(`Erro HTTP ${res.status}`);
  const text = await res.text();
  if (!text||text.length<10) throw new Error("Planilha vazia");
  const parseCSVLine = line => {
    const result=[]; let cur="",inQ=false;
    for (let i=0;i<line.length;i++) {
      const ch=line[i];
      if (ch==='"'&&line[i+1]==='"') { cur+='"';i++; }
      else if (ch==='"') { inQ=!inQ; }
      else if (ch===','&&!inQ) { result.push(cur.trim());cur=""; }
      else { cur+=ch; }
    }
    result.push(cur.trim()); return result;
  };
  const lines = text.split("\n").filter(l=>l.trim());
  const headers = parseCSVLine(lines[0]);
  return lines.slice(1).map(line => {
    const vals=parseCSVLine(line); const obj={};
    headers.forEach((h,i)=>{ obj[h]=(vals[i]??"").replace(/^"|"$/g,"").trim(); });
    return obj;
  }).filter(r=>Object.values(r).some(v=>v));
};

// ─── PROCESSA GERENCIAL ───────────────────────────────────────────────────────
const processGerencial = raw => {
  const keys = raw.length?Object.keys(raw[0]):[];
  const find = (...kws) => keys.find(k=>kws.some(kw=>k.toUpperCase().includes(kw.toUpperCase())))||null;
  const get  = (r,...kws) => { const k=find(...kws); return k?(r[k]??""):""; };
  return raw.map((r,i) => {
    const grupo  = get(r,"Grupo Cliente").toString().trim();
    const cr     = get(r,"CR","CONTRATO").toString().trim();
    if (!grupo||!cr) return null;
    const statusReal = get(r,"Status Real","Status").toString().trim()||"N/A";
    const mesVig     = get(r,"Mês Vigência","MES").toString().trim();
    const inicioNeg  = parseData(get(r,"Inicio Negociação","Início Negociação","INICIO NEG"));
    const dataAprov  = parseData(get(r,"Data de Aprovação","DATA APROVAÇÃO"));
    const devido     = get(r,"Devido","DEVIDO").toString().trim();
    const aplicado   = get(r,"Aplicado","APLICADO").toString().trim();
    const mesNum     = parseMesNum(mesVig);
    const mesAtual   = new Date().getMonth()+1;
    let status = statusReal;
    if (!dataAprov && inicioNeg && mesNum && mesNum>mesAtual) status="Em Negociação Antecipada";
    else if (!dataAprov && inicioNeg) status="Em Negociação";
    else if (!dataAprov && !inicioNeg && mesNum && mesNum>mesAtual) status="Aguardando Reajuste";
    const diasNeg = dataAprov
      ? diasEntre(inicioNeg,dataAprov)
      : (inicioNeg ? diasEntre(inicioNeg,new Date()) : null);
    return {
      _id:i, grupo, cr,
      descr:   get(r,"DESCRI").toString().trim(),
      tipo:    get(r,"Tipo de Negócio","Tipo").toString().trim()||"Reajuste",
      status, statusReal, mes:mesVig,
      farmer:  get(r,"Responsável Farmer","FARMER").toString().trim(),
      info:    get(r,"Informações").toString().trim(),
      devido, aplicado, inicioNeg, dataAprov, diasNeg,
      concluido: !!dataAprov,
      _edited:false,
    };
  }).filter(Boolean);
};

// ─── PROCESSA PRODUTIVIDADE ───────────────────────────────────────────────────
const matchPessoa = s => EQUIPE.find(p=>(s||"").toUpperCase().includes(p.toUpperCase()))||null;
const matchSin    = v => {
  const s=(v||"").toUpperCase();
  if(s.includes("VERDE")) return "Verde";
  if(s.includes("AMARELO")) return "Amarelo";
  if(s.includes("VERM")) return "Vermelho";
  return null;
};
const matchCatP = (ativ,tipo) => {
  const c=((ativ||"")+" "+(tipo||"")).toUpperCase();
  if(c.includes("REAJUSTE")||c.includes("NOTIFICAÇ")||c.includes("CARTA DE REAJUSTE")) return "Reajuste";
  if(c.includes("RENOVAÇ")) return "Renovação";
  if(c.includes("UP-SELLING")||c.includes("UP SELLING")||c.includes("UPSELLING")) return "Up Selling";
  if(c.includes("DEFESA")) return "Defesa de Território";
  if(c.includes("ALTERAÇ")&&c.includes("ESCOPO")) return "Alteração de Escopo";
  if(c.includes("BID")||c.includes("COTAÇ")) return "BID/Cotação";
  return "Outros";
};

const processProd = raw => {
  const keys = raw.length?Object.keys(raw[0]):[];
  const find = (...kws) => keys.find(k=>kws.some(kw=>k.toUpperCase().includes(kw.toUpperCase())))||null;
  const get  = (r,...kws) => { const k=find(...kws); return k?(r[k]??""):""; };
  return raw.map(r => {
    const respF = get(r,"Responsável Farmer","RESP. FARMER").toString().trim();
    const respH = get(r,"Responsável Hunter","RESP. HUNTER").toString().trim();
    const pessoa = matchPessoa(respF)||matchPessoa(respH);
    if (!pessoa) return null;
    const ativ  = get(r,"ATIVIDADE","ATIVIDADE1").toString().toUpperCase().trim();
    const tipo  = get(r,"Tipo de Negócio","TIPO DE PROPOSTA").toString().trim();
    const status= get(r,"Status","STATUS").toString().trim();
    const mes   = get(r,"Mês","MES").toString().trim();
    const semana= get(r,"Semana","SEMANA").toString().trim();
    const escopo= get(r,"Escopo Atuação","ESCOPO").toString().trim();
    const grupo = get(r,"Grupo Cliente","GRUPO CLIENTE").toString().trim();
    const nProp = get(r,"Nº Proposta","N PROPOSTA").toString().trim();
    const obs   = get(r,"OBS.","OBS").toString().trim();
    const valPleito = num(get(r,"Valor Final / Pleito","COM REAJUSTE","VALOR PROPOSTA"));
    const valAprov  = num(get(r,"APROVADO PELO CLIENTE (R$)"));
    return {
      pessoa, grupo, nProp, mes, semana, escopo, obs, ativ, tipo, status,
      valPleito, valAprov,
      pctAceito: num(get(r,"% Reajuste Aceito","APROVADO PELO CLIENTE (%)")),
      isRevisao: (get(r,"Nº da Revisão","Revisão")||"0").toString()!=="0",
      sinalizacao: matchSin(get(r,"Semáforo","Sinalização")),
      categoria: matchCatP(ativ,tipo),
    };
  }).filter(Boolean);
};

// ─── EXPORTAR ─────────────────────────────────────────────────────────────────
const exportExcel = data => {
  const rows = data.map(r=>({
    "Grupo Cliente":r.grupo,"CR":r.cr,"Descrição CR":r.descr,
    "Tipo de Negócio":r.tipo,"Mês Vigência":r.mes,
    "Responsável Farmer":r.farmer,"Status Real":r.status,
    "Devido":r.devido,"Aplicado":r.aplicado,
    "Início Negociação":r.inicioNeg?fmtData(r.inicioNeg):"",
    "Data Aprovação":r.dataAprov?fmtData(r.dataAprov):"",
    "Dias Negociação":r.diasNeg??"",
    "Informações":r.info,"Editado":r._edited?"Sim":"",
  }));
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Gerencial");
  XLSX.writeFile(wb,"Gerencial_Atualizado.xlsx");
};

// ─── INLINE SELECT ────────────────────────────────────────────────────────────
const InlineSelect = ({ value, options, colorMap, onChange }) => {
  const [open,setOpen] = useState(false);
  const [pos,setPos]   = useState({top:0,left:0});
  const ref = useRef(null);
  const c   = colorMap?.[value]||{dot:"#9CA3AF",text:"#6B7280",bg:"#F3F4F6"};

  const handleOpen = e => {
    e.stopPropagation();
    if (!open&&ref.current) {
      const r=ref.current.getBoundingClientRect();
      const menuH=options.length*42;
      const openUp=window.innerHeight-r.bottom<menuH+8;
      setPos({top:openUp?r.top-menuH-4:r.bottom+4,left:r.left});
    }
    setOpen(o=>!o);
  };

  return (
    <div style={{position:"relative",display:"inline-block"}}>
      <div ref={ref} onClick={handleOpen}
           style={{display:"inline-flex",alignItems:"center",gap:5,cursor:"pointer",
                   padding:"2px 8px",borderRadius:99,background:c.bg,
                   border:`1.5px solid ${open?c.dot:"transparent"}`,userSelect:"none"}}>
        {colorMap&&<div style={{width:6,height:6,borderRadius:99,background:c.dot,flexShrink:0}}/>}
        <span style={{fontSize:11,fontWeight:600,color:c.text,whiteSpace:"nowrap"}}>{value}</span>
        <span style={{fontSize:8,color:c.text,opacity:.6}}>▼</span>
      </div>
      {open&&(<>
        <div onClick={()=>setOpen(false)} style={{position:"fixed",inset:0,zIndex:500}}/>
        <div style={{position:"fixed",top:pos.top,left:pos.left,zIndex:501,
                     background:"#fff",borderRadius:10,boxShadow:"0 8px 24px rgba(0,0,0,.18)",
                     border:"1px solid #E2E8F0",minWidth:180,overflow:"hidden"}}>
          {options.map(opt=>{
            const oc=colorMap?.[opt]||{dot:"#9CA3AF",text:"#374151"};
            return(
              <div key={opt} onClick={e=>{e.stopPropagation();onChange(opt);setOpen(false);}}
                   style={{display:"flex",alignItems:"center",gap:8,padding:"9px 14px",
                            cursor:"pointer",background:opt===value?"#F8FAFC":"#fff"}}
                   onMouseOver={e=>e.currentTarget.style.background="#F1F5F9"}
                   onMouseOut={e=>e.currentTarget.style.background=opt===value?"#F8FAFC":"#fff"}>
                {colorMap&&<div style={{width:8,height:8,borderRadius:99,background:oc.dot}}/>}
                <span style={{fontSize:13,color:oc.text,fontWeight:opt===value?600:400}}>{opt}</span>
                {opt===value&&<span style={{marginLeft:"auto",fontSize:11,color:oc.dot}}>✓</span>}
              </div>
            );
          })}
        </div>
      </>)}
    </div>
  );
};

// ─── EXPORT BTN ───────────────────────────────────────────────────────────────
const ExportBtn = ({ n, onExport }) => {
  const [open,setOpen]=useState(false);
  return (
    <div style={{position:"fixed",bottom:24,right:24,zIndex:300,
                 display:"flex",flexDirection:"column",alignItems:"flex-end",gap:8}}>
      {open&&(
        <div style={{background:"#fff",borderRadius:12,padding:"14px 18px",
                     boxShadow:"0 8px 24px rgba(0,0,0,.18)",border:"1px solid #E2E8F0",
                     display:"flex",flexDirection:"column",gap:10,minWidth:200}}>
          <div style={{fontSize:12,color:"#64748B"}}>
            <span style={{fontWeight:700,color:"#F59E0B"}}>{n}</span> alteraç{n===1?"ão":"ões"}
          </div>
          <button onClick={onExport} style={{padding:"9px 16px",borderRadius:8,background:"#16A34A",
            border:"none",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>
            ⬇ Exportar Excel
          </button>
          <button onClick={()=>setOpen(false)} style={{padding:"6px",borderRadius:8,
            background:"#F1F5F9",border:"none",color:"#64748B",fontSize:12,cursor:"pointer"}}>
            Minimizar
          </button>
        </div>
      )}
      <button onClick={()=>setOpen(o=>!o)}
        style={{width:44,height:44,borderRadius:99,background:open?"#64748B":"#16A34A",
                border:"none",color:"#fff",fontSize:18,cursor:"pointer",
                boxShadow:"0 4px 16px rgba(0,0,0,.25)",display:"flex",
                alignItems:"center",justifyContent:"center",position:"relative"}}>
        {open?"✕":"⬇"}
        {!open&&<span style={{position:"absolute",top:-4,right:-4,background:"#F59E0B",
                               color:"#fff",fontSize:10,fontWeight:700,minWidth:16,height:16,
                               borderRadius:99,display:"flex",alignItems:"center",
                               justifyContent:"center",padding:"0 3px",border:"2px solid #fff"}}>{n}</span>}
      </button>
    </div>
  );
};

// ─── PILL ─────────────────────────────────────────────────────────────────────
const Pill = ({ label, active, onClick, color }) => (
  <button onClick={onClick} style={{
    padding:"3px 10px",borderRadius:99,border:"none",cursor:"pointer",
    fontFamily:"inherit",fontSize:11,fontWeight:500,transition:"all .15s",
    background:active?(color||DARK):"rgba(255,255,255,.08)",
    color:active?"#fff":"rgba(255,255,255,.55)",
    whiteSpace:"nowrap",
  }}>{label}</button>
);

// ─── FILTRO LATERAL (para Produtividade) ──────────────────────────────────────
const FiltroLateral = ({ data, filtros, setFiltros }) => {
  const mesesDisp  = ORDEM_MES.filter(m=>data.some(r=>r.mes.toLowerCase().includes(m.substring(0,3).toLowerCase())));
  const semanasDisp= [...new Set(data.map(r=>r.semana).filter(Boolean))].sort();
  const scopesDisp = [...new Set(data.map(r=>r.escopo).filter(Boolean))].sort();

  const toggle = (field, val) => setFiltros(prev=>{
    const s = new Set(prev[field]);
    s.has(val)?s.delete(val):s.add(val);
    return {...prev,[field]:s};
  });
  const clear = field => setFiltros(prev=>({...prev,[field]:new Set()}));

  const Section = ({ label, field, items, colorMap }) => (
    <div style={{marginBottom:16}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
        <span style={{fontSize:10,fontWeight:700,color:"rgba(255,255,255,.35)",
                      textTransform:"uppercase",letterSpacing:".08em"}}>{label}</span>
        {filtros[field].size>0&&(
          <button onClick={()=>clear(field)} style={{fontSize:9,color:"rgba(255,255,255,.3)",
            background:"none",border:"none",cursor:"pointer",padding:0}}>limpar</button>
        )}
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
        {items.map(item=>(
          <button key={item} onClick={()=>toggle(field,item)} style={{
            padding:"3px 9px",borderRadius:99,border:"none",cursor:"pointer",
            fontFamily:"inherit",fontSize:11,fontWeight:500,transition:"all .15s",
            background:filtros[field].has(item)?(colorMap?.[item]||"#60A5FA"):"rgba(255,255,255,.08)",
            color:filtros[field].has(item)?"#fff":"rgba(255,255,255,.5)",
            whiteSpace:"nowrap",
          }}>{item}</button>
        ))}
      </div>
    </div>
  );

  return (
    <div style={{padding:"16px 12px",borderTop:"1px solid rgba(255,255,255,.08)",
                 overflowY:"auto",flex:1}}>
      <div style={{fontSize:10,fontWeight:700,color:"rgba(255,255,255,.3)",
                   textTransform:"uppercase",letterSpacing:".1em",marginBottom:14}}>Filtros</div>
      <Section label="Analista" field="pessoa"
        items={EQUIPE} colorMap={COR_P}/>
      <Section label="Status" field="status"
        items={["Aprovado","Em Negociação","Recusado"]}/>
      <Section label="Mês" field="mes" items={mesesDisp}/>
      <Section label="Semana" field="semana" items={semanasDisp}/>
      <Section label="Escopo" field="escopo" items={scopesDisp}/>
      <Section label="Semáforo" field="sin"
        items={["Verde","Amarelo","Vermelho"]} colorMap={SIN_COR}/>
      <Section label="Tipo" field="cat" items={TIPO_P} colorMap={TIPO_COR_P}/>
    </div>
  );
};

// ─── GERENCIAL ────────────────────────────────────────────────────────────────
function SecaoGerencial({ data, setData }) {
  const [fMeses,  setFMeses]  = useState(new Set());
  const [fFarmer, setFFarmer] = useState("Todos");
  const [fStatus, setFStatus] = useState(new Set());
  const [fTipo,   setFTipo]   = useState("Todos");
  const [busca,   setBusca]   = useState("");
  const [tooltip, setTooltip] = useState({visible:false,text:"",x:0,y:0});

  const editRow = (id,field,value) =>
    setData(prev=>prev.map(r=>r._id===id?{...r,[field]:value,_edited:true}:r));

  const mesesDisp  = ORDEM_MES.filter(m=>data.some(r=>r.mes===m));
  const todosMeses = fMeses.size===0;
  const toggleMes  = m=>setFMeses(prev=>{const n=new Set(prev);n.has(m)?n.delete(m):n.add(m);return n;});
  const todosStat  = fStatus.size===0;

  // Farmers: só nomes reais da equipe
  const farmersDisp = ["Todos",...EQUIPE.filter(p=>data.some(r=>r.farmer.toUpperCase().includes(p.toUpperCase())))];

  const filtrado = data.filter(r => {
    if (!todosMeses&&!fMeses.has(r.mes)) return false;
    if (fFarmer!=="Todos"&&!r.farmer.toUpperCase().includes(fFarmer.toUpperCase())) return false;
    if (!todosStat&&!fStatus.has(r.status)&&!fStatus.has(r.statusReal)) return false;
    if (fTipo!=="Todos"&&r.tipo!==fTipo) return false;
    if (busca&&![r.grupo,r.cr,r.descr].some(v=>v.toUpperCase().includes(busca.toUpperCase()))) return false;
    return true;
  });

  const cardBase = todosMeses?data:data.filter(r=>fMeses.has(r.mes));
  const buildCard = rows => ({
    total:rows.length,
    por:[
      {s:"Concluído",      n:rows.filter(r=>r.status==="Concluído"||r.statusReal==="Concluído").length},
      {s:"Pendente Cliente",n:rows.filter(r=>r.status==="Pendente Cliente").length},
      {s:"Pendente GPS",   n:rows.filter(r=>r.status==="Pendente GPS").length},
      {s:"Em Negociação",  n:rows.filter(r=>r.status==="Em Negociação"||r.status==="Em Negociação Antecipada").length},
      {s:"Negócio Perdido",n:rows.filter(r=>r.status==="Negócio Perdido"||r.statusReal==="Negócio Perdido").length},
      {s:"N/A",            n:rows.filter(r=>r.status==="N/A").length},
    ].filter(x=>x.n>0)
  });
  const cardGeral = buildCard(cardBase);
  const cardsTipo = TIPO_G.map(t=>({tipo:t,...buildCard(cardBase.filter(r=>r.tipo===t))})).filter(c=>c.total>0);

  const grupos = (()=>{
    const by={};
    filtrado.forEach(r=>{if(!by[r.grupo])by[r.grupo]=[];by[r.grupo].push(r);});
    return Object.entries(by).map(([g,crs])=>({g,crs})).sort((a,b)=>a.g.localeCompare(b.g));
  })();

  const nEditados = data.filter(r=>r._edited).length;
  const showTip=(e,t)=>{if(t)setTooltip({visible:true,text:t,x:e.clientX,y:e.clientY});};
  const moveTip=e=>{if(tooltip.visible)setTooltip(p=>({...p,x:e.clientX,y:e.clientY}));};
  const hideTip=()=>setTooltip(p=>({...p,visible:false}));

  return (
    <div>
      {tooltip.visible&&tooltip.text&&(
        <div style={{position:"fixed",zIndex:9999,pointerEvents:"none",
          left:tooltip.x+14,top:tooltip.y-10,background:DARK,color:"#F8FAFC",
          padding:"8px 12px",borderRadius:8,fontSize:12,maxWidth:300,lineHeight:1.5,
          boxShadow:"0 4px 16px rgba(0,0,0,.3)",whiteSpace:"pre-wrap"}}>{tooltip.text}</div>
      )}

      {/* FILTROS */}
      <div style={{background:"#fff",borderRadius:10,padding:"10px 14px",marginBottom:16,
                   display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",
                   boxShadow:"0 1px 3px rgba(0,0,0,.06)"}}>
        <input value={busca} onChange={e=>setBusca(e.target.value)}
          placeholder="Buscar..." style={{padding:"5px 10px",borderRadius:8,
          border:"1px solid #E2E8F0",fontSize:12,color:DARK,outline:"none",
          width:160,fontFamily:"inherit"}}/>
        <div style={{width:1,height:20,background:"#E2E8F0"}}/>
        <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
          <span style={{fontSize:10,fontWeight:600,color:"#94A3B8"}}>MÊS</span>
          <button onClick={()=>setFMeses(new Set())} style={{padding:"3px 10px",borderRadius:99,
            border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:11,fontWeight:500,
            background:todosMeses?DARK:"#F1F4F8",color:todosMeses?"#fff":"#64748B"}}>Todos</button>
          {mesesDisp.map(m=>(
            <button key={m} onClick={()=>toggleMes(m)} style={{padding:"3px 10px",borderRadius:99,
              border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:11,fontWeight:500,
              background:fMeses.has(m)?DARK:"#F1F4F8",color:fMeses.has(m)?"#fff":"#64748B"}}>{m}</button>
          ))}
        </div>
        <div style={{width:1,height:20,background:"#E2E8F0"}}/>
        <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
          <span style={{fontSize:10,fontWeight:600,color:"#94A3B8"}}>FARMER</span>
          {farmersDisp.map(f=>(
            <button key={f} onClick={()=>setFFarmer(f)} style={{padding:"3px 10px",borderRadius:99,
              border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:11,fontWeight:500,
              background:fFarmer===f?(COR_P[f]||DARK):"#F1F4F8",
              color:fFarmer===f?"#fff":"#64748B"}}>{f}</button>
          ))}
        </div>
        <div style={{width:1,height:20,background:"#E2E8F0"}}/>
        <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
          <span style={{fontSize:10,fontWeight:600,color:"#94A3B8"}}>STATUS</span>
          <button onClick={()=>setFStatus(new Set())} style={{padding:"3px 10px",borderRadius:99,
            border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:11,fontWeight:500,
            background:todosStat?DARK:"#F1F4F8",color:todosStat?"#fff":"#64748B"}}>Todos</button>
          {["Concluído","Pendente Cliente","Pendente GPS","Em Negociação","Negócio Perdido","N/A"].map(s=>(
            <button key={s} onClick={()=>setFStatus(prev=>{const n=new Set(prev);n.has(s)?n.delete(s):n.add(s);return n;})}
              style={{padding:"3px 10px",borderRadius:99,border:"none",cursor:"pointer",
                fontFamily:"inherit",fontSize:11,fontWeight:500,
                background:fStatus.has(s)?(STATUS_COR[s]?.dot||DARK):"#F1F4F8",
                color:fStatus.has(s)?"#fff":"#64748B"}}>{s}</button>
          ))}
        </div>
        <div style={{width:1,height:20,background:"#E2E8F0"}}/>
        <div style={{display:"flex",gap:4,alignItems:"center",flexWrap:"wrap"}}>
          <span style={{fontSize:10,fontWeight:600,color:"#94A3B8"}}>TIPO</span>
          {["Todos",...TIPO_G].map(t=>(
            <button key={t} onClick={()=>setFTipo(t)} style={{padding:"3px 10px",borderRadius:99,
              border:"none",cursor:"pointer",fontFamily:"inherit",fontSize:11,fontWeight:500,
              background:fTipo===t?DARK:"#F1F4F8",color:fTipo===t?"#fff":"#64748B"}}>{t}</button>
          ))}
        </div>
        <span style={{marginLeft:"auto",fontSize:11,color:"#94A3B8",whiteSpace:"nowrap"}}>
          {filtrado.length} CRs · {grupos.length} grupos
        </span>
      </div>

      {/* CARDS */}
      <div style={{display:"grid",gridTemplateColumns:`repeat(${1+cardsTipo.length},1fr)`,
                   gap:12,marginBottom:20}}>
        <div style={{background:DARK,borderRadius:12,padding:"16px 18px",boxShadow:"0 4px 14px rgba(26,35,50,.25)"}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
            <div>
              <div style={{fontSize:10,color:"rgba(255,255,255,.5)",marginBottom:3}}>Total Geral</div>
              <div style={{fontSize:32,fontWeight:800,color:"#fff",lineHeight:1}}>{cardGeral.total}</div>
            </div>
            <div style={{fontSize:24,opacity:.3}}>📋</div>
          </div>
          {cardGeral.por.map(({s,n})=>{
            const c=STATUS_COR[s]||{dot:"#9CA3AF",text:"rgba(255,255,255,.6)"};
            const p=cardGeral.total>0?Math.round(n/cardGeral.total*100):0;
            return(
              <div key={s} style={{display:"flex",alignItems:"center",gap:6,marginBottom:5}}>
                <div style={{width:7,height:7,borderRadius:99,background:c.dot,flexShrink:0}}/>
                <span style={{fontSize:11,color:"rgba(255,255,255,.6)",flex:1}}>{s}</span>
                <span style={{fontSize:12,fontWeight:700,color:"#fff",minWidth:20,textAlign:"right"}}>{n}</span>
                <span style={{fontSize:10,color:"rgba(255,255,255,.3)",minWidth:32,textAlign:"right"}}>({p}%)</span>
              </div>
            );
          })}
        </div>
        {cardsTipo.map(({tipo,total,por})=>{
          const conc=por.find(p=>p.s==="Concluído")?.n||0;
          return(
            <div key={tipo} style={{background:"#fff",borderRadius:12,padding:"16px 18px",
                                    borderLeft:`4px solid ${DARK}`,boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
                <div>
                  <div style={{fontSize:10,color:"#94A3B8",marginBottom:3}}>{tipo}</div>
                  <div style={{fontSize:32,fontWeight:800,color:DARK,lineHeight:1}}>{total}</div>
                  <div style={{fontSize:10,color:"#94A3B8",marginTop:3}}>
                    {total>0?Math.round(conc/total*100):0}% concluído
                  </div>
                </div>
                <div style={{background:"#F1F4F8",borderRadius:8,padding:"5px 9px",
                              fontSize:11,fontWeight:700,color:DARK,height:"fit-content"}}>{tipo}</div>
              </div>
              {por.map(({s,n})=>{
                const c=STATUS_COR[s]||{dot:"#9CA3AF"};
                const p=total>0?Math.round(n/total*100):0;
                return(
                  <div key={s} style={{display:"flex",alignItems:"center",gap:6,marginBottom:5}}>
                    <div style={{width:7,height:7,borderRadius:99,background:c.dot,flexShrink:0}}/>
                    <span style={{fontSize:11,color:"#64748B",flex:1}}>{s}</span>
                    <span style={{fontSize:12,fontWeight:700,color:DARK,minWidth:20,textAlign:"right"}}>{n}</span>
                    <span style={{fontSize:10,color:"#94A3B8",minWidth:32,textAlign:"right"}}>({p}%)</span>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>

      {/* TABELA */}
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {grupos.length===0&&(
          <div style={{background:"#fff",borderRadius:12,padding:"48px 0",
                       textAlign:"center",color:"#94A3B8",fontSize:14}}>Nenhum CR encontrado.</div>
        )}
        {grupos.map(({g,crs})=>{
          const nEdit=crs.filter(r=>r._edited).length;
          return(
            <div key={g} style={{background:"#fff",borderRadius:12,overflow:"hidden",
                                  boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
              <div style={{background:DARK2,padding:"10px 16px",display:"flex",
                            alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:13,fontWeight:700,color:"#fff"}}>{g}</span>
                  <span style={{fontSize:10,color:"rgba(255,255,255,.4)"}}>{crs.length} CR{crs.length!==1?"s":""}</span>
                  {nEdit>0&&<span style={{fontSize:9,background:"#F59E0B",color:"#fff",
                    padding:"1px 6px",borderRadius:99,fontWeight:600}}>{nEdit} editado{nEdit!==1?"s":""}</span>}
                </div>
                <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                  {["Concluído","Pendente Cliente","Pendente GPS","Em Negociação","Negócio Perdido"].map(s=>{
                    const n=crs.filter(c=>c.status===s||c.statusReal===s).length;
                    if(!n) return null;
                    const sc=STATUS_COR[s];
                    return(
                      <span key={s} style={{fontSize:10,fontWeight:600,padding:"1px 8px",
                        borderRadius:99,background:sc.bg,color:sc.text}}>{n} {s}</span>
                    );
                  })}
                </div>
              </div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
                  <colgroup>
                    <col style={{width:80}}/><col style={{width:"auto"}}/><col style={{width:60}}/>
                    <col style={{width:80}}/><col style={{width:65}}/><col style={{width:65}}/>
                    <col style={{width:80}}/><col style={{width:110}}/><col style={{width:110}}/>
                  </colgroup>
                  <thead>
                    <tr style={{background:"#F8FAFC"}}>
                      {["CR","Descrição","Mês","Farmer","Devido","Aplicado","Tempo Neg.","Tipo","Status"].map((h,i)=>(
                        <th key={h} style={{textAlign:i>2?"center":"left",padding:"7px 12px",
                          fontSize:10,color:"#94A3B8",fontWeight:600,
                          borderBottom:"1px solid #E2E8F0",whiteSpace:"nowrap"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {crs.map((cr,ci)=>(
                      <tr key={cr._id}
                          style={{background:cr._edited?"#FFFBEB":ci%2===0?"#fff":"#FAFBFC",
                                  borderTop:"1px solid #F1F4F8"}}
                          onMouseOver={e=>{if(!cr._edited)e.currentTarget.style.background="#EFF6FF";}}
                          onMouseOut={e=>{e.currentTarget.style.background=cr._edited?"#FFFBEB":ci%2===0?"#fff":"#FAFBFC";}}>
                        <td style={{padding:"7px 12px",fontSize:11,color:"#64748B",fontFamily:"monospace",
                                    overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                          {cr._edited&&<span style={{color:"#F59E0B",marginRight:3}}>●</span>}
                          {cr.cr}
                        </td>
                        <td style={{padding:"7px 12px",fontSize:11,color:DARK,
                                    cursor:cr.info?"help":"default",overflow:"hidden"}}
                            onMouseEnter={e=>showTip(e,cr.info)}
                            onMouseMove={moveTip} onMouseLeave={hideTip}>
                          <span style={{display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                            {cr.descr}
                          </span>
                        </td>
                        <td style={{padding:"7px 12px",fontSize:11,color:"#64748B",textAlign:"center",whiteSpace:"nowrap"}}>{cr.mes}</td>
                        <td style={{padding:"7px 12px",fontSize:11,color:"#64748B",textAlign:"center",
                                    overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{cr.farmer||"—"}</td>
                        <td style={{padding:"7px 12px",fontSize:11,textAlign:"center",whiteSpace:"nowrap",
                                    color:cr.devido&&cr.devido!=="—"?"#7C3AED":"#94A3B8",
                                    fontWeight:cr.devido&&cr.devido!=="—"?600:400}}>
                          {pctFmt(cr.devido)}
                        </td>
                        <td style={{padding:"7px 12px",fontSize:11,textAlign:"center",whiteSpace:"nowrap",
                                    color:cr.aplicado&&cr.aplicado!=="—"?"#059669":"#94A3B8",
                                    fontWeight:cr.aplicado&&cr.aplicado!=="—"?600:400}}>
                          {pctFmt(cr.aplicado)}
                        </td>
                        <td style={{padding:"7px 12px",textAlign:"center",whiteSpace:"nowrap"}}>
                          {cr.diasNeg!=null?(
                            <span style={{fontWeight:600,fontSize:11,
                              color:cr.concluido?"#16A34A":"#D97706",
                              background:cr.concluido?"#DCFCE7":"#FEF9C3",
                              padding:"2px 7px",borderRadius:99}}>
                              {cr.diasNeg}d {cr.concluido?"✓":"⏳"}
                            </span>
                          ):"—"}
                        </td>
                        <td style={{padding:"5px 12px",textAlign:"center"}}>
                          <InlineSelect value={cr.tipo} options={TIPO_G} onChange={v=>editRow(cr._id,"tipo",v)}/>
                        </td>
                        <td style={{padding:"5px 12px",textAlign:"center"}}>
                          <InlineSelect value={cr.status} options={STATUS_G} colorMap={STATUS_COR}
                            onChange={v=>editRow(cr._id,"status",v)}/>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          );
        })}
      </div>
      {nEditados>0&&<ExportBtn n={nEditados} onExport={()=>exportExcel(data)}/>}
    </div>
  );
}

// ─── PRODUTIVIDADE ────────────────────────────────────────────────────────────
function SecaoProdutividade({ data, subPag, filtros }) {
  const filt = data.filter(r => {
    if (filtros.pessoa.size>0&&!filtros.pessoa.has(r.pessoa)) return false;
    if (filtros.mes.size>0&&![...filtros.mes].some(m=>r.mes.toLowerCase().includes(m.substring(0,3).toLowerCase()))) return false;
    if (filtros.semana.size>0&&!filtros.semana.has(r.semana)) return false;
    if (filtros.escopo.size>0&&!filtros.escopo.has(r.escopo)) return false;
    if (filtros.sin.size>0&&!filtros.sin.has(r.sinalizacao)) return false;
    if (filtros.cat.size>0&&!filtros.cat.has(r.categoria)) return false;
    if (filtros.status.size>0) {
      const isAprov = r.status.toUpperCase().includes("APROVADO");
      const isNeg   = r.status.toUpperCase().includes("NEGOCI");
      const isRecus = r.status.toUpperCase().includes("RECUS");
      if (filtros.status.has("Aprovado")&&!isAprov) return false;
      if (filtros.status.has("Em Negociação")&&!isNeg) return false;
      if (filtros.status.has("Recusado")&&!isRecus) return false;
    }
    return true;
  });

  // ── VISÃO GERAL ─────────────────────────────────────────────────────────────
  if (subPag==="visao") {
    const somaAprov  = filt.reduce((s,r)=>s+r.valAprov,0);
    const somaPleito = filt.reduce((s,r)=>s+r.valPleito,0);
    const proprias   = filt.filter(r=>!r.isRevisao);
    const aprovadas  = filt.filter(r=>r.status.toUpperCase().includes("APROVADO"));
    const txAprov    = proprias.length>0?Math.round(aprovadas.length/proprias.length*100):0;
    return (
      <div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:24}}>
          {[
            {l:"Propostas",v:proprias.length,icon:"📋"},
            {l:"Valor em negociação",v:brl(somaPleito),icon:"🎯"},
            {l:"Valor aprovado",v:brl(somaAprov),icon:"✅",hl:"#16A34A"},
            {l:"Taxa de aprovação",v:`${txAprov}%`,icon:"📈",hl:txAprov>50?"#16A34A":"#D97706"},
          ].map(k=>(
            <div key={k.l} style={{background:"#fff",borderRadius:12,padding:"16px 18px",
                                    boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
              <div style={{fontSize:20,marginBottom:8}}>{k.icon}</div>
              <div style={{fontSize:11,color:"#94A3B8",marginBottom:4}}>{k.l}</div>
              <div style={{fontSize:20,fontWeight:700,color:k.hl||DARK}}>{k.v}</div>
            </div>
          ))}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
          {EQUIPE.filter(p=>filt.some(r=>r.pessoa===p)).map(p=>{
            const rows=filt.filter(r=>r.pessoa===p);
            const propP=rows.filter(r=>!r.isRevisao);
            const aprovP=rows.filter(r=>r.status.toUpperCase().includes("APROVADO"));
            const porCat=TIPO_P.map(cat=>({cat,n:rows.filter(r=>r.categoria===cat).length})).filter(x=>x.n>0);
            return(
              <div key={p} style={{background:"#fff",borderRadius:14,padding:20,
                                    borderTop:`3px solid ${COR_P[p]}`,
                                    boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:14}}>
                  <div>
                    <div style={{fontSize:11,color:"#94A3B8",marginBottom:3}}>Analista</div>
                    <div style={{fontSize:17,fontWeight:700,color:COR_P[p]}}>{p}</div>
                    <div style={{fontSize:11,color:"#94A3B8",marginTop:2}}>
                      {propP.length} proposta{propP.length!==1?"s":""} · {rows.filter(r=>r.isRevisao).length} revisões
                    </div>
                  </div>
                  <div style={{background:"#DCFCE7",borderRadius:8,padding:"4px 10px",textAlign:"center"}}>
                    <div style={{fontSize:17,fontWeight:700,color:"#16A34A"}}>{aprovP.length}</div>
                    <div style={{fontSize:10,color:"#16A34A"}}>aprovados</div>
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
                  {[
                    {l:"Valor pleito",v:brl(rows.reduce((s,r)=>s+r.valPleito,0))},
                    {l:"Valor aprovado",v:brl(rows.reduce((s,r)=>s+r.valAprov,0)),hl:true},
                  ].map(k=>(
                    <div key={k.l} style={{background:"#F8FAFC",borderRadius:8,padding:"8px 10px"}}>
                      <div style={{fontSize:10,color:"#94A3B8",marginBottom:2}}>{k.l}</div>
                      <div style={{fontSize:12,fontWeight:600,color:k.hl?"#16A34A":DARK}}>{k.v}</div>
                    </div>
                  ))}
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                  {porCat.map(({cat,n})=>(
                    <span key={cat} style={{fontSize:10,fontWeight:600,padding:"2px 7px",borderRadius:99,
                      background:`${TIPO_COR_P[cat]||"#6B7280"}18`,color:TIPO_COR_P[cat]||"#6B7280",
                      border:`1px solid ${TIPO_COR_P[cat]||"#6B7280"}30`}}>{cat} ({n})</span>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ── POR CLIENTE ──────────────────────────────────────────────────────────────
  if (subPag==="clientes") {
    const grupos = (()=>{
      const by={};
      filt.forEach(r=>{
        if(!by[r.grupo])by[r.grupo]={rows:[],sin:null};
        by[r.grupo].rows.push(r);
        if(!by[r.grupo].sin&&r.sinalizacao)by[r.grupo].sin=r.sinalizacao;
      });
      return Object.entries(by).map(([g,{rows,sin}])=>({g,rows,sin})).sort((a,b)=>a.g.localeCompare(b.g));
    })();
    return (
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {grupos.map(({g,rows,sin})=>{
          const aprov=rows.filter(r=>r.status.toUpperCase().includes("APROVADO")).length;
          return(
            <div key={g} style={{background:"#fff",borderRadius:12,overflow:"hidden",
                                  boxShadow:"0 1px 4px rgba(0,0,0,.07)",
                                  borderLeft:`4px solid ${sin?SIN_COR[sin]:"#E2E8F0"}`}}>
              <div style={{padding:"12px 18px",display:"flex",alignItems:"center",
                            justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <span style={{fontSize:14,fontWeight:700,color:DARK}}>{g}</span>
                  {sin&&<span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:99,
                    background:SIN_BG[sin],color:SIN_COR[sin]}}>{sin}</span>}
                  <span style={{fontSize:11,color:"#94A3B8"}}>{rows.length} interações</span>
                </div>
                <div style={{display:"flex",gap:20}}>
                  {[
                    {l:"Valor pleito",v:brl(rows.reduce((s,r)=>s+r.valPleito,0))},
                    {l:"Valor aprovado",v:brl(rows.reduce((s,r)=>s+r.valAprov,0)),c:"#16A34A"},
                    {l:"Aprovações",v:`${aprov}/${rows.length}`},
                  ].map(k=>(
                    <div key={k.l} style={{textAlign:"right"}}>
                      <div style={{fontSize:10,color:"#94A3B8",marginBottom:2}}>{k.l}</div>
                      <div style={{fontSize:13,fontWeight:600,color:k.c||DARK}}>{k.v}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          );
        })}
        {!grupos.length&&<div style={{background:"#fff",borderRadius:12,padding:"48px 0",
          textAlign:"center",color:"#94A3B8",fontSize:14}}>Nenhum cliente encontrado.</div>}
      </div>
    );
  }

  // ── HISTÓRICO ────────────────────────────────────────────────────────────────
  return (
    <div style={{background:"#fff",borderRadius:12,overflow:"hidden",
                 boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
          <thead>
            <tr style={{background:"#F8FAFC"}}>
              {["Resp.","Mês","Semana","Cliente","Proposta","Tipo","Escopo","Val. Pleito","Val. Aprovado","Status","Rev."].map(h=>(
                <th key={h} style={{textAlign:"left",padding:"8px 12px",fontSize:10,
                  color:"#94A3B8",fontWeight:600,whiteSpace:"nowrap",borderBottom:"1px solid #E2E8F0"}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filt.map((r,i)=>{
              const isAprov=r.status.toUpperCase().includes("APROVADO");
              const isRecus=r.status.toUpperCase().includes("RECUS");
              const sc=isAprov?{text:"#16A34A",bg:"#DCFCE7"}:isRecus?{text:"#DC2626",bg:"#FEE2E2"}:{text:"#D97706",bg:"#FEF9C3"};
              return(
                <tr key={i} style={{background:i%2===0?"#fff":"#F8FAFC",borderTop:"1px solid #F1F4F8"}}>
                  <td style={{padding:"7px 12px",borderLeft:`3px solid ${COR_P[r.pessoa]||"#E2E8F0"}`,
                    color:COR_P[r.pessoa],fontWeight:600,whiteSpace:"nowrap"}}>{r.pessoa}</td>
                  <td style={{padding:"7px 12px",color:"#94A3B8",whiteSpace:"nowrap"}}>{r.mes}</td>
                  <td style={{padding:"7px 12px",color:"#94A3B8",whiteSpace:"nowrap",fontSize:10}}>{r.semana}</td>
                  <td style={{padding:"7px 12px",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:DARK}}>{r.grupo}</td>
                  <td style={{padding:"7px 12px",color:"#94A3B8",whiteSpace:"nowrap",fontFamily:"monospace",fontSize:10}}>{r.nProp}</td>
                  <td style={{padding:"7px 12px",whiteSpace:"nowrap"}}>
                    <span style={{fontSize:10,fontWeight:600,padding:"2px 7px",borderRadius:99,
                      background:`${TIPO_COR_P[r.categoria]||"#6B7280"}18`,
                      color:TIPO_COR_P[r.categoria]||"#6B7280"}}>{r.categoria}</span>
                  </td>
                  <td style={{padding:"7px 12px",color:"#94A3B8",whiteSpace:"nowrap",fontSize:10}}>{r.escopo||"—"}</td>
                  <td style={{padding:"7px 12px",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.valPleito>0?brl(r.valPleito):"—"}</td>
                  <td style={{padding:"7px 12px",color:"#16A34A",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.valAprov>0?brl(r.valAprov):"—"}</td>
                  <td style={{padding:"7px 12px"}}>
                    <span style={{fontSize:10,fontWeight:500,padding:"2px 7px",borderRadius:99,
                      background:sc.bg,color:sc.text,whiteSpace:"nowrap"}}>{r.status}</span>
                  </td>
                  <td style={{padding:"7px 12px",textAlign:"center",
                    color:r.isRevisao?"#DC2626":"#94A3B8",fontSize:10}}>{r.isRevisao?"Rev":"—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {!filt.length&&<div style={{textAlign:"center",padding:"40px 0",color:"#94A3B8",fontSize:14}}>
          Nenhum registro encontrado.</div>}
      </div>
    </div>
  );
}

// ─── APP PRINCIPAL ────────────────────────────────────────────────────────────
const FILTROS_INIT = () => ({
  pessoa:new Set(),mes:new Set(),semana:new Set(),
  escopo:new Set(),sin:new Set(),cat:new Set(),status:new Set()
});

export default function App() {
  const [gData,   setGData]   = useState([]);
  const [pData,   setPData]   = useState([]);
  const [loading, setLoading] = useState(true);
  const [erro,    setErro]    = useState(null);
  const [ultimaAt,setUltimaAt]= useState(null);
  const [secao,   setSecao]   = useState("gerencial");
  const [subPag,  setSubPag]  = useState("visao");
  const [menuOpen,setMenuOpen]= useState(false);
  const [filtroP, setFiltroP] = useState(FILTROS_INIT);

  const buscarDados = useCallback(async()=>{
    setLoading(true);setErro(null);
    try {
      const [rawG,rawP] = await Promise.all([
        fetchCSV("2073814116"),
        fetchCSV("1622380363"),
      ]);
      setGData(processGerencial(rawG));
      setPData(processProd(rawP));
      setUltimaAt(new Date());
    } catch(e) { setErro(e.message); }
    finally { setLoading(false); }
  },[]);

  useEffect(()=>{ buscarDados(); },[buscarDados]);

  const navItems = [
    { id:"gerencial",     label:"Gerencial",     subs:[] },
    { id:"produtividade", label:"Produtividade",  subs:[
      {id:"visao",    label:"Visão Geral"},
      {id:"clientes", label:"Por Cliente"},
      {id:"historico",label:"Histórico"},
    ]},
  ];

  const navTitle = secao==="gerencial"?"Gerencial"
    : navItems.find(n=>n.id==="produtividade")?.subs.find(s=>s.id===subPag)?.label||"Produtividade";

  return (
    <div style={{display:"flex",minHeight:"100vh",background:"#F1F4F8",
                 fontFamily:"system-ui,-apple-system,sans-serif",color:DARK}}>

      {/* Overlay quando menu aberto */}
      {menuOpen&&(
        <div onClick={()=>setMenuOpen(false)}
             style={{position:"fixed",inset:0,zIndex:90,background:"rgba(0,0,0,.3)"}}/>
      )}

      {/* ── MENU LATERAL ────────────────────────────────────────────────────── */}
      <div style={{
        width: menuOpen?MENU_W_OPEN:MENU_W_CLOSE,
        minHeight:"100vh",background:MENU_BG,
        display:"flex",flexDirection:"column",
        position:"fixed",top:0,left:0,zIndex:100,
        transition:"width .2s",overflow:"hidden",
        boxShadow: menuOpen?"4px 0 20px rgba(0,0,0,.3)":"none",
      }}>
        {/* Logo + toggle */}
        <div style={{padding:menuOpen?"18px 16px 14px":"14px 0",display:"flex",
                     alignItems:"center",justifyContent:menuOpen?"space-between":"center",
                     borderBottom:"1px solid rgba(255,255,255,.07)",minHeight:60}}>
          {menuOpen&&(
            <div style={{lineHeight:1}}>
              <div style={{fontSize:9,fontWeight:400,color:"rgba(255,255,255,.35)",letterSpacing:".14em"}}>GRUPO</div>
              <div style={{fontSize:24,fontWeight:800,color:"#fff",letterSpacing:"-.01em"}}>GPS</div>
            </div>
          )}
          <button onClick={()=>setMenuOpen(o=>!o)} style={{
            width:28,height:28,borderRadius:6,border:"none",
            background:"rgba(255,255,255,.07)",color:"rgba(255,255,255,.5)",
            cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",
            fontSize:12,flexShrink:0,
          }}>
            {menuOpen?"◀":"▶"}
          </button>
        </div>

        {/* Nav */}
        <nav style={{padding:"8px 0"}}>
          {navItems.map(({id,label,subs})=>(
            <div key={id}>
              <div onClick={()=>{setSecao(id);setSubPag(subs[0]?.id||"visao");setMenuOpen(true);}}
                   style={{
                     padding:menuOpen?"9px 16px":"9px 0",
                     display:"flex",alignItems:"center",gap:8,cursor:"pointer",
                     justifyContent:menuOpen?"flex-start":"center",
                     background:secao===id?"rgba(255,255,255,.07)":"transparent",
                     borderLeft:secao===id?"3px solid #60A5FA":"3px solid transparent",
                     transition:"all .15s",
                   }}>
                {!menuOpen&&(
                  <span style={{fontSize:11,fontWeight:700,width:MENU_W_CLOSE,textAlign:"center",
                    color:secao===id?"#60A5FA":"rgba(255,255,255,.4)"}}>
                    {label.charAt(0)}
                  </span>
                )}
                {menuOpen&&(
                  <span style={{fontSize:13,fontWeight:600,
                    color:secao===id?"#fff":"rgba(255,255,255,.55)"}}>
                    {label}
                  </span>
                )}
              </div>
              {menuOpen&&secao===id&&subs.length>0&&subs.map(sub=>(
                <div key={sub.id} onClick={()=>setSubPag(sub.id)}
                     style={{padding:"6px 16px 6px 28px",cursor:"pointer",fontSize:12,
                       color:subPag===sub.id?"#60A5FA":"rgba(255,255,255,.4)",
                       fontWeight:subPag===sub.id?600:400,
                       background:subPag===sub.id?"rgba(96,165,250,.07)":"transparent",
                       transition:"all .15s",
                     }}>
                  {sub.label}
                </div>
              ))}
            </div>
          ))}
        </nav>

        {/* Filtros Produtividade no menu */}
        {menuOpen&&secao==="produtividade"&&(
          <FiltroLateral data={pData} filtros={filtroP} setFiltros={setFiltroP}/>
        )}

        {/* Última atualização */}
        {menuOpen&&ultimaAt&&(
          <div style={{padding:"10px 16px",borderTop:"1px solid rgba(255,255,255,.06)",
                       fontSize:10,color:"rgba(255,255,255,.25)",marginTop:"auto"}}>
            Atualizado às {ultimaAt.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}
          </div>
        )}
      </div>

      {/* ── CONTEÚDO ────────────────────────────────────────────────────────── */}
      <div style={{flex:1,marginLeft:MENU_W_CLOSE,display:"flex",flexDirection:"column",
                   minWidth:0,transition:"margin-left .2s"}}>
        {/* Header */}
        <div style={{background:DARK,padding:"0 24px",display:"flex",alignItems:"center",
                     justifyContent:"space-between",minHeight:54,gap:12,flexShrink:0}}>
          <div>
            <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>{navTitle}</div>
            <div style={{fontSize:10,color:"rgba(255,255,255,.35)"}}>Grupo GPS · Controle Comercial</div>
          </div>
          <button onClick={buscarDados} disabled={loading} style={{
            padding:"6px 14px",borderRadius:8,background:"rgba(255,255,255,.1)",
            border:"1px solid rgba(255,255,255,.15)",color:"#fff",fontSize:11,
            fontWeight:600,cursor:loading?"not-allowed":"pointer",
            display:"flex",alignItems:"center",gap:5,opacity:loading?.7:1,flexShrink:0,
          }}>
            {loading?"⏳":"⟳"} {loading?"Atualizando...":"Atualizar"}
          </button>
        </div>

        {/* Conteúdo */}
        <div style={{flex:1,padding:"20px 24px",overflowY:"auto",overflowX:"hidden"}}>
          {loading&&(
            <div style={{background:"#fff",borderRadius:14,padding:"80px 0",textAlign:"center",
                         boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
              <div style={{fontSize:32,marginBottom:12}}>⏳</div>
              <div style={{fontSize:15,fontWeight:600,color:DARK,marginBottom:6}}>Carregando dados...</div>
              <div style={{fontSize:13,color:"#94A3B8"}}>Buscando planilha do Google Sheets</div>
            </div>
          )}
          {erro&&!loading&&(
            <div style={{background:"#FEF2F2",borderRadius:14,padding:"40px",
                         textAlign:"center",border:"1px solid #FECACA"}}>
              <div style={{fontSize:32,marginBottom:12}}>⚠️</div>
              <div style={{fontSize:15,fontWeight:600,color:"#DC2626",marginBottom:8}}>
                Não foi possível carregar os dados
              </div>
              <div style={{fontSize:13,color:"#EF4444",marginBottom:20}}>{erro}</div>
              <button onClick={buscarDados} style={{padding:"10px 24px",borderRadius:8,
                background:"#DC2626",border:"none",color:"#fff",fontSize:13,
                fontWeight:600,cursor:"pointer"}}>Tentar novamente</button>
            </div>
          )}
          {!loading&&!erro&&secao==="gerencial"&&(
            <SecaoGerencial data={gData} setData={setGData}/>
          )}
          {!loading&&!erro&&secao==="produtividade"&&(
            <SecaoProdutividade data={pData} subPag={subPag} filtros={filtroP}/>
          )}
        </div>
      </div>
    </div>
  );
}
