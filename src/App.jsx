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

const DARK  = "#1a2332";
const DARK2 = "#243044";
const MENU_BG = "#0F1F35";

// Equipe e cores (usadas em Gerencial e Produtividade)
const EQUIPE   = ["Mariana","Wilder","Giovanni","Carla","Darlan"];
const COR_P    = { Mariana:"#7C3AED",Wilder:"#0369A1",Giovanni:"#059669",Carla:"#D97706",Darlan:"#DC2626" };
const SIN_COR  = { Verde:"#16A34A",Amarelo:"#D97706",Vermelho:"#DC2626" };
const SIN_BG   = { Verde:"#DCFCE7",Amarelo:"#FEF9C3",Vermelho:"#FEE2E2" };
const TIPO_P   = ["Reajuste","Renovação","Up Selling","Defesa de Território","Alteração de Escopo","BID/Cotação","Outros"];
const TIPO_COR_P = { "Reajuste":"#7C3AED","Renovação":"#0891B2","Up Selling":"#059669","Defesa de Território":"#D97706","Alteração de Escopo":"#0369A1","BID/Cotação":"#DC2626","Outros":"#6B7280" };
const MENU_W_OPEN  = 220;
const MENU_W_CLOSE = 48;

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const matchPessoa = s => EQUIPE.find(p => (s||"").toUpperCase().includes(p.toUpperCase())) || null;

const matchSin = v => {
  const s = (v||"").toString().toUpperCase();
  if (s.includes("VERDE"))   return "verde";
  if (s.includes("AMARELO")) return "amarelo";
  if (s.includes("VERM"))    return "vermelho";
  return null;
};

const matchCatP = (ativ, tipo) => {
  const c = ((ativ||"")+" "+(tipo||"")).toUpperCase();
  if (c.includes("REAJUSTE")||c.includes("NOTIFICAÇ")||c.includes("CARTA DE REAJUSTE")) return "Reajuste";
  if (c.includes("RENOVAÇ")) return "Renovação";
  if (c.includes("UP-SELLING")||c.includes("UP SELLING")||c.includes("AUMENTO DE ESCOPO")) return "Up Selling";
  if (c.includes("DEFESA DE TERRIT")) return "Defesa de Território";
  if ((c.includes("ALTERAÇ")&&c.includes("ESCOPO"))||c.includes("REVISÃO DE ESCOPO")||c.includes("ADITIVO CONTRATUAL")) return "Alteração de Escopo";
  if (c.includes("BID")||c.includes("COTAÇ")) return "BID/Cotação";
  return "Outros";
};
const brl = v => (!v&&v!==0)?"—"
  : new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL",maximumFractionDigits:0}).format(v);
const num = v => {
  if (v==null||v==="") return 0;
  const s = String(v).trim();
  // Currency: "R$ 1.234,56" -> strip R$, spaces, thousand dots, convert comma
  // Decimal: "0.0756" -> keep as is
  const cleaned = s.replace(/R\$\s*/g,"").replace(/\s/g,"");
  // If it has comma as decimal separator (Brazilian): "1.234,56"
  if (/\d\.\d{3},/.test(cleaned) || /,\d{1,2}$/.test(cleaned)) {
    return parseFloat(cleaned.replace(/\./g,"").replace(",",".")) || 0;
  }
  return parseFloat(cleaned) || 0;
};
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
  if (!raw.length) return [];
  // Usa índice fixo baseado na ordem das colunas da planilha:
  // 0:Mês Vigência, 1:Grupo Cliente, 2:CR, 3:DESCRI CR, 4:Responsável Farmer,
  // 5:Status Real, 6:Tipo de Negócio, 7:Devido, 8:Aplicado,
  // 9:Inicio Negociação, 10:Data de Aprovação, 11:Informações
  const headers = Object.keys(raw[0]);
  const idx = kws => {
    const i = headers.findIndex(h => kws.some(kw => h.toUpperCase().trim() === kw.toUpperCase().trim()));
    return i >= 0 ? i : null;
  };
  const idxMes    = idx(["Mês Vigência","MES VIGENCIA"]);
  const idxGrupo  = idx(["Grupo Cliente","GRUPO CLIENTE"]);
  const idxCR     = idx(["CR"]);
  const idxDescr  = idx(["DESCRI CR","DESCRICAO CR","DESCRI"]);
  const idxFarmer = idx(["Responsável Farmer","RESPONSAVEL FARMER"]);
  const idxStatus = idx(["Status Real","STATUS REAL"]);
  const idxTipo   = idx(["Tipo de Negócio","TIPO DE NEGOCIO"]);
  const idxDevido = idx(["Devido","DEVIDO"]);
  const idxAplic  = idx(["Aplicado","APLICADO"]);
  const idxInicio = idx(["Inicio Negociação","Início Negociação","INICIO NEGOCIACAO"]);
  const idxAprov  = idx(["Data de Aprovação","DATA DE APROVACAO"]);
  const idxInfo   = idx(["Informações","INFORMACOES"]);

  const getByIdx = (r, i) => i !== null ? (Object.values(r)[i] ?? "") : "";

  return raw.map((r, i) => {
    const grupo  = getByIdx(r, idxGrupo).toString().trim();
    const cr     = getByIdx(r, idxCR).toString().trim();
    // Ignora linhas sem grupo válido ou CR inválido
    if (!grupo || grupo.length > 60) return null;
    if (!cr || cr.length > 15 || !/^\d/.test(cr.trim())) return null;
    const statusReal = getByIdx(r, idxStatus).toString().trim() || "N/A";
    const mesVig     = getByIdx(r, idxMes).toString().trim();
    const inicioNeg  = parseData(getByIdx(r, idxInicio));
    const dataAprov  = parseData(getByIdx(r, idxAprov));
    const devido     = getByIdx(r, idxDevido).toString().trim();
    const aplicado   = getByIdx(r, idxAplic).toString().trim();
    const mesNum     = parseMesNum(mesVig);
    const mesAtual   = new Date().getMonth() + 1;
    let status = statusReal;
    if (!dataAprov && inicioNeg && mesNum && mesNum > mesAtual) status = "Em Negociação Antecipada";
    else if (!dataAprov && inicioNeg) status = "Em Negociação";
    else if (!dataAprov && !inicioNeg && mesNum && mesNum > mesAtual) status = "Aguardando Reajuste";
    const diasNeg = dataAprov
      ? diasEntre(inicioNeg, dataAprov)
      : (inicioNeg ? diasEntre(inicioNeg, new Date()) : null);
    return {
      _id: i, grupo, cr,
      descr:   getByIdx(r, idxDescr).toString().trim(),
      tipo:    getByIdx(r, idxTipo).toString().trim() || "Reajuste",
      status, statusReal, mes: mesVig,
      farmer:  getByIdx(r, idxFarmer).toString().trim(),
      info:    getByIdx(r, idxInfo).toString().trim(),
      devido, aplicado, inicioNeg, dataAprov, diasNeg,
      concluido: !!dataAprov,
      _edited: false,
    };
  }).filter(Boolean);
};

const processProd = raw => {
  if(!raw||!raw.length) return [];
  const keys=Object.keys(raw[0]);
  const col=(...kws)=>keys.find(k=>kws.some(kw=>k.toUpperCase().replace(/\s+/g," ").trim().includes(kw.toUpperCase())))||null;
  const val=(r,...kws)=>{const k=col(...kws);return k?r[k]:"";};
  return raw.map(r=>{
    const ativ=(val(r,"ATIVIDADE1")||val(r,"ATIVIDADE")||"").toString().trim().toUpperCase();
    const tipo=(val(r,"TIPO DE PROPOSTA")||"").toString().trim().toUpperCase();
    const obs=(val(r,"OBS")||"").toString().trim();
    const cValAtual=col("VALOR CONTRATO ATUAL");
    const valAtual=cValAtual?numP(r[cValAtual]):0;
    const cValPleito=col("REAJUSTE + PLEITO","COM REAJUSTE + PLEITO");
    const cValReaj=col("VALOR CONTRATO COM REAJUSTE");
    const valPleito=(cValPleito?numP(r[cValPleito]):0)||(cValReaj?numP(r[cValReaj]):0);
    const row={
      isRevisao:(val(r,"REVISÃO","REVISAO")||"").toString().toUpperCase().includes("REVIS"),
      nProposta:(val(r,"Nº PROPOSTA","N PROPOSTA","PROPOSTA")||"").toString().trim(),
      grupoCliente:(val(r,"GRUPO CLIENTE")||"").toString().trim(),
      cliente:(val(r,"CLIENTE")||"").toString().trim(),
      respFarmer:(val(r,"RESP. FARMER","RESPONSÁVEL FARMER")||"").toString().trim(),
      respHunter:(val(r,"RESP. HUNTER","RESPONSÁVEL HUNTER")||"").toString().trim(),
      atividade:ativ,tipoProposta:tipo,obs,
      valAtual,valPleito,diferenca:valPleito-valAtual,
      pctReaj:   numP(val(r,"REAJUSTE CONTRATUAL (%)")),
      pctPleito: numP(val(r,"PLEITO (%)")),
      aprovR:    numP(val(r,"APROVADO PELO CLIENTE (R$)")),
      aprovPct:  numP(val(r,"APROVADO PELO CLIENTE (%)")),
      status:(val(r,"STATUS")||"").toString().trim(),
      mes:(val(r,"Mês","MES")||"").toString().trim(),
      semana:(val(r,"Semana","SEMANA")||"").toString().trim(),
      escopo:(val(r,"Escopo Atuação","ESCOPO")||"").toString().trim(),
      sinalizacao:matchSinP(val(r,"Sinalização","SINALIZACAO","SINALIZ")||""),
      fezPec:val(r,"FEZ PEC"),fezAbertura:val(r,"FEZ ABERTURA"),
      fezProposta:val(r,"FEZ PROPOSTA COMERCIAL"),fezCarta:val(r,"FEZ CARTA DE REAJUSTE"),
      fezNotif:val(r,"FEZ CARTA DE NOTIF","FEZ NOTIF"),
    };
    row.responsavel=matchPessoaP(row.respFarmer,row.respHunter);
    row.categoria=matchCatP(ativ,tipo);
    row.entregaveis=matchEntP(row);
    return row;
  }).filter(r=>r.responsavel);
};

const contarEntP = rows => {
  const c={};let t=0;
  for(const r of rows) for(const e of r.entregaveis){c[e]=(c[e]||0)+1;t++;}
  return Object.entries(c).map(([cat,qtd])=>({cat,qtd,p:t?Math.round(qtd/t*100):0})).sort((a,b)=>b.qtd-a.qtd);
};

const clienteStatsP = (ativs) => {
  const by={};
  for(const r of ativs){
    const g=r.grupoCliente||r.cliente;
    if(!by[g]) by[g]={rows:[],sin:null};
    by[g].rows.push(r);
    if(!by[g].sin&&r.sinalizacao) by[g].sin=r.sinalizacao;
  }
  return Object.entries(by).map(([g,{rows,sin}])=>{
    const wR=rows.filter(r=>r.pctReaj>0),wA=rows.filter(r=>r.aprovPct>0);
    const avg=(arr,fn)=>arr.length?arr.reduce((s,r)=>s+fn(r),0)/arr.length:null;
    return{
      grupo:g,rows,sin,
      avgReaj:avg(wR,r=>r.pctReaj),avgPleito:avg(wR,r=>r.pctPleito),avgAprov:avg(wA,r=>r.aprovPct),
      totalAtual:rows.reduce((s,r)=>s+r.valAtual,0),
      totalPleito:rows.reduce((s,r)=>s+r.valPleito,0),
      totalDif:rows.reduce((s,r)=>s+r.diferenca,0),
    };
  }).sort((a,b)=>b.rows.length-a.rows.length);
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
    return Object.entries(by).map(([g,crs])=>({g,crs})).sort((a,b)=>(a.g||"").localeCompare(b.g||""));
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

// ─── PRODUTIVIDADE (baseada no app original) ──────────────────────────────────
const CAT_COM   = ["Reajuste","Renovação","Up Selling","Defesa de Território","Alteração de Escopo","Outros"];
const CAT_COM_C = { "Reajuste":"#7C3AED","Renovação":"#0891B2","Up Selling":"#059669","Defesa de Território":"#D97706","Alteração de Escopo":"#0369A1","Outros":"#6B7280" };
const CAT_ENT_C = { "PEC":"#0369A1","Abertura de Custo":"#059669","Proposta Comercial":"#D97706","Carta de Reajuste":"#7C3AED","Notificação de Reajuste":"#DC2626","Revisão de Escopo":"#0891B2" };
const SIN_C2    = { verde:"#16A34A",amarelo:"#D97706",vermelho:"#DC2626" };
const SIN_BG2   = { verde:"#DCFCE7",amarelo:"#FEF9C3",vermelho:"#FEE2E2" };
const SIN_LB2   = { verde:"Boa Negociação",amarelo:"Moderada",vermelho:"Difícil" };
const COLORS2   = { MARIANA:"#7C3AED",WILDER:"#0369A1",GIOVANNI:"#059669",CARLA:"#D97706",DARLAN:"#DC2626" };

const fmtPct2 = v => {
  if (v==null||v===""||isNaN(v)) return "—";
  const n = parseFloat(v);
  if (n===0) return "—";
  // If value > 1, it's already in % form (e.g. 6.86 = 6.86%)
  // If value <= 1, it's decimal (e.g. 0.0686 = 6.86%)
  const pct = n > 1 ? n : n * 100;
  return `${pct.toFixed(2)}%`;
};

// Reprocessa dados de produtividade com colunas corretas
const processProdFull = raw => {
  if (!raw?.length) return [];
  const headers = Object.keys(raw[0]);
  const col = (...kws) => headers.find(k => kws.some(kw => k.toUpperCase().replace(/\s+/g," ").trim().includes(kw.toUpperCase()))) || null;
  const val = (r, ...kws) => { const k=col(...kws); return k ? (r[k]??"") : ""; };

  // Agrupa por proposta base para pegar última revisão
  const propBase = np => (np||"").toString().trim().replace(/\s*[-–]?\s*rev\.?\s*[\d]+$/i,"").trim();
  const byProp = {};
  raw.forEach((r,i) => {
    const np  = val(r,"Nº Proposta","N PROPOSTA").toString().trim();
    const rev = parseInt(val(r,"Nº da Revisão","REVISÃO","REVISAO"))||0;
    const key = `${val(r,"Grupo Cliente")}||${propBase(np)}`;
    if (!byProp[key] || rev > byProp[key].rev) byProp[key] = {r, rev, i};
  });
  // Marca quais são última revisão
  const ultimasIdx = new Set(Object.values(byProp).map(x=>x.i));

  return raw.map((r,i) => {
    const respF = val(r,"Responsável Farmer","RESP. FARMER").toString().trim();
    const respH = val(r,"Responsável Hunter","RESP. HUNTER").toString().trim();
    const responsavel = matchPessoa(respF) || matchPessoa(respH);
    if (!responsavel) return null;

    const tipo    = val(r,"Tipo de Negócio","TIPO DE NEGOCIO","TIPO DE PROPOSTA").toString().trim();
    const obs     = val(r,"OBS.","OBS").toString().trim();
    const semana  = val(r,"Semana","SEMANA").toString().trim();
    const nRev    = parseInt(val(r,"Nº da Revisão","REVISÃO","REVISAO"))||0;
    const isRevisao = nRev > 0;
    const isUltima  = ultimasIdx.has(i);

    const valAtual  = num(val(r,"Valor Contrato Atual"));
    const valPleito = num(val(r,"Valor Final / Pleito","Valor Proposta"));
    const isSim = v => ["SIM","S"].includes((v||"").toString().toUpperCase().trim());

    const found = new Set();
    if(isSim(val(r,"Fez PEC?")))                         found.add("PEC");
    if(isSim(val(r,"Fez Abertura de Custo?")))           found.add("Abertura de Custo");
    if(isSim(val(r,"Fez Proposta Comercial?")))          found.add("Proposta Comercial");
    if(isSim(val(r,"Fez Carta de Reajuste?")))           found.add("Carta de Reajuste");
    if(isSim(val(r,"Fez Notificação de Reajuste?")))     found.add("Notificação de Reajuste");

    return {
      responsavel, isRevisao, isUltima, nRev,
      nProposta:    val(r,"Nº Proposta","N PROPOSTA").toString().trim(),
      grupoCliente: val(r,"Grupo Cliente").toString().trim(),
      unidade:      val(r,"Unidade / Filial","UNIDADE","FILIAL").toString().trim(),
      escopo:       val(r,"Escopo Atuação","ESCOPO").toString().trim(),
      mes:          val(r,"Mês","MES").toString().trim(),
      semana,
      tipo,
      obs,
      status:       val(r,"Status","STATUS").toString().trim(),
      sinalizacao:  matchSin(val(r,"Semáforo (Dificuldade)","Semáforo","Sinalização","")),
      valAtual, valPleito,
      diferenca:    valPleito - valAtual,
      pctReaj:      num(val(r,"% Reajuste Contrato")),
      pctPleito:    num(val(r,"% Reajuste Pleito")),
      aprovPct:     num(val(r,"% Reajuste Aceito")),
      aprovR:       num(val(r,"Valor Proposta")),
      entregaveis:  found.size ? [...found] : [],
      categoria:    matchCatP("", tipo),
    };
  }).filter(Boolean);
};

// Filtra apenas última revisão por proposta para cálculos de valor
const ultimasRevisoes = rows => rows.filter(r => r.isUltima);

const TagP = ({label, color}) => (
  <span style={{fontSize:9,padding:"2px 7px",background:`${color}15`,color,
                border:`1px solid ${color}30`,whiteSpace:"nowrap",borderRadius:4,fontWeight:500}}>{label}</span>
);

const BarP = ({p, color}) => (
  <div style={{flex:1,height:5,background:"#F1F4F8",borderRadius:3,overflow:"hidden"}}>
    <div style={{width:`${Math.min(p,100)}%`,height:"100%",background:color,borderRadius:3,transition:"width .4s"}}/>
  </div>
);

function SecaoProdutividade({ rawData, subPag, setSubPag, filtros }) {
  const data = React.useMemo(() => processProdFull(rawData), [rawData]);
  const [fCat,    setFCat]    = useState("Todas");
  const [obsOpen, setObsOpen] = useState(null);

  // Usa filtros do menu lateral
  const fResp = filtros?.fResp || "Todos";
  const fStat = filtros?.fStat || "Todos";
  const fMes  = filtros?.fMes  || "Todos";
  const fSem  = filtros?.fSem  || "Todas";
  const fEsc  = filtros?.fEsc  || "Todos";
  const fSin  = filtros?.fSin  || "Todos";
  const fTipo = filtros?.fTipo || "Todos";

  const ativs = data.filter(r => r.isUltima);
  const revs  = data.filter(r => r.isRevisao);

  // Opções de filtro derivadas dos dados
  const mesesDisp   = ORDEM_MES.filter(m => data.some(r => r.mes && m.toUpperCase().startsWith(r.mes.toUpperCase().substring(0,3))));
  const semanasDisp = [...new Set(data.map(r=>r.semana).filter(Boolean))].sort();
  const escoposDisp = [...new Set(data.map(r=>r.escopo).filter(Boolean))].sort();
  const tiposDisp   = [...new Set(data.map(r=>r.tipo).filter(Boolean))].sort();

  const applyAll = rows => {
    let r = rows;
    if (fMes  !== "Todos") r = r.filter(x => x.mes && fMes.toUpperCase().startsWith(x.mes.toUpperCase().substring(0,3)));
    if (fSem  !== "Todas") r = r.filter(x => x.semana === fSem);
    if (fResp !== "Todos") r = r.filter(x => x.responsavel === fResp);
    if (fEsc  !== "Todos") r = r.filter(x => x.escopo === fEsc);
    if (fSin  !== "Todos") r = r.filter(x => x.sinalizacao === fSin);
    if (fTipo !== "Todos") r = r.filter(x => x.tipo === fTipo);
    if (fStat === "Aprovado")      r = r.filter(x => x.status.toUpperCase().includes("APROVADO"));
    else if (fStat === "Em Negociação") r = r.filter(x => x.status.toUpperCase().includes("NEGOCI"));
    else if (fStat === "Recusado") r = r.filter(x => x.status.toUpperCase().includes("RECUS"));
    return r;
  };
  const applyMesSem = rows => {
    let r = rows;
    if (fMes !== "Todos") r = r.filter(x => x.mes && fMes.toUpperCase().startsWith(x.mes.toUpperCase().substring(0,3)));
    if (fSem !== "Todas") r = r.filter(x => x.semana === fSem);
    return r;
  };
  const applyStatus = rows => {
    if (fStat === "Aprovado")      return rows.filter(r => r.status.toUpperCase().includes("APROVADO"));
    if (fStat === "Em Negociação") return rows.filter(r => r.status.toUpperCase().includes("NEGOCI"));
    if (fStat === "Recusado")      return rows.filter(r => r.status.toUpperCase().includes("RECUS"));
    return rows;
  };
  const filterRows = rows => applyAll(rows);

  const clienteStats = () => {
    const base = applyAll(data);
    const by = {};
    for (const r of base) {
      const g = (r.grupoCliente||"").trim();
      if (!g || g.length < 2) continue;
      if (!by[g]) by[g] = {rows:[],sin:null};
      by[g].rows.push(r);
      if (!by[g].sin && r.sinalizacao) by[g].sin = r.sinalizacao;
    }
    return Object.entries(by).map(([g,{rows,sin}]) => {
      const ults = rows.filter(r=>r.isUltima);
      const wR = ults.filter(r=>r.pctReaj>0), wA = ults.filter(r=>r.aprovPct>0);
      const avg = (arr,fn) => arr.length ? arr.reduce((s,r)=>s+fn(r),0)/arr.length : null;
      return {
        grupo:g, rows, sin,
        avgReaj:   avg(wR, r=>r.pctReaj),    // decimal (0.0756)
        avgPleito: avg(wR, r=>r.pctPleito),  // decimal
        avgAprov:  avg(wA, r=>r.aprovPct),   // decimal
        totalAtual:  ults.reduce((s,r)=>s+r.valAtual,0),
        totalPleito: ults.reduce((s,r)=>s+r.valPleito,0),
        totalDif:    ults.reduce((s,r)=>s+r.diferenca,0),
      };
    }).sort((a,b) => b.rows.length - a.rows.length);
  };

  const contarEnt = rows => {
    const c = {}; let t = 0;
    for (const r of rows) for (const e of r.entregaveis) { c[e]=(c[e]||0)+1; t++; }
    return Object.entries(c).map(([cat,qtd]) => ({cat,qtd,p:t?Math.round(qtd/t*100):0})).sort((a,b)=>b.qtd-a.qtd);
  };

  const PillP = ({label, active, color, onClick}) => (
    <button onClick={onClick} style={{
      padding:"4px 12px", borderRadius:99, border:"none", cursor:"pointer",
      fontFamily:"inherit", fontSize:11, fontWeight:500, transition:"all .15s",
      background: active ? (color||DARK) : "#F1F4F8",
      color:      active ? "#fff" : "#64748B",
      whiteSpace:"nowrap",
    }}>{label}</button>
  );



  // ── VISÃO GERAL ──────────────────────────────────────────────────────────────
  if (subPag === "visao") {
    const base      = applyAll(ativs);
    const cStats    = clienteStats();
    const scopeCat  = cat => base.filter(r => r.categoria === cat);

    return (
      <div>


        {/* KPIs */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10,marginBottom:20}}>
          {[
            {l:"Propostas",       v:applyAll(data).filter(r=>!r.isRevisao).length},
            {l:"Revisões",        v:applyAll(data).filter(r=>r.isRevisao).length},
            {l:"Contrato Atual",  v:brl(base.reduce((s,r)=>s+r.valAtual,0))},
            {l:"Valor c/ Pleito", v:brl(base.reduce((s,r)=>s+r.valPleito,0))},
            {l:"Diferença",       v:brl(base.reduce((s,r)=>s+r.diferenca,0)), hl:true},
          ].map(i=>(
            <div key={i.l} style={{background:i.hl?"#F0FDF4":"#F8FAFC",borderRadius:10,
                                    padding:"14px 16px",boxShadow:"0 1px 4px rgba(0,0,0,.07)",
                                    borderLeft:`3px solid ${i.hl?"#16A34A":"#E2E8F0"}`}}>
              <div style={{fontSize:10,color:"#94A3B8",marginBottom:4}}>{i.l}</div>
              <div style={{fontSize:20,fontWeight:700,color:i.hl?"#16A34A":DARK}}>{i.v}</div>
            </div>
          ))}
        </div>

        {/* Analistas */}
        <div style={{fontSize:11,fontWeight:600,letterSpacing:".06em",textTransform:"uppercase",
                     color:"#94A3B8",marginBottom:12}}>Analistas</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14,marginBottom:24}}>
          {["MARIANA","WILDER","GIOVANNI"].map(p => {
            const sp    = base.filter(r=>r.responsavel===p);
            const revP  = revs.filter(r=>r.responsavel===p);
            const aprov = sp.filter(r=>r.status.toUpperCase().includes("APROVADO"));
            const cor   = COLORS2[p];
            return (
              <div key={p} style={{background:`${cor}08`,borderRadius:12,padding:18,
                                    borderTop:`3px solid ${cor}`,boxShadow:"0 1px 4px rgba(0,0,0,.07)",
                                    border:`1px solid ${cor}20`}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
                  <div>
                    <div style={{fontSize:10,color:"#94A3B8",marginBottom:2}}>Analista</div>
                    <div style={{fontSize:16,fontWeight:700,color:cor}}>{p}</div>
                    <div style={{fontSize:10,color:"#94A3B8",marginTop:2}}>
                      {sp.length} atividades · {revP.length} revisões
                    </div>
                  </div>
                  <div style={{textAlign:"right",fontSize:11}}>
                    <div style={{color:"#16A34A"}}>✓ {aprov.length} aprovados</div>
                    <div style={{color:"#94A3B8",marginTop:2}}>{revP.length} revisões</div>
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:12}}>
                  {[
                    {l:"Contrato Atual", v:brl(sp.reduce((s,r)=>s+r.valAtual,0))},
                    {l:"Com Pleito",     v:brl(sp.reduce((s,r)=>s+r.valPleito,0))},
                    {l:"Diferença",      v:brl(sp.reduce((s,r)=>s+r.diferenca,0)), hl:true},
                  ].map(k=>(
                    <div key={k.l} style={{background:"#fff",borderRadius:8,padding:"7px 8px",
                                            border:"1px solid #E2E8F0"}}>
                      <div style={{fontSize:9,color:"#94A3B8",marginBottom:2}}>{k.l}</div>
                      <div style={{fontSize:11,fontWeight:600,color:k.hl?"#16A34A":DARK}}>{k.v}</div>
                    </div>
                  ))}
                </div>
                <div style={{borderTop:"1px solid #F1F4F8",paddingTop:10}}>
                  <div style={{fontSize:9,color:"#94A3B8",marginBottom:6,textTransform:"uppercase",letterSpacing:".06em"}}>Por Tipo</div>
                  {CAT_COM.map(cat=>{
                    const d=sp.filter(r=>r.categoria===cat);
                    if(!d.length) return null;
                    const dif=d.reduce((s,r)=>s+r.diferenca,0);
                    return(
                      <div key={cat} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          <div style={{width:6,height:6,borderRadius:99,background:CAT_COM_C[cat]}}/>
                          <span style={{fontSize:11,color:"#64748B"}}>{cat}</span>
                        </div>
                        <div style={{display:"flex",gap:8}}>
                          <span style={{fontSize:11,color:cor}}>{d.length}x</span>
                          <span style={{fontSize:11,color:"#16A34A",minWidth:60,textAlign:"right"}}>{dif>0?brl(dif):"—"}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>

        {/* Liderança */}
        <div style={{fontSize:11,fontWeight:600,letterSpacing:".06em",textTransform:"uppercase",
                     color:"#94A3B8",marginBottom:12}}>Liderança</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:14,marginBottom:24}}>
          {["Carla","Darlan"].map(p=>{
            const sp   = base.filter(r=>r.responsavel===p);
            const revP = revs.filter(r=>r.responsavel===p);
            const aprov= sp.filter(r=>r.status.toUpperCase().includes("APROVADO"));
            const cor  = COR_P[p]||"#D97706";
            const porCat = [...new Set(sp.map(r=>r.tipo).filter(Boolean))].map(cat=>({cat,n:sp.filter(r=>r.tipo===cat).length})).filter(x=>x.n>0).sort((a,b)=>b.n-a.n);
            return (
              <div key={p} style={{background:`${cor}08`,borderRadius:12,padding:18,
                                    borderTop:`3px solid ${cor}`,boxShadow:"0 1px 4px rgba(0,0,0,.07)",
                                    border:`1px solid ${cor}20`}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
                  <div>
                    <div style={{fontSize:10,color:"#94A3B8",marginBottom:2}}>Liderança</div>
                    <div style={{fontSize:16,fontWeight:700,color:cor}}>{p}</div>
                    <div style={{fontSize:10,color:"#94A3B8",marginTop:2}}>{sp.length} interações · {revP.length} revisões</div>
                  </div>
                  <div style={{background:"#DCFCE7",borderRadius:8,padding:"4px 10px",textAlign:"center"}}>
                    <div style={{fontSize:16,fontWeight:700,color:"#16A34A"}}>{aprov.length}</div>
                    <div style={{fontSize:9,color:"#16A34A"}}>aprovados</div>
                  </div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,marginBottom:10}}>
                  {[
                    {l:"Interações",  v:sp.length},
                    {l:"Atual",       v:brl(sp.reduce((s,r)=>s+r.valAtual,0))},
                    {l:"Com Pleito",  v:brl(sp.reduce((s,r)=>s+r.valPleito,0))},
                    {l:"Diferença",   v:brl(sp.reduce((s,r)=>s+r.diferenca,0)), hl:true},
                  ].map(k=>(
                    <div key={k.l} style={{background:"#fff",borderRadius:8,padding:"7px 8px",
                                            border:"1px solid #E2E8F0"}}>
                      <div style={{fontSize:9,color:"#94A3B8",marginBottom:2}}>{k.l}</div>
                      <div style={{fontSize:11,fontWeight:600,color:k.hl?"#16A34A":DARK}}>{k.v}</div>
                    </div>
                  ))}
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                  {porCat.map(({cat,n})=>(
                    <TagP key={cat} label={`${cat} ${n}×`} color={CAT_COM_C[cat]||TIPO_COR_P[cat]||"#6B7280"}/>
                  ))}
                </div>
              </div>
            );
          })}
        </div>

        {/* Semáforo */}
        <div style={{fontSize:11,fontWeight:600,letterSpacing:".06em",textTransform:"uppercase",
                     color:"#94A3B8",marginBottom:12}}>Semáforo de Clientes</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14}}>
          {["verde","amarelo","vermelho"].map(sin=>{
            const list=cStats.filter(c=>c.sin===sin);
            return(
              <div key={sin} style={{background:`${SIN_C2[sin]}08`,borderRadius:12,padding:16,
                                      borderTop:`3px solid ${SIN_C2[sin]}`,
                                      boxShadow:"0 1px 4px rgba(0,0,0,.07)",
                                      border:`1px solid ${SIN_C2[sin]}25`}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                  <div style={{display:"flex",alignItems:"center",gap:7}}>
                    <div style={{width:9,height:9,borderRadius:99,background:SIN_C2[sin]}}/>
                    <span style={{fontSize:13,fontWeight:600,color:SIN_C2[sin]}}>{SIN_LB2[sin]}</span>
                  </div>
                  <div style={{fontSize:22,fontWeight:800,color:SIN_C2[sin]}}>{list.length}</div>
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:4}}>
                  {list.slice(0,5).map(c=>(
                    <div key={c.grupo} style={{display:"flex",justifyContent:"space-between",
                      padding:"5px 10px",background:SIN_BG2[sin],borderRadius:6}}>
                      <span style={{fontSize:11,color:SIN_C2[sin],fontWeight:500}}>{c.grupo}</span>
                      <span style={{fontSize:10,color:SIN_C2[sin]}}>{c.rows.length}×</span>
                    </div>
                  ))}
                  {list.length>5&&<div style={{fontSize:10,color:"#94A3B8",textAlign:"center",paddingTop:3}}>+{list.length-5} grupos</div>}
                  {!list.length&&<div style={{fontSize:11,color:"#94A3B8",textAlign:"center",padding:"10px 0"}}>Nenhum classificado</div>}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ── ESFORÇO ──────────────────────────────────────────────────────────────────
  if (subPag === "esforco") {
    const base     = applyAll(data);
    const analBase = base;
    const entGeral = contarEnt(analBase);
    return (
      <div>
        <div style={{marginBottom:20}}>
          <div style={{fontSize:11,fontWeight:600,letterSpacing:".06em",color:"#94A3B8",
                       textTransform:"uppercase",marginBottom:4}}>Esforço Operacional</div>
          <div style={{fontSize:20,fontWeight:700,color:DARK}}>
            Mesa de <span style={{color:"#2563EB"}}>Trabalho</span>
          </div>
          <div style={{fontSize:12,color:"#94A3B8",marginTop:2}}>{analBase.length} interações no período</div>
        </div>

        {/* Distribuição geral */}
        <div style={{background:"#fff",borderRadius:12,padding:18,marginBottom:16,
                     boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
          <div style={{fontSize:11,fontWeight:600,color:"#94A3B8",textTransform:"uppercase",
                       letterSpacing:".06em",marginBottom:14}}>Distribuição do Time</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            {entGeral.map(({cat,qtd,p})=>(
              <div key={cat} style={{padding:"10px 14px",background:"#F8FAFC",borderRadius:8,
                                     borderLeft:`4px solid ${CAT_ENT_C[cat]||"#94A3B8"}`}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                  <span style={{fontSize:12,fontWeight:500,color:CAT_ENT_C[cat]||DARK}}>{cat}</span>
                  <div style={{display:"flex",gap:10,alignItems:"center"}}>
                    <span style={{fontSize:11,color:"#94A3B8"}}>{qtd}×</span>
                    <span style={{fontSize:15,fontWeight:700,color:DARK}}>{p}%</span>
                  </div>
                </div>
                <BarP p={p} color={CAT_ENT_C[cat]||"#94A3B8"}/>
              </div>
            ))}
          </div>
        </div>

        {/* Por analista */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14}}>
          {["Mariana","Wilder","Giovanni"].map(p=>{
            const rows  = applyAll(data).filter(r=>r.responsavel===p);
            const ents  = contarEnt(rows);
            const cor   = COR_P[p]||"#7C3AED";
            return(
              <div key={p} style={{background:`${cor}08`,borderRadius:12,padding:18,
                                    borderTop:`3px solid ${cor}`,boxShadow:"0 1px 4px rgba(0,0,0,.07)",
                                    border:`1px solid ${cor}20`}}>
                <div style={{marginBottom:14}}>
                  <div style={{fontSize:10,color:"#94A3B8",marginBottom:2}}>Analista</div>
                  <div style={{fontSize:16,fontWeight:700,color:cor}}>{p}</div>
                  <div style={{fontSize:10,color:"#94A3B8",marginTop:2}}>
                    {rows.filter(r=>!r.isRevisao).length} atividades · {rows.filter(r=>r.isRevisao).length} revisões
                  </div>
                </div>
                {!ents.length&&<div style={{color:"#94A3B8",fontSize:12}}>Sem checklist preenchido</div>}
                {ents.map(({cat,qtd,p:pr})=>(
                  <div key={cat} style={{marginBottom:10}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                      <div style={{display:"flex",alignItems:"center",gap:5}}>
                        <div style={{width:8,height:8,borderRadius:99,background:CAT_ENT_C[cat]||"#94A3B8"}}/>
                        <span style={{fontSize:12,color:"#64748B"}}>{cat}</span>
                      </div>
                      <div style={{display:"flex",gap:8}}>
                        <span style={{fontSize:11,color:"#94A3B8"}}>{qtd}×</span>
                        <span style={{fontSize:12,fontWeight:600,color:cor}}>{pr}%</span>
                      </div>
                    </div>
                    <BarP p={pr} color={cor}/>
                  </div>
                ))}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ── POR CLIENTE ───────────────────────────────────────────────────────────────
  if (subPag === "clientes") {
    const cStats  = clienteStats();
    const filtered = fSin==="Todos" ? cStats : cStats.filter(c=>c.sin===fSin);
    return (
      <div>
        <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}>
          <span style={{fontSize:12,color:"#94A3B8"}}>{filtered.length} clientes</span>
        </div>

        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {filtered.map(({grupo,rows,sin,avgReaj,avgPleito,avgAprov,totalAtual,totalPleito,totalDif})=>{
            const filtRows = fResp==="Todos" ? rows : rows.filter(r=>r.responsavel===fResp);
            return(
              <div key={grupo} style={{background:"#fff",borderRadius:12,overflow:"hidden",
                                        boxShadow:"0 1px 4px rgba(0,0,0,.07)",
                                        borderLeft:`4px solid ${sin?SIN_C2[sin]:"#E2E8F0"}`}}>
                <div style={{padding:"12px 18px",display:"flex",justifyContent:"space-between",
                              alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
                  <div>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                      <span style={{fontSize:14,fontWeight:700,color:DARK}}>{grupo}</span>
                      {sin
                        ?<span style={{fontSize:10,padding:"2px 8px",background:SIN_BG2[sin],
                            color:SIN_C2[sin],borderRadius:99,fontWeight:600}}>{SIN_LB2[sin]}</span>
                        :<span style={{fontSize:10,padding:"2px 8px",background:"#F1F4F8",
                            color:"#94A3B8",borderRadius:99}}>Sem classificação</span>}
                    </div>
                    <div style={{fontSize:10,color:"#94A3B8"}}>
                      {filtRows.length} atividades · {filtRows.filter(r=>r.status.toUpperCase().includes("APROVADO")).length} aprovadas
                    </div>
                  </div>
                  <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
                    {[
                      {l:"Total Atual",   v:totalAtual>0?brl(totalAtual):"—"},
                      {l:"Total Pleito",  v:totalPleito>0?brl(totalPleito):"—"},
                      {l:"Ganho",         v:totalDif>0?brl(totalDif):"—",   c:"#16A34A"},
                      {l:"% Reaj.",       v:avgReaj!=null?fmtPct2(avgReaj):"—"},
                      {l:"% Pleito",      v:avgPleito!=null?fmtPct2(avgPleito):"—"},
                      {l:"% Aceito",      v:avgAprov!=null?fmtPct2(avgAprov):"—", c:avgAprov!=null?"#16A34A":undefined},
                    ].map(i=>(
                      <div key={i.l} style={{textAlign:"right"}}>
                        <div style={{fontSize:9,color:"#94A3B8",marginBottom:2,textTransform:"uppercase",letterSpacing:".04em"}}>{i.l}</div>
                        <div style={{fontSize:13,fontWeight:600,color:i.c||DARK}}>{i.v}</div>
                      </div>
                    ))}
                  </div>
                </div>
                {/* Tabela interna de unidades — uma linha por unidade (última revisão) */}
                {filtRows.length>0&&(()=>{
                  // Agrupa por unidade, pega última revisão
                  const byUnit = {};
                  filtRows.forEach(r => {
                    const key = (r.unidade||r.cliente||"").trim();
                    if (!byUnit[key]) byUnit[key] = {rows:[],last:null};
                    byUnit[key].rows.push(r);
                    if (!byUnit[key].last||r.nRev>byUnit[key].last.nRev) byUnit[key].last=r;
                  });
                  const unitRows = Object.values(byUnit).map(({rows,last})=>({
                    ...last, totalRevs: rows.filter(x=>x.nRev>0).length
                  })).sort((a,b)=>(a.unidade||"").localeCompare(b.unidade||""));
                  return (
                  <div style={{overflowX:"auto",borderTop:"1px solid #F1F4F8"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                      <thead>
                        <tr style={{background:"#F8FAFC"}}>
                          {["Unidade / Filial","Tipo","Rev.","Contrato Atual","Com Pleito","Ganho","% Reajuste","% Aceito","Status","Obs"].map(h=>(
                            <th key={h} style={{textAlign:"left",padding:"7px 12px",fontSize:10,
                              color:"#94A3B8",fontWeight:600,whiteSpace:"nowrap",
                              borderBottom:"1px solid #E2E8F0"}}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {unitRows.map((r,i)=>{
                          const isAprov=r.status.toUpperCase().includes("APROVADO");
                          return(
                            <tr key={i} style={{background:i%2===0?"#fff":"#FAFBFC",
                                                borderTop:"1px solid #F1F4F8"}}>
                              <td style={{padding:"7px 12px",fontWeight:500,color:DARK,whiteSpace:"nowrap"}}>{r.unidade}</td>
                              <td style={{padding:"7px 12px"}}><TagP label={r.tipo||r.categoria} color={CAT_COM_C[r.categoria]||"#6B7280"}/></td>
                              <td style={{padding:"7px 12px",textAlign:"center",
                                color:r.totalRevs>0?"#7C3AED":"#94A3B8",fontWeight:r.totalRevs>0?600:400}}>
                                {r.totalRevs||0}</td>
                              <td style={{padding:"7px 12px",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.valAtual>0?brl(r.valAtual):"—"}</td>
                              <td style={{padding:"7px 12px",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.valPleito>0?brl(r.valPleito):"—"}</td>
                              <td style={{padding:"7px 12px",color:"#16A34A",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.diferenca>0?brl(r.diferenca):"—"}</td>
                              <td style={{padding:"7px 12px",color:"#64748B",whiteSpace:"nowrap"}}>{r.pctReaj>0?fmtPct2(r.pctReaj):"—"}</td>
                              <td style={{padding:"7px 12px",color:"#16A34A",whiteSpace:"nowrap"}}>{r.aprovPct>0?fmtPct2(r.aprovPct):"—"}</td>
                              <td style={{padding:"7px 12px"}}>
                                <span style={{fontSize:10,fontWeight:500,padding:"2px 7px",borderRadius:99,
                                  background:isAprov?"#DCFCE7":"#FEF9C3",
                                  color:isAprov?"#16A34A":"#D97706",whiteSpace:"nowrap"}}>{r.status}</span>
                              </td>
                              <td style={{padding:"7px 12px"}}>
                                {r.obs&&(
                                  <button onClick={()=>setObsOpen(obsOpen===`${grupo}-${i}`?null:`${grupo}-${i}`)}
                                    style={{fontSize:9,padding:"2px 7px",borderRadius:4,border:"1px solid #E2E8F0",
                                      background:"#F8FAFC",color:"#64748B",cursor:"pointer",fontFamily:"inherit"}}>
                                    OBS
                                  </button>
                                )}
                                {obsOpen===`${grupo}-${i}`&&r.obs&&(
                                  <div style={{position:"fixed",zIndex:9999,background:"#1E293B",color:"#F8FAFC",
                                    padding:"10px 14px",borderRadius:8,fontSize:11,maxWidth:300,lineHeight:1.5,
                                    boxShadow:"0 4px 16px rgba(0,0,0,.3)",top:"50%",left:"50%",
                                    transform:"translate(-50%,-50%)",cursor:"pointer"}}
                                    onClick={()=>setObsOpen(null)}>
                                    {r.obs}
                                  </div>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  );
                })()}
              </div>
            );
          })}
          {!filtered.length&&<div style={{background:"#fff",borderRadius:12,padding:"48px 0",
            textAlign:"center",color:"#94A3B8",fontSize:14}}>Nenhum cliente encontrado.</div>}
        </div>
      </div>
    );
  }

  // ── HISTÓRICO ─────────────────────────────────────────────────────────────────
  // Histórico: uma linha por unidade (última revisão) com contagem de revisões
  const filtHistorico = (() => {
    const all = filterRows(data);
    const by = {};
    all.forEach(r => {
      // Chave simples: grupo + unidade (uma linha por unidade)
      const key = `${(r.grupoCliente||"").trim()}||${(r.unidade||"").trim()}`;
      if (!by[key]) by[key] = { rows:[], last:null };
      by[key].rows.push(r);
      if (!by[key].last || r.nRev > by[key].last.nRev) by[key].last = r;
    });
    return Object.values(by)
      .filter(x => x.last)
      .map(({rows, last}) => ({
        ...last,
        totalRevisoes: rows.filter(r => r.nRev > 0).length,
      }))
      .sort((a,b) =>
        (a.grupoCliente||"").localeCompare(b.grupoCliente||"") ||
        (a.unidade||"").localeCompare(b.unidade||"")
      );
  })();
  return (
    <div>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:12}}>
        <span style={{fontSize:12,color:"#94A3B8"}}>{filtHistorico.length} registros</span>
      </div>
      <div style={{background:"#fff",borderRadius:12,overflow:"hidden",boxShadow:"0 1px 4px rgba(0,0,0,.07)"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
            <thead>
              <tr style={{background:"#F8FAFC"}}>
                {["Mês","Sem.","Resp.","Escopo","Grupo","Unidade","Tipo","Rev.","Contrato Atual","Com Pleito","Ganho","% Reaj","% Aceito","Semáforo","Status","OBS"].map(h=>(
                  <th key={h} style={{textAlign:"left",padding:"8px 10px",fontSize:10,
                    color:"#94A3B8",fontWeight:600,whiteSpace:"nowrap",
                    borderBottom:"1px solid #E2E8F0"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtHistorico.map((r,i)=>{
                const isAprov=r.status.toUpperCase().includes("APROVADO");
                const isRecus=r.status.toUpperCase().includes("RECUS");
                const sc=isAprov?{t:"#16A34A",bg:"#DCFCE7"}:isRecus?{t:"#DC2626",bg:"#FEE2E2"}:{t:"#D97706",bg:"#FEF9C3"};
                return(
                  <tr key={i} style={{background:i%2===0?"#fff":"#F8FAFC",borderTop:"1px solid #F1F4F8"}}>
                    <td style={{padding:"7px 10px",color:"#94A3B8",whiteSpace:"nowrap"}}>{r.mes}</td>
                    <td style={{padding:"7px 10px",color:"#94A3B8",whiteSpace:"nowrap",fontSize:10}}>{r.semana}</td>
                    <td style={{padding:"7px 10px",color:COLORS2[r.responsavel],fontWeight:600,
                      whiteSpace:"nowrap",borderLeft:`3px solid ${COLORS2[r.responsavel]||"#E2E8F0"}`}}>
                      {r.responsavel.charAt(0)+r.responsavel.slice(1).toLowerCase()}
                    </td>
                    <td style={{padding:"7px 10px",color:"#94A3B8",whiteSpace:"nowrap",fontSize:10}}>{r.escopo||"—"}</td>
                    <td style={{padding:"7px 10px",maxWidth:110,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:DARK}}>{r.grupoCliente}</td>
                    <td style={{padding:"7px 10px",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontWeight:500}}>{r.unidade}</td>
                    <td style={{padding:"7px 10px"}}><TagP label={r.tipo||r.categoria} color={CAT_COM_C[r.categoria]||"#6B7280"}/></td>
                    <td style={{padding:"7px 10px",textAlign:"center",
                      color:r.totalRevisoes>0?"#7C3AED":"#94A3B8",fontWeight:600}}>
                      {r.totalRevisoes||0}
                    </td>
                    <td style={{padding:"7px 10px",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.valAtual>0?brl(r.valAtual):"—"}</td>
                    <td style={{padding:"7px 10px",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.valPleito>0?brl(r.valPleito):"—"}</td>
                    <td style={{padding:"7px 10px",color:"#16A34A",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{r.diferenca>0?brl(r.diferenca):"—"}</td>
                    <td style={{padding:"7px 10px",color:"#64748B",whiteSpace:"nowrap"}}>{r.pctReaj>0?fmtPct2(r.pctReaj):"—"}</td>
                    <td style={{padding:"7px 10px",color:"#16A34A",whiteSpace:"nowrap"}}>{r.aprovPct>0?fmtPct2(r.aprovPct):"—"}</td>
                    <td style={{padding:"7px 10px"}}>
                      {r.sinalizacao&&(
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          <div style={{width:7,height:7,borderRadius:99,background:SIN_C2[r.sinalizacao]}}/>
                          <span style={{fontSize:10,color:SIN_C2[r.sinalizacao],fontWeight:500}}>{SIN_LB2[r.sinalizacao]}</span>
                        </div>
                      )}
                    </td>
                    <td style={{padding:"7px 10px"}}>
                      <span style={{fontSize:10,fontWeight:500,padding:"2px 7px",borderRadius:99,
                        background:sc.bg,color:sc.t,whiteSpace:"nowrap"}}>{r.status}</span>
                    </td>
                    <td style={{padding:"7px 10px",maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",
                      whiteSpace:"nowrap",fontSize:10,color:"#94A3B8"}}
                      title={r.obs}>{r.obs||"—"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {!filtHistorico.length&&<div style={{textAlign:"center",padding:"40px 0",color:"#94A3B8",fontSize:14}}>
            Nenhum registro encontrado.</div>}
        </div>
      </div>
    </div>
  );
}

// ─── APP PRINCIPAL ────────────────────────────────────────────────────────────
export default function App() {
  const [gData,   setGData]   = useState([]);
  const [pData,   setPData]   = useState([]);
  const [loading, setLoading] = useState(true);
  const [erro,    setErro]    = useState(null);
  const [ultimaAt,setUltimaAt]= useState(null);
  const [secao,   setSecao]   = useState("gerencial");
  const [subPag,  setSubPag]  = useState("visao");
  const [menuOpen,setMenuOpen]= useState(false);

  // Filtros de Produtividade — no menu lateral
  const [pFResp, setPFResp] = useState("Todos");
  const [pFStat, setPFStat] = useState("Todos");
  const [pFMes,  setPFMes]  = useState("Todos");
  const [pFSem,  setPFSem]  = useState("Todas");
  const [pFEsc,  setPFEsc]  = useState("Todos");
  const [pFSin,  setPFSin]  = useState("Todos");
  const [pFTipo, setPFTipo] = useState("Todos");

  const buscarDados = useCallback(async()=>{
    setLoading(true);setErro(null);
    try {
      const [rawG,rawP] = await Promise.all([
        fetchCSV("2073814116"),
        fetchCSV("1622380363"),
      ]);
      setGData(processGerencial(rawG));
      setPData(rawP); // raw rows — processed inside SecaoProdutividade
      setUltimaAt(new Date());
    } catch(e) { setErro(e.message); }
    finally { setLoading(false); }
  },[]);

  useEffect(()=>{ buscarDados(); },[buscarDados]);

  const navItems = [
    { id:"gerencial",     label:"Gerencial",     subs:[] },
    { id:"produtividade", label:"Produtividade",  subs:[
      {id:"visao",    label:"Visão Geral"},
      {id:"esforco",  label:"Esforço"},
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
        height:"100vh",background:MENU_BG,
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
        {menuOpen&&secao==="produtividade"&&pData.length>0&&(()=>{
          const procData = processProdFull(pData);
          const mesesDisp  = ORDEM_MES.filter(m=>procData.some(r=>r.mes&&m.toUpperCase().startsWith(r.mes.toUpperCase().substring(0,3))));
          const semanasDisp= [...new Set(procData.map(r=>r.semana).filter(Boolean))].sort();
          const escoposDisp= [...new Set(procData.map(r=>r.escopo).filter(Boolean))].sort();
          const tiposDisp  = [...new Set(procData.map(r=>r.tipo).filter(Boolean))].sort();

          const BtnF = ({label,active,color,onClick}) => (
            <button onClick={onClick} style={{
              padding:"3px 9px",borderRadius:99,border:"none",cursor:"pointer",
              fontFamily:"inherit",fontSize:11,fontWeight:500,transition:"all .15s",
              background:active?(color||"#60A5FA"):"rgba(255,255,255,.08)",
              color:active?"#fff":"rgba(255,255,255,.5)",whiteSpace:"nowrap",
              marginBottom:3,
            }}>{label}</button>
          );

          const Sec = ({label,children}) => (
            <div style={{marginBottom:14}}>
              <div style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,.3)",
                           textTransform:"uppercase",letterSpacing:".1em",marginBottom:6}}>{label}</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:3}}>{children}</div>
            </div>
          );

          return (
            <div style={{overflowY:"auto",padding:"14px 12px",
                         borderTop:"1px solid rgba(255,255,255,.07)",
                         flex:1, minHeight:0}}>
              <div style={{fontSize:10,fontWeight:700,color:"rgba(255,255,255,.25)",
                           textTransform:"uppercase",letterSpacing:".1em",marginBottom:14}}>
                Filtros
              </div>

              <Sec label="Responsável">
                {["Todos","Mariana","Wilder","Giovanni","Carla","Darlan"].map(p=>(
                  <BtnF key={p} label={p} active={pFResp===p}
                    color={COR_P[p]||"#60A5FA"} onClick={()=>setPFResp(p)}/>
                ))}
              </Sec>

              <Sec label="Status">
                {["Todos","Aprovado","Em Negociação","Recusado"].map(s=>(
                  <BtnF key={s} label={s} active={pFStat===s} onClick={()=>setPFStat(s)}/>
                ))}
              </Sec>

              <Sec label="Mês">
                {["Todos",...mesesDisp].map(m=>(
                  <BtnF key={m} label={m} active={pFMes===m} onClick={()=>setPFMes(m)}/>
                ))}
              </Sec>

              <Sec label="Semana">
                {["Todas",...semanasDisp].map(s=>(
                  <BtnF key={s} label={s} active={pFSem===s} onClick={()=>setPFSem(s)}/>
                ))}
              </Sec>

              <Sec label="Escopo">
                {["Todos",...escoposDisp].map(e=>(
                  <BtnF key={e} label={e} active={pFEsc===e} onClick={()=>setPFEsc(e)}/>
                ))}
              </Sec>

              <Sec label="Semáforo">
                {[
                  {v:"Todos",l:"Todos"},
                  {v:"verde",l:"Boa Negociação",c:"#16A34A"},
                  {v:"amarelo",l:"Moderada",c:"#D97706"},
                  {v:"vermelho",l:"Difícil",c:"#DC2626"},
                ].map(({v,l,c})=>(
                  <BtnF key={v} label={l} active={pFSin===v} color={c} onClick={()=>setPFSin(v)}/>
                ))}
              </Sec>

              <Sec label="Tipo de Negócio">
                {["Todos",...tiposDisp].map(t=>(
                  <BtnF key={t} label={t} active={pFTipo===t} onClick={()=>setPFTipo(t)}/>
                ))}
              </Sec>

              <button onClick={()=>{setPFResp("Todos");setPFStat("Todos");setPFMes("Todos");
                setPFSem("Todas");setPFEsc("Todos");setPFSin("Todos");setPFTipo("Todos");}}
                style={{width:"100%",padding:"6px",borderRadius:8,border:"1px solid rgba(255,255,255,.15)",
                        background:"transparent",color:"rgba(255,255,255,.4)",fontSize:11,
                        cursor:"pointer",fontFamily:"inherit",marginTop:4}}>
                Limpar filtros
              </button>
            </div>
          );
        })()}

        {/* Última atualização */}
        {menuOpen&&ultimaAt&&(
          <div style={{padding:"10px 16px",borderTop:"1px solid rgba(255,255,255,.06)",
                       fontSize:10,color:"rgba(255,255,255,.25)"}}>
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
            <SecaoProdutividade rawData={pData} subPag={subPag} setSubPag={setSubPag}
              filtros={{fResp:pFResp,fStat:pFStat,fMes:pFMes,fSem:pFSem,fEsc:pFEsc,fSin:pFSin,fTipo:pFTipo}}/>
          )}
        </div>
      </div>
    </div>
  );
}
