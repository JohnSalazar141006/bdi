import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, LineChart, Line, ResponsiveContainer, CartesianGrid } from "recharts";
import { LayoutDashboard, Scale, FileSpreadsheet, User, FileText, MessageSquare, Download, CheckCircle2, AlertTriangle, TrendingUp, Plus, Trash2, BarChart3, History, RotateCcw, Printer, Undo2, Redo2, Upload, ShieldCheck } from "lucide-react";

/* ===========================================================
   SISTEMA CONTABLE MARGARITA — aplicación web
   Panel · módulos con vista previa en vivo · gráficos · agente
   =========================================================== */

const MESES=["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const EMPRESAS={
  "Importadora Canaima, C.A.":{rif:"J-06500004-0",modo:"porcentaje",unidad:"Unidades",rubros:["Prendas de vestir","Calzados","Accesorios"],ventas:{"Prendas de vestir":0.36,Calzados:0.54,Accesorios:0.10},unidades:{"Prendas de vestir":0.41,Calzados:0.46,Accesorios:0.13}},
  "Importadora Opus, C.A.":{rif:"",modo:"porcentaje",unidad:"Unidades",rubros:["Prendas de vestir","Calzados","Accesorios"],ventas:{"Prendas de vestir":0.37,Calzados:0.54,Accesorios:0.09},unidades:{"Prendas de vestir":0.41,Calzados:0.46,Accesorios:0.13}},
  "Von Road, C.A.":{rif:"",modo:"porcentaje",unidad:"Unidades",rubros:["Prendas de vestir","Calzados","Accesorios"],ventas:{"Prendas de vestir":0.37,Calzados:0.54,Accesorios:0.09},unidades:{"Prendas de vestir":0.41,Calzados:0.46,Accesorios:0.13}},
  "Plascan, C.A.":{rif:"",modo:"rubro",unidad:"KG",rubros:["Plásticos","Envases","Alimentos"]},
};
const COLORES=["#0E7C6B","#C8892B","#3E6B8C","#9C5BA8","#B5544A"];
const fmt=(n)=>(n??0).toLocaleString("es-VE",{minimumFractionDigits:2,maximumFractionDigits:2});
const fmt0=(n)=>(n??0).toLocaleString("es-VE",{maximumFractionDigits:0});
const slug=(s)=>(s||"doc").split(",")[0].replace(/\W+/g,"_");
const sig=(a,m)=>{const i=MESES.indexOf(m);return i===11?[a+1,"Enero"]:[a,MESES[i+1]];};

/* ---------------- MOTORES ---------------- */
function calcComprobacion(d){
  const base=d.ingresos||0,compras=(d.pct_compras||0)*base,invF=(d.pct_inventario_final||0)*base;
  const gc=+d.gasto_compra||0;
  const disp=(d.inventario_inicial||0)+compras+gc,costo=disp-invF,bruto=base-costo;
  const util=(d.pct_utilidad_neta||0)*base,gastos=bruto-util,neto=bruto-gastos;
  const superavit=util+(d.superavit_acumulado_anterior||0),invMerc=invF;
  const anc=(d.mejoras||0)+(d.mobiliario||0)-(d.depreciacion||0);
  const obl=d.obligaciones||[],pasC=(d.cuentas_por_pagar||0)+obl.reduce((s,o)=>s+(+o.monto||0),0);
  const capN=(d.capital||0)+(d.reserva_legal||0)+superavit,totPP=pasC+capN;
  const anticipo=totPP-((d.efectivo||0)+(d.cuentas_por_cobrar||0)+invMerc+anc);
  const actC=(d.efectivo||0)+(d.cuentas_por_cobrar||0)+invMerc+anticipo,totA=actC+anc;
  return {...d,tipo:"Comprobación"+(d.fiscal?" (Fiscal)":" (Real)"),compras,inventario_final:invF,costo_de_ventas:costo,
    beneficio_bruto:bruto,gastos_generales:gastos,beneficio_neto:neto,islr:0,reserva:0,utilidad_neta:util,superavit,
    inventario_mercancias:invMerc,total_activo_no_corriente:anc,total_pasivo_corriente:pasC,capital_neto:capN,
    total_pasivo_patrimonio:totPP,anticipo_a_proveedores:anticipo,total_activo_corriente:actC,total_activos:totA,descuadre:totA-totPP};
}
function calcCierre(d){
  const base=d.ingresos||0,gc=+d.gasto_compra||0,disp=(d.inventario_inicial||0)+(d.compras||0)+gc,costo=disp-(d.inventario_final||0);
  const bruto=base-costo,gastos=d.gastos_generales||0,neto=bruto-gastos;
  const util=neto-(d.islr||0)-(d.apartado_reserva_legal||0),superavit=util+(d.resultados_acumulados||0);
  const invMerc=d.inventario_final||0,anc=(d.mejoras||0)+(d.mobiliario||0)-(d.depreciacion||0);
  const obl=d.obligaciones||[],pasC=(d.cuentas_por_pagar||0)+obl.reduce((s,o)=>s+(+o.monto||0),0);
  const capN=(d.capital||0)+(d.reserva_legal||0)+superavit,totPP=pasC+capN;
  const actC=(d.efectivo||0)+(d.cuentas_por_cobrar||0)+invMerc+(d.anticipo||0),totA=actC+anc;
  return {...d,tipo:"Cierre"+(d.fiscal?" (Fiscal)":" (Real)"),compras:d.compras||0,costo_de_ventas:costo,beneficio_bruto:bruto,
    beneficio_neto:neto,utilidad_neta:util,superavit,inventario_mercancias:invMerc,anticipo_a_proveedores:d.anticipo||0,
    total_activo_no_corriente:anc,gastos_generales:gastos,total_pasivo_corriente:pasC,capital_neto:capN,
    total_pasivo_patrimonio:totPP,total_activo_corriente:actC,total_activos:totA,descuadre:totA-totPP};
}
function analizar(c){
  const al=[]; const ing=c.ingresos||0;
  if(Math.abs(c.descuadre)>=1) al.push(["error","El balance NO cuadra: descuadre de "+fmt(c.descuadre)+" Bs."]);
  if(c.anticipo_a_proveedores<0) al.push(["warn","Anticipo a proveedores negativo ("+fmt(c.anticipo_a_proveedores)+"): el pasivo+patrimonio no alcanza a cubrir los activos. Revisa los saldos."]);
  if(c.gastos_generales<0) al.push(["warn","Gastos generales negativos: la utilidad objetivo supera el beneficio bruto. Baja el % de utilidad."]);
  if(ing){
    const mb=c.beneficio_bruto/ing, mu=(c.utilidad_neta||0)/ing;
    if(mb<0.05) al.push(["warn","Margen bruto muy bajo ("+(mb*100).toFixed(1)+"%): el costo de ventas se está comiendo casi todo."]);
    else if(mb>0.6) al.push(["info","Margen bruto alto ("+(mb*100).toFixed(1)+"%): verifica el % de compras, suele rondar el 80%."]);
    if(mu<0) al.push(["warn","Utilidad neta negativa: el período cierra en pérdida."]);
    else if(mu>0&&mu<0.15) al.push(["info","Margen neto del "+(mu*100).toFixed(1)+"%."]);
  }
  if(ing&&c.inventario_inicial===0) al.push(["info","Inventario inicial en 0: recuerda que debe ser el inventario final del año anterior."]);
  if(!al.length||(al.length===1&&al[0][0]==="info")) al.unshift(["ok","Balance cuadrado e indicadores dentro de lo normal."]);
  return al;
}
function fmtNum(ws,cols){ // formato Bs a celdas numéricas
  const r=XLSX.utils.decode_range(ws['!ref']);
  for(let R=r.s.r;R<=r.e.r;R++)for(let C=r.s.c;C<=r.e.c;C++){
    if(cols&&!cols.includes(C))continue;
    const cell=ws[XLSX.utils.encode_cell({r:R,c:C})];
    if(cell&&typeof cell.v==="number"){cell.z='#,##0.00;(#,##0.00)';}
  }
}
function excelBalance(c){
  const E=[["INGRESOS:",""],["  Ingresos por Ventas",c.ingresos],["",""],
    ["COSTO DE VENTAS:",""],["  Inventario Inicial",c.inventario_inicial],["  Más: Compras",c.compras],
    ["  Más: Gastos y fletes en compras",+c.gasto_compra||0],
    ["  Menos: Inventario Final",c.inventario_final],["  Costo de Ventas",c.costo_de_ventas],["",""],
    ["BENEFICIO BRUTO EN VENTAS",c.beneficio_bruto],["",""],["GASTOS DE OPERACIONES:",""],["  Gastos Generales",c.gastos_generales],
    ["BENEFICIO NETO EN EL PERÍODO",c.beneficio_neto],["",""],["ISLR",c.islr||0],["APARTADO RESERVA LEGAL",c.reserva||c.apartado_reserva_legal||0],
    ["UTILIDAD NETA DEL EJERCICIO",c.utilidad_neta],["MÁS: RESULTADOS ACUMULADOS AÑOS ANT.",c.superavit_acumulado_anterior||c.resultados_acumulados||0],["SUPERÁVIT",c.superavit]];
  const S=[["ACTIVO",""],["ACTIVO CORRIENTE:",""],["  Efectivo en Caja y Bancos",c.efectivo||0],["  Cuentas por Cobrar",c.cuentas_por_cobrar||0],
    ["  Inventario de Mercancías",c.inventario_mercancias],["  Anticipo a Proveedores",c.anticipo_a_proveedores],
    ["  Total Activo Corriente",c.total_activo_corriente],["",""],["ACTIVO NO CORRIENTE:",""],["  Mejoras e Instalaciones",c.mejoras||0],
    ["  Mobiliarios y Equipos",c.mobiliario||0],["  Menos: Depreciación Acumulada",c.depreciacion||0],["  Total Activo No Corriente",c.total_activo_no_corriente],
    ["TOTAL ACTIVOS",c.total_activos],["",""],["PASIVO Y PATRIMONIO",""],["PASIVO CORRIENTE:",""],
    ...(c.obligaciones||[]).map(o=>["  "+o.nombre,+o.monto||0]),["  Cuentas y Efectos por Pagar",c.cuentas_por_pagar||0],
    ["  Total Pasivo Corriente",c.total_pasivo_corriente],["",""],["PATRIMONIO:",""],["  Capital",c.capital||0],["  Reserva Legal",c.reserva_legal||0],
    ["  Superávit",c.superavit],["  Capital Neto",c.capital_neto],["TOTAL PASIVO Y PATRIMONIO",c.total_pasivo_patrimonio]];
  const filas=Math.max(E.length,S.length);
  const aoa=[[c.empresa,"","","",""],["ESTADO DE RESULTADOS","","","ESTADO DE SITUACIÓN FINANCIERA",""],
    [`Período ${c.periodo||c.fecha_cierre||""}`,"","",`AL ${c.fecha_cierre||""}   ·   RIF ${c.rif||""}`,""],["","","","",""]];
  for(let i=0;i<filas;i++){const e=E[i]||["",""],s=S[i]||["",""];aoa.push([e[0],e[1],"",s[0],s[1]]);}
  aoa.push(["","","","",""],["CUADRE (debe ser 0)",c.descuadre,"",c.presidente||"",""],["","","","PRESIDENTE",""]);
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=[{wch:36},{wch:20},{wch:3},{wch:38},{wch:20}];
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:4}}];
  fmtNum(ws,[1,4]);
  const inf=[["INFORME DE COMPILACIÓN DEL CONTADOR PÚBLICO INDEPENDIENTE"],[],
    ["Señores"],[c.empresa],["RIF: "+(c.rif||"")],["Presente.-"],[],
    ["He compilado el Estado de Situación Financiera de "+c.empresa+" al "+(c.fecha_cierre||"")+" y el"],
    ["correspondiente Estado de Resultados por el período "+(c.periodo||c.fecha_cierre||"")+", expresados en "+(c.fiscal?"bolívares (cifras fiscales)":"dólares (cifras reales)")+"."],[],
    ["La compilación se limita a presentar, en forma de estados financieros, información que es"],
    ["representación de la administración. No he auditado ni revisado dichos estados financieros y,"],
    ["en consecuencia, no expreso opinión ni ninguna otra forma de seguridad sobre ellos."],[],
    ["Porlamar, "+(c.fecha_cierre||"")],[],[],
    ["________________________________"],["LCDA. ROSELYS SALAZAR"],["CONTADOR PÚBLICO COLEGIADO"],["C.P.C N° 78.838"]];
  const wsi=XLSX.utils.aoa_to_sheet(inf); wsi['!cols']=[{wch:84}];
  const cap=+c.capital||0, rl=+c.reserva_legal||0, apRes=+c.reserva||+c.apartado_reserva_legal||0;
  const supAnt=+c.superavit_acumulado_anterior||+c.resultados_acumulados||0, utl=+c.utilidad_neta||0, sup=+c.superavit||0;
  const vp=[["ESTADO DE VARIACIÓN EN LAS CUENTAS DE PATRIMONIO","","","",""],[`AL ${c.fecha_cierre||""}`,"","","",""],["","Capital","Reserva Legal","Superávit","Total"],
    ["Saldos al inicio del período",cap,rl-apRes,supAnt,cap+(rl-apRes)+supAnt],["Resultado del período",0,apRes,utl,apRes+utl],
    ["Saldos al cierre del período",cap,rl,sup,cap+rl+sup]];
  const wsv=XLSX.utils.aoa_to_sheet(vp); wsv['!cols']=[{wch:34},{wch:18},{wch:16},{wch:20},{wch:20}]; wsv['!merges']=[{s:{r:0,c:0},e:{r:0,c:4}}]; fmtNum(wsv,[1,2,3,4]);
  const m=c.fiscal?"Bs.":"$";
  const notas=[["NOTAS A LOS ESTADOS FINANCIEROS"],[`${c.empresa} — AL ${c.fecha_cierre||""}`],[],
    ["NOTA 1 — CONSTITUCIÓN Y ACTIVIDAD"],[`${c.empresa}, RIF ${c.rif||""}, domiciliada en Isla de Margarita, estado Nueva Esparta, dedicada a la actividad comercial.`],[`Los estados financieros se expresan en ${c.fiscal?"bolívares (cifras fiscales)":"dólares (cifras reales)"}.`],[],
    ["NOTA 2 — EFECTIVO EN CAJA Y BANCOS"],[`Disponibilidad en caja y cuentas bancarias por ${m} ${fmt(c.efectivo||0)}.`],[],
    ["NOTA 3 — MOBILIARIO Y EQUIPOS"],[`Se registran al costo de adquisición ${m} ${fmt(c.mobiliario||0)}, neto de depreciación acumulada de ${m} ${fmt(Math.abs(c.depreciacion||0))}.`],[],
    ["NOTA 4 — MEJORAS E INSTALACIONES"],[`Mejoras e instalaciones por ${m} ${fmt(c.mejoras||0)}.`],[],
    ["NOTA 5 — CAPITAL"],[`Capital social ${m} ${fmt(cap)}; reserva legal ${m} ${fmt(rl)}.`],[],
    ["NOTA 6 — GASTOS"],[`Gastos generales del período ${m} ${fmt(c.gastos_generales||0)}.`],[],
    ["NOTA 7 — POLÍTICA CONTABLE"],["Los estados financieros se preparan sobre la base del costo histórico, conforme a principios de contabilidad de aceptación general."],["La administración es responsable de la preparación y presentación de la información."]];
  const wsn=XLSX.utils.aoa_to_sheet(notas); wsn['!cols']=[{wch:100}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,wsi,"Informe");
  XLSX.utils.book_append_sheet(wb,ws,((c.fiscal?"Fiscal Bs. ":"Real ")+(c.fecha_cierre||"")).slice(0,31));
  XLSX.utils.book_append_sheet(wb,wsv,"Variación Patrimonial");
  XLSX.utils.book_append_sheet(wb,wsn,"Notas");
  XLSX.writeFile(wb,`${c.tipo.replace(/\W+/g,"_")}_${slug(c.empresa)}_${c.fecha_cierre||""}.xlsx`);
}
function calcPersona(d,tasa){
  const props=d.propiedades||[],tp=props.reduce((s,p)=>s+(+p.valor||0),0);
  const activos=tp+(+d.cuentas_bancarias||0)+(+d.cuentas_por_cobrar||0);
  const prest=d.prestamos||[],pasivos=prest.reduce((s,p)=>s+(+p.monto||0),0)+(+d.cuentas_por_pagar||0);
  return {...d,activos,pasivos,patrimonio:activos-pasivos,tasa};
}
function excelPersona(c){
  const d=(v)=>c.tasa?v/c.tasa:0;
  const aoa=[[`Balance Personal — ${c.nombre}`,"",""],[`Cédula ${c.cedula||""}   ·   AL ${c.fecha_cierre||""}   ·   Tasa BCV ${fmt(c.tasa)}`,"",""],["","",""],["ACTIVO","Bs.","US$"],
    ...(c.propiedades||[]).map(p=>["  "+p.nombre,+p.valor||0,d(+p.valor||0)]),["  Cuentas bancarias",+c.cuentas_bancarias||0,d(+c.cuentas_bancarias||0)],
    ["  Cuentas por cobrar",+c.cuentas_por_cobrar||0,d(+c.cuentas_por_cobrar||0)],["TOTAL ACTIVO",c.activos,d(c.activos)],["","",""],["PASIVO","Bs.","US$"],
    ...(c.prestamos||[]).map(p=>["  "+p.nombre,+p.monto||0,d(+p.monto||0)]),["  Cuentas por pagar",+c.cuentas_por_pagar||0,d(+c.cuentas_por_pagar||0)],
    ["TOTAL PASIVO",c.pasivos,d(c.pasivos)],["","",""],["PATRIMONIO NETO",c.patrimonio,d(c.patrimonio)]];
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=[{wch:38},{wch:22},{wch:18}]; ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}];
  fmtNum(ws,[1,2]);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Balance Personal");
  XLSX.writeFile(wb,`Balance_Personal_${slug(c.nombre)}.xlsx`);
}
function serieDesag(e,ano,mes,totalBs,totalU,tasa,proy){
  const make=(A,M,bs,u,p)=>{const o={ano:A,mes:M,proy:p,bs:{},usd:{},u:{}};e.rubros.forEach(r=>{o.bs[r]=bs*(e.ventas?.[r]??1/e.rubros.length);o.u[r]=u*(e.unidades?.[r]??1/e.rubros.length);o.usd[r]=tasa?o.bs[r]/tasa:0;});return o;};
  const regs=[make(ano,mes,totalBs,totalU,false)];
  if(proy){let bs=totalBs,u=totalU,A=ano,M=mes;const cm=proy.crecimiento||0,meta=proy.meta||1,n=proy.meses||12;
    for(let i=0;i<n;i++){[A,M]=sig(A,M);bs*=(1+cm);u*=(1+cm);regs.push(make(A,M,bs*meta,u*meta,true));}}
  return regs;
}
function excelDesag(empresa,e,regs){
  const head=["Año","Mes"];e.rubros.forEach(r=>head.push(r+" Bs"));head.push("Total Bs");e.rubros.forEach(r=>head.push(r+" $"));head.push("Total $");e.rubros.forEach(r=>head.push(r+" "+e.unidad));head.push("Total "+e.unidad);
  const fila=(g)=>{const tb=e.rubros.reduce((s,r)=>s+g.bs[r],0),td=e.rubros.reduce((s,r)=>s+g.usd[r],0),tu=e.rubros.reduce((s,r)=>s+g.u[r],0);return [g.ano,g.mes,...e.rubros.map(r=>g.bs[r]),tb,...e.rubros.map(r=>g.usd[r]),td,...e.rubros.map(r=>g.u[r]),tu];};
  const reales=regs.filter(g=>!g.proy), proy=regs.filter(g=>g.proy);
  const aoa=[[`${empresa} — Desagregación de las Ventas y Unidades`],[],head,...reales.map(fila)];
  if(proy.length){aoa.push([],["PROYECCIÓN (metas para el banco)"],...proy.map(fila));}
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  const nc=head.length; ws['!cols']=Array.from({length:nc},(_,i)=>({wch:i<2?10:16}));
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:nc-1}}];
  fmtNum(ws,Array.from({length:nc-2},(_,i)=>i+2));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Detallado");
  XLSX.writeFile(wb,`Desagregacion_${slug(empresa)}.xlsx`);
}
function descargarInforme(d){
  const hoy=new Date().toLocaleDateString("es-VE",{day:"numeric",month:"long",year:"numeric"});
  const persona=!!d.cedula,idTxt=persona?`cédula de identidad N° ${d.cedula}`:`RIF N° ${d.rif||""}`;
  const cuerpo=persona?`hace constar que el(la) ciudadano(a) <b>${d.nombre}</b>, ${idTxt}, percibe ingresos mensuales${d.ingreso_mensual?` por <b>Bs. ${fmt(d.ingreso_mensual)}</b>`:""} producto de sus actividades económicas, al cierre del período terminado el ${d.fecha_cierre||hoy}.`:`hace constar que la empresa <b>${d.nombre}</b>, ${idTxt}, mantiene su contabilidad al día al cierre del período terminado el ${d.fecha_cierre||hoy}.`;
  const html=`<html><head><meta charset="utf-8"></head><body style="font-family:Arial;max-width:680px;margin:40px auto;line-height:1.6"><div style="text-align:center;font-weight:bold">LCDA. ROSELYS SALAZAR<br><span style="font-weight:normal">CONTADOR PÚBLICO COLEGIADO · C.P.C N° 78.838</span></div><hr><p style="text-align:right">Porlamar, ${hoy}</p><p><b>Señores:</b><br><b>${d.dirigido_a||""}</b><br>Presente.-</p><p style="text-align:center"><b>${persona?"CERTIFICACIÓN DE INGRESOS":"CONSTANCIA"}</b></p><p style="text-align:justify">Quien suscribe, LCDA. ROSELYS SALAZAR, Contador Público Colegiado, C.P.C N° 78.838, ${cuerpo} La presente se expide para <b>${d.tramite||""}</b>.</p><br><br><div style="text-align:center">_____________________________<br><b>LCDA. ROSELYS SALAZAR</b><br>C.P.C N° 78.838</div></body></html>`;
  const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([html],{type:"application/msword"}));a.download=`${persona?"Certificacion":"Constancia"}_${slug(d.nombre)}.doc`;a.click();
}

/* ---------------- deshacer / rehacer ---------------- */
function useUndoable(initial){
  const [h,setH]=useState({past:[],present:initial,future:[]});
  const set=(u)=>setH(s=>{const n=typeof u==="function"?u(s.present):u;if(n===s.present)return s;return {past:[...s.past,s.present].slice(-60),present:n,future:[]};});
  const undo=()=>setH(s=>s.past.length?{past:s.past.slice(0,-1),present:s.past[s.past.length-1],future:[s.present,...s.future]}:s);
  const redo=()=>setH(s=>s.future.length?{past:[...s.past,s.present],present:s.future[0],future:s.future.slice(1)}:s);
  return [h.present,set,undo,redo,h.past.length>0,h.future.length>0];
}
function UndoBar({undo,redo,canU,canR}){
  useEffect(()=>{const k=(e)=>{const z=(e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="z";
    if(z&&!e.shiftKey){e.preventDefault();undo();}else if(((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="y")||(z&&e.shiftKey)){e.preventDefault();redo();}};
    window.addEventListener("keydown",k);return ()=>window.removeEventListener("keydown",k);},[undo,redo]);
  return (<div className="undobar"><button disabled={!canU} onClick={undo} title="Deshacer (Ctrl+Z)"><Undo2 size={15}/></button><button disabled={!canR} onClick={redo} title="Rehacer (Ctrl+Y / Ctrl+Shift+Z)"><Redo2 size={15}/></button></div>);
}

/* ---------------- UI helpers ---------------- */
const soloNum=(s)=>String(s??"").replace(/\D/g,"");
const conPuntos=(s)=>{const n=soloNum(s);return n.replace(/\B(?=(\d{3})+(?!\d))/g,".");};
const isoToVe=(iso)=>{if(!iso)return"";const [y,m,d]=iso.split("-");return `${d}-${m}-${y}`;};
const veToIso=(ve)=>{if(!ve)return"";const p=String(ve).split(/[-/]/);if(p.length!==3)return"";const [d,m,y]=p;return `${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`;};
const Field=({label,value,onChange,type="number",options,placeholder,format})=>{
  const inputType=format==="date"?"date":(format?"text":(type==="number"?"number":"text"));
  const shown=format==="date"?veToIso(value):(format==="id"?conPuntos(value):(value??""));
  const handle=(e)=>{const raw=e.target.value;onChange(format==="date"?isoToVe(raw):(format==="id"?conPuntos(raw):raw));};
  return (<label className="fld"><span>{label}</span>
    {options?(<select value={value??""} onChange={e=>onChange(e.target.value)}>{options.map(o=>(<option key={o} value={o}>{o}</option>))}</select>)
    :(<input type={inputType} inputMode={format==="id"?"numeric":undefined} value={shown} placeholder={placeholder} onChange={handle}/>)}
  </label>);
};
const Badge=({ok})=> ok
  ? <span className="badge ok"><CheckCircle2 size={15}/> Cuadra</span>
  : <span className="badge no"><AlertTriangle size={15}/> Revisar</span>;
const Money=({v})=><span className="mono">{fmt(v)}</span>;
const Linea=({l,v,b,t})=>(<div className={"ln"+(b?" b":"")+(t?" t":"")}><span>{l}</span><Money v={v}/></div>);

/* ---------------- VISTAS ---------------- */
function Panel({tasa,bitacora,ir}){
  const e=EMPRESAS["Importadora Canaima, C.A."];
  const pie=e.rubros.map((r,i)=>({name:r,value:e.ventas[r]}));
  const acciones=[["Comprobación",Scale,"comprobacion"],["Cierre",FileSpreadsheet,"cierre"],["Persona natural",User,"persona"],["Desagregación",BarChart3,"desag"],["Informe",FileText,"informe"],["Agente",MessageSquare,"agente"]];
  return (
    <div className="view">
      <div className="cards">
        <div className="card hero"><div className="lbl">Tasa BCV de hoy</div><div className="big mono">{fmt(tasa)}</div><div className="sub">Bs por dólar · variable del sistema</div></div>
        <div className="card"><div className="lbl">Empresas</div><div className="big mono">{Object.keys(EMPRESAS).length}</div><div className="sub">configuradas · escala a ~50</div></div>
        <div className="card"><div className="lbl">Acciones registradas</div><div className="big mono">{bitacora.length}</div><div className="sub">en la memoria del sistema</div></div>
      </div>
      <div className="grid2">
        <div className="panelbox"><h3>Acciones rápidas</h3><div className="qa">
          {acciones.map(([t,Ic,v])=>(<button key={v} className="qbtn" onClick={()=>ir(v)}><Ic size={18}/>{t}</button>))}
        </div></div>
        <div className="panelbox"><h3>Mezcla de rubros — Canaima</h3>
          <ResponsiveContainer width="100%" height={180}>
            <PieChart><Pie data={pie} dataKey="value" nameKey="name" innerRadius={42} outerRadius={70} paddingAngle={2}>
              {pie.map((x,i)=><Cell key={i} fill={COLORES[i]}/>)}</Pie><Tooltip formatter={(v)=>(v*100).toFixed(0)+"%"}/></PieChart>
          </ResponsiveContainer>
          <div className="leg">{pie.map((x,i)=>(<span key={i}><i style={{background:COLORES[i]}}/>{x.name} {(x.value*100).toFixed(0)}%</span>))}</div>
        </div>
      </div>
      <div className="panelbox"><h3>Actividad reciente</h3>
        {bitacora.length===0?<p className="muted">Aún no hay actividad. Genera tu primer documento.</p>
        :bitacora.slice(0,8).map((b,i)=>(<div className="actrow" key={i}><time className="mono">{b.t}</time><span>{b.txt}</span></div>))}
      </div>
    </div>
  );
}

function FormComprobacion({tasa,onDone}){
  const [f,setF,undo,redo,canU,canR]=useUndoable({empresa:"Importadora Canaima, C.A.",fecha_cierre:"28-02-2026",periodo:"01-01-26 al 28-02-26",fiscal:false,
    ingresos:6200000000,inventario_inicial:385774132.44,pct_compras:0.82,pct_inventario_final:0.011,pct_utilidad_neta:0.028,pct_cxc:0.15,pct_cxp:0.6,pct_efectivo:0.12,
    superavit_acumulado_anterior:1513176770.14,capital:22000000,reserva_legal:2200000,mejoras:662084389.89,mobiliario:22000073.58,
    depreciacion:79662609.36,efectivo:900000000,cuentas_por_cobrar:1200000000,cuentas_por_pagar:0,
    obligaciones:[{nombre:"Banesco Panamá",monto:1979744979.61},{nombre:"Banco de Venezuela",monto:236447357.92}]});
  const set=(k)=>(v)=>setF(s=>({...s,[k]:k.startsWith("pct")||["ingresos","inventario_inicial","superavit_acumulado_anterior","capital","reserva_legal","mejoras","mobiliario","depreciacion","efectivo","cuentas_por_cobrar","cuentas_por_pagar"].includes(k)?(v===""?"":+v):v}));
  const c=calcComprobacion(f);
  const autocompletar=()=>{const cc=calcComprobacion(f);setF(s=>({...s,
    cuentas_por_cobrar:Math.round((+f.pct_cxc||0.15)*(+f.ingresos||0)),
    cuentas_por_pagar:Math.round((+f.pct_cxp||0.6)*cc.inventario_final),
    efectivo:Math.round((+f.pct_efectivo||0.12)*(+f.ingresos||0))}));};
  const pie=[["Efectivo",f.efectivo],["Cuentas x cobrar",f.cuentas_por_cobrar],["Inventario",c.inventario_mercancias],["Anticipo",c.anticipo_a_proveedores],["No corriente",c.total_activo_no_corriente]].map(([n,v],i)=>({name:n,value:Math.max(+v,0)}));
  return (
    <div className="view two">
      <div className="form">
        <div className="fh"><h2>Balance de Comprobación</h2><UndoBar undo={undo} redo={redo} canU={canU} canR={canR}/></div>
        <Field label="Empresa" value={f.empresa} onChange={set("empresa")} options={Object.keys(EMPRESAS)} type="text"/>
        <div className="r2"><Field label="Fecha de corte" value={f.fecha_cierre} onChange={set("fecha_cierre")} type="text" format="date"/>
          <Field label="Pista" value={f.fiscal?"Fiscal":"Real"} onChange={v=>setF(s=>({...s,fiscal:v==="Fiscal"}))} options={["Real","Fiscal"]} type="text"/></div>
        <Field label="Ingresos por ventas (del archivo)" value={f.ingresos} onChange={set("ingresos")}/>
        <Field label="Inventario inicial (= final año anterior)" value={f.inventario_inicial} onChange={set("inventario_inicial")}/>
        <div className="r2"><Field label="% Compras" value={f.pct_compras} onChange={set("pct_compras")}/><Field label="% Inventario final" value={f.pct_inventario_final} onChange={set("pct_inventario_final")}/></div>
        <Field label="Gastos y fletes en compras (Bs)" value={f.gasto_compra} onChange={set("gasto_compra")}/>
        <Field label="% Utilidad neta (objetivo)" value={f.pct_utilidad_neta} onChange={set("pct_utilidad_neta")}/>
        <Field label="Superávit acumulado 31-12 anterior" value={f.superavit_acumulado_anterior} onChange={set("superavit_acumulado_anterior")}/>
        <div className="r2"><Field label="Capital" value={f.capital} onChange={set("capital")}/><Field label="Reserva legal" value={f.reserva_legal} onChange={set("reserva_legal")}/></div>
        <div className="r3"><Field label="% Ctas x cobrar (ventas)" value={f.pct_cxc} onChange={set("pct_cxc")}/><Field label="% Ctas x pagar (inv. final)" value={f.pct_cxp} onChange={set("pct_cxp")}/><Field label="% Efectivo (ventas)" value={f.pct_efectivo} onChange={set("pct_efectivo")}/></div>
        <button type="button" className="mini" onClick={autocompletar}>⚡ Autocompletar con la lógica</button>
        <div className="r2"><Field label="Efectivo" value={f.efectivo} onChange={set("efectivo")}/><Field label="Cuentas por cobrar" value={f.cuentas_por_cobrar} onChange={set("cuentas_por_cobrar")}/></div>
        <Field label="Cuentas por pagar (saldo proveedores)" value={f.cuentas_por_pagar} onChange={set("cuentas_por_pagar")}/>
        <div className="r3"><Field label="Mejoras" value={f.mejoras} onChange={set("mejoras")}/><Field label="Mobiliario" value={f.mobiliario} onChange={set("mobiliario")}/><Field label="Depreciación" value={f.depreciacion} onChange={set("depreciacion")}/></div>
        <Lista titulo="Obligaciones bancarias" items={f.obligaciones} campo="monto" onChange={v=>setF(s=>({...s,obligaciones:v}))}/>
      </div>
      <div className="preview">
        <div className="pvhead"><span>{c.tipo}</span><Badge ok={Math.abs(c.descuadre)<1}/></div>
        <div className="stmt">
          <div className="sec">Estado de Resultados</div>
          <Linea l="Ingresos por ventas" v={c.ingresos}/><Linea l="Costo de ventas" v={c.costo_de_ventas}/>
          <Linea l="Beneficio bruto" v={c.beneficio_bruto} b/><Linea l="Gastos generales" v={c.gastos_generales}/>
          <Linea l="ISLR / Reserva" v={0}/><Linea l="Utilidad neta del ejercicio" v={c.utilidad_neta} b/>
          <Linea l="(+) Acumulado anterior" v={c.superavit_acumulado_anterior}/><Linea l="Superávit" v={c.superavit} b/>
          <div className="sec">Situación Financiera</div>
          <Linea l="Total activo corriente" v={c.total_activo_corriente}/>
          <Linea l="Anticipo a proveedores (cuadre)" v={c.anticipo_a_proveedores}/>
          <Linea l="Total activos" v={c.total_activos} b/><Linea l="Total pasivo + patrimonio" v={c.total_pasivo_patrimonio} b/>
          <Linea l="Descuadre" v={c.descuadre} t/>
        </div>
        <ResponsiveContainer width="100%" height={150}>
          <PieChart><Pie data={pie} dataKey="value" nameKey="name" innerRadius={30} outerRadius={58} paddingAngle={2}>{pie.map((x,i)=><Cell key={i} fill={COLORES[i]}/>)}</Pie><Tooltip formatter={(v)=>fmt0(v)}/></PieChart>
        </ResponsiveContainer>
        <div className="alerts">{analizar(c).map((a,i)=>(<div className={"alert a-"+a[0]} key={i}>{a[1]}</div>))}</div>
        <div className="dlrow"><button className="dl" onClick={()=>{excelBalance(c);onDone(`Comprobación ${f.empresa} ${f.fecha_cierre}`,"balance",c);}}><Download size={16}/> Excel</button><button className="dl2" onClick={()=>imprimirEstado(c)}>Imprimir / PDF</button></div>
      </div>
    </div>
  );
}
function Lista({titulo,items,campo,onChange}){
  const upd=(i,k,v)=>{const n=items.map((it,j)=>j===i?{...it,[k]:k===campo?(v===""?"":+v):v}:it);onChange(n);};
  return (<div className="lista"><div className="lhdr">{titulo}<button onClick={()=>onChange([...items,{nombre:"",[campo]:0}])}><Plus size={14}/></button></div>
    {items.map((it,i)=>(<div className="lrow" key={i}>
      <input value={it.nombre} placeholder="Nombre" onChange={e=>upd(i,"nombre",e.target.value)}/>
      <input type="number" value={it[campo]} onChange={e=>upd(i,campo,e.target.value)}/>
      <button onClick={()=>onChange(items.filter((_,j)=>j!==i))}><Trash2 size={13}/></button></div>))}</div>);
}

function FormCierre({onDone}){
  const [f,setF,undo,redo,canU,canR]=useUndoable({empresa:"Importadora Canaima, C.A.",rif:"J-06500004-0",fecha_cierre:"31-12-2025",periodo:"01-01-25 al 31-12-25",fiscal:false,
    ingresos:35637987420.35,inventario_inicial:147104321.67,compras:29273042867.08,inventario_final:385774132.44,gastos_generales:5032083823.75,
    islr:310497068.81,apartado_reserva_legal:1600000,resultados_acumulados:499645411.86,efectivo:2642546966.74,cuentas_por_cobrar:1774771773.53,
    anticipo:0,cuentas_por_pagar:0,capital:22000000,reserva_legal:2200000,mejoras:662084389.89,mobiliario:22000073.58,depreciacion:79662609.36,
    obligaciones:[{nombre:"Banesco Panamá",monto:1979744979.61},{nombre:"Banesco Venezuela",monto:1508350661.66},{nombre:"Banco de Venezuela",monto:236447357.92},{nombre:"Banco Activo",monto:145594957.49}]});
  const num=(k)=>(v)=>setF(s=>({...s,[k]:v===""?"":+v}));
  const c=calcCierre(f);
  return (<div className="view two">
    <div className="form"><div className="fh"><h2>Balance de Cierre Anual</h2><UndoBar undo={undo} redo={redo} canU={canU} canR={canR}/></div>
      <Field label="Empresa" value={f.empresa} onChange={v=>setF(s=>({...s,empresa:v,rif:EMPRESAS[v]?.rif||""}))} options={Object.keys(EMPRESAS)} type="text"/>
      <Field label="Fecha de cierre" value={f.fecha_cierre} onChange={v=>setF(s=>({...s,fecha_cierre:v}))} type="text" format="date"/>
      <Field label="Ingresos" value={f.ingresos} onChange={num("ingresos")}/>
      <div className="r2"><Field label="Inventario inicial" value={f.inventario_inicial} onChange={num("inventario_inicial")}/><Field label="Compras" value={f.compras} onChange={num("compras")}/></div>
      <Field label="Gastos y fletes en compras" value={f.gasto_compra} onChange={num("gasto_compra")}/>
      <div className="r2"><Field label="Inventario final" value={f.inventario_final} onChange={num("inventario_final")}/><Field label="Gastos generales" value={f.gastos_generales} onChange={num("gastos_generales")}/></div>
      <div className="r2"><Field label="ISLR" value={f.islr} onChange={num("islr")}/><Field label="Reserva legal (apartado)" value={f.apartado_reserva_legal} onChange={num("apartado_reserva_legal")}/></div>
      <Field label="Resultados acumulados" value={f.resultados_acumulados} onChange={num("resultados_acumulados")}/>
      <div className="r2"><Field label="Efectivo" value={f.efectivo} onChange={num("efectivo")}/><Field label="Cuentas por cobrar" value={f.cuentas_por_cobrar} onChange={num("cuentas_por_cobrar")}/></div>
      <div className="r3"><Field label="Mejoras" value={f.mejoras} onChange={num("mejoras")}/><Field label="Mobiliario" value={f.mobiliario} onChange={num("mobiliario")}/><Field label="Depreciación" value={f.depreciacion} onChange={num("depreciacion")}/></div>
      <Lista titulo="Obligaciones bancarias" items={f.obligaciones} campo="monto" onChange={v=>setF(s=>({...s,obligaciones:v}))}/>
    </div>
    <div className="preview"><div className="pvhead"><span>{c.tipo}</span><Badge ok={Math.abs(c.descuadre)<1}/></div>
      <div className="stmt"><div className="sec">Resultados</div>
        <Linea l="Ingresos" v={c.ingresos}/><Linea l="Costo de ventas" v={c.costo_de_ventas}/><Linea l="Beneficio bruto" v={c.beneficio_bruto} b/>
        <Linea l="Gastos generales" v={c.gastos_generales}/><Linea l="Utilidad neta" v={c.utilidad_neta} b/><Linea l="Superávit" v={c.superavit} b/>
        <div className="sec">Situación Financiera</div><Linea l="Total activos" v={c.total_activos} b/><Linea l="Total pasivo + patrimonio" v={c.total_pasivo_patrimonio} b/><Linea l="Descuadre" v={c.descuadre} t/></div>
      <div className="alerts">{analizar(c).map((a,i)=>(<div className={"alert a-"+a[0]} key={i}>{a[1]}</div>))}</div>
      <div className="dlrow"><button className="dl" onClick={()=>{excelBalance(c);onDone(`Cierre ${f.empresa} ${f.fecha_cierre}`,"balance",c);}}><Download size={16}/> Excel</button><button className="dl2" onClick={()=>imprimirEstado(c)}>Imprimir / PDF</button></div>
    </div></div>);
}

function FormPersona({tasa,onDone}){
  const [f,setF,undo,redo,canU,canR]=useUndoable({nombre:"Walid Darwiche",cedula:"15.005.017",fecha_cierre:"31-05-2026",cuentas_bancarias:2000000000,cuentas_por_cobrar:0,cuentas_por_pagar:0,
    propiedades:[{nombre:"Casa",valor:80000000000},{nombre:"Vehículo",valor:9000000000}],prestamos:[{nombre:"Préstamo Banesco",monto:15000000000}]});
  const [tasaCorte,setTasaCorte]=useState(tasa);
  useEffect(()=>{(async()=>{
    const hoy=new Date().toISOString().slice(0,10), iso=veToIso(f.fecha_cierre);
    if(!iso||iso===hoy){setTasaCorte(tasa);return;}
    try{const r=await window.storage.get("tasa:"+iso);setTasaCorte(r?JSON.parse(r.value).valor:tasa);}catch(e){setTasaCorte(tasa);}
  })();},[f.fecha_cierre,tasa]);
  const c=calcPersona(f,tasaCorte);
  return (<div className="view two">
    <div className="form"><div className="fh"><h2>Balance Personal</h2><UndoBar undo={undo} redo={redo} canU={canU} canR={canR}/></div>
      <div className="r2"><Field label="Nombre" value={f.nombre} onChange={v=>setF(s=>({...s,nombre:v}))} type="text"/><Field label="Cédula" value={f.cedula} onChange={v=>setF(s=>({...s,cedula:v}))} type="text" format="id"/></div>
      <Field label="Fecha de corte" value={f.fecha_cierre} onChange={v=>setF(s=>({...s,fecha_cierre:v}))} type="text" format="date"/>
      <Lista titulo="Propiedades y bienes" items={f.propiedades} campo="valor" onChange={v=>setF(s=>({...s,propiedades:v}))}/>
      <Field label="Cuentas bancarias" value={f.cuentas_bancarias} onChange={v=>setF(s=>({...s,cuentas_bancarias:v===""?"":+v}))}/>
      <Lista titulo="Préstamos / pasivos" items={f.prestamos} campo="monto" onChange={v=>setF(s=>({...s,prestamos:v}))}/>
    </div>
    <div className="preview"><div className="pvhead"><span>Patrimonio personal</span></div>
      <div className="stmt"><Linea l="Total activos (Bs)" v={c.activos} b/><Linea l="Total pasivos (Bs)" v={c.pasivos}/><Linea l="Patrimonio (Bs)" v={c.patrimonio} b/>
        <Linea l="Patrimonio ($)" v={tasaCorte?c.patrimonio/tasaCorte:0} t/></div>
      <div className="pvsub">Conversión a la tasa del corte: {fmt(tasaCorte)} Bs/$ {veToIso(f.fecha_cierre)===new Date().toISOString().slice(0,10)?"(en vivo)":""}</div>
      <ResponsiveContainer width="100%" height={160}><BarChart data={[{n:"Activos",v:c.activos},{n:"Pasivos",v:c.pasivos},{n:"Patrimonio",v:c.patrimonio}]}><CartesianGrid strokeDasharray="3 3" stroke="#E6E1D4"/><XAxis dataKey="n" fontSize={11}/><YAxis tickFormatter={fmt0} fontSize={9} width={70}/><Tooltip formatter={fmt0}/><Bar dataKey="v" radius={[4,4,0,0]}>{[0,1,2].map(i=><Cell key={i} fill={COLORES[i]}/>)}</Bar></BarChart></ResponsiveContainer>
      <div className="dlrow"><button className="dl" onClick={()=>{excelPersona(c);onDone(`Balance personal ${f.nombre}`,"persona",c);}}><Download size={16}/> Excel</button><button className="dl2" onClick={()=>imprimirPersona(c)}>Imprimir / PDF</button></div>
    </div></div>);
}

function FormDesag({tasa,onDone}){
  const [f,setF,undo,redo,canU,canR]=useUndoable({empresa:"Importadora Canaima, C.A.",mes:"Mayo",ano:2026,total:5000000000,unidades:250000,proy:true,crecimiento:0.02,meta:1.10});
  const e=EMPRESAS[f.empresa];
  const regs=serieDesag(e,f.ano,f.mes,+f.total||0,+f.unidades||0,tasa,f.proy?{meses:12,crecimiento:+f.crecimiento||0,meta:+f.meta||1}:null);
  const base=regs[0];
  const barData=e.rubros.map(r=>({name:r.split(" ")[0],Bs:base.bs[r]}));
  const lineData=regs.map(g=>({mes:g.mes.slice(0,3),Total:e.rubros.reduce((s,r)=>s+g.bs[r],0)}));
  return (<div className="view two">
    <div className="form"><div className="fh"><h2>Desagregación de Ventas</h2><UndoBar undo={undo} redo={redo} canU={canU} canR={canR}/></div>
      <Field label="Empresa" value={f.empresa} onChange={v=>setF(s=>({...s,empresa:v}))} options={Object.keys(EMPRESAS)} type="text"/>
      <div className="r2"><Field label="Mes base" value={f.mes} onChange={v=>setF(s=>({...s,mes:v}))} options={MESES} type="text"/><Field label="Año" value={f.ano} onChange={v=>setF(s=>({...s,ano:+v}))}/></div>
      <Field label="Total ventas (Bs)" value={f.total} onChange={v=>setF(s=>({...s,total:v===""?"":+v}))}/>
      <Field label={"Total unidades ("+e.unidad+")"} value={f.unidades} onChange={v=>setF(s=>({...s,unidades:v===""?"":+v}))}/>
      <label className="chk"><input type="checkbox" checked={f.proy} onChange={e=>setF(s=>({...s,proy:e.target.checked}))}/> Proyectar 12 meses (para el banco)</label>
      {f.proy&&<div className="r2"><Field label="Crecimiento mensual" value={f.crecimiento} onChange={v=>setF(s=>({...s,crecimiento:v===""?"":+v}))}/><Field label="Factor meta (>1 supera)" value={f.meta} onChange={v=>setF(s=>({...s,meta:v===""?"":+v}))}/></div>}
    </div>
    <div className="preview"><div className="pvhead"><span>{f.empresa.split(",")[0]} · {f.mes} {f.ano}</span></div>
      <div className="pvsub">Reparto por rubro (Bs)</div>
      <ResponsiveContainer width="100%" height={150}><BarChart data={barData}><CartesianGrid strokeDasharray="3 3" stroke="#E6E1D4"/><XAxis dataKey="name" fontSize={11}/><YAxis tickFormatter={fmt0} fontSize={9} width={70}/><Tooltip formatter={fmt0}/><Bar dataKey="Bs" radius={[4,4,0,0]}>{barData.map((x,i)=><Cell key={i} fill={COLORES[i]}/>)}</Bar></BarChart></ResponsiveContainer>
      {f.proy&&<><div className="pvsub"><TrendingUp size={13}/> Proyección 12 meses</div>
      <ResponsiveContainer width="100%" height={140}><LineChart data={lineData}><CartesianGrid strokeDasharray="3 3" stroke="#E6E1D4"/><XAxis dataKey="mes" fontSize={10}/><YAxis tickFormatter={fmt0} fontSize={9} width={70}/><Tooltip formatter={fmt0}/><Line dataKey="Total" stroke="#C8892B" strokeWidth={2} dot={false}/></LineChart></ResponsiveContainer></>}
      <button className="dl" onClick={()=>{excelDesag(f.empresa,e,regs);onDone(`Desagregación ${f.empresa} ${f.mes}`,"desag",{empresa:f.empresa,regs});}}><Download size={16}/> Descargar Excel</button>
    </div></div>);
}

function FormInforme({onDone}){
  const [f,setF]=useState({tipo:"Empresa",nombre:"Importadora Canaima, C.A.",rif:"J-06500004-0",cedula:"",dirigido_a:"Banco Banesco, C.A.",tramite:"actualización de expediente",fecha_cierre:"31 de mayo de 2026",ingreso_mensual:0});
  const set=(k)=>(v)=>setF(s=>({...s,[k]:v}));
  const persona=f.tipo==="Persona";
  const enviar=()=>{const payload={...f,cedula:persona?f.cedula:"",rif:persona?"":f.rif};descargarInforme(payload);onDone(`Informe ${f.nombre}`,"informe",payload);};
  return (<div className="view two">
    <div className="form"><h2>Informes</h2>
      <Field label="Tipo" value={f.tipo} onChange={set("tipo")} options={["Empresa","Persona"]} type="text"/>
      <Field label="Nombre" value={f.nombre} onChange={set("nombre")} type="text"/>
      {persona?<Field label="Cédula" value={f.cedula} onChange={set("cedula")} type="text" format="id"/>:<Field label="RIF (solo números)" value={f.rif} onChange={set("rif")} type="text" format="id"/>}
      <Field label="Dirigido a" value={f.dirigido_a} onChange={set("dirigido_a")} type="text"/>
      <Field label="Trámite" value={f.tramite} onChange={set("tramite")} options={["solicitud de crédito","trámites correspondientes","actualización de expediente","apertura de cuenta"]} type="text"/>
      <Field label="Fecha de cierre" value={f.fecha_cierre} onChange={set("fecha_cierre")} type="text"/>
      {persona&&<Field label="Ingreso mensual (Bs)" value={f.ingreso_mensual} onChange={v=>setF(s=>({...s,ingreso_mensual:+v||0}))}/>}
    </div>
    <div className="preview carta"><div className="pvhead"><span>{persona?"Certificación de ingresos":"Constancia"}</span></div>
      <div className="doc">
        <div className="ctr b">LCDA. ROSELYS SALAZAR</div><div className="ctr">CONTADOR PÚBLICO COLEGIADO · C.P.C N° 78.838</div><hr/>
        <p className="rgt">Porlamar, _______</p><p><b>Señores:</b><br/><b>{f.dirigido_a}</b><br/>Presente.-</p>
        <p className="ctr b">{persona?"CERTIFICACIÓN DE INGRESOS":"CONSTANCIA"}</p>
        <p className="just">Quien suscribe hace constar que {persona?"el(la) ciudadano(a)":"la empresa"} <b>{f.nombre}</b>, {persona?`cédula N° ${f.cedula}`:`RIF N° ${f.rif}`}, {persona?`percibe ingresos mensuales${f.ingreso_mensual?` por Bs. ${fmt(f.ingreso_mensual)}`:""}`:"mantiene su contabilidad al día"} al cierre del {f.fecha_cierre}. La presente se expide para <b>{f.tramite}</b>.</p>
      </div>
      <button className="dl" onClick={enviar}><Download size={16}/> Descargar documento</button>
    </div></div>);
}

function imprimirEstado(c){
  const fila=(l,v,b)=>`<tr class="${b?'b':''}"><td>${l}</td><td class="n">${v==null?'':fmt(v)}</td></tr>`;
  const eerr=[['Ingresos por Ventas',c.ingresos],['Costo de Ventas',c.costo_de_ventas],['Beneficio Bruto en Ventas',c.beneficio_bruto,1],['Gastos Generales',c.gastos_generales],['ISLR',c.islr||0],['Reserva Legal',c.reserva||c.apartado_reserva_legal||0],['Utilidad Neta del Ejercicio',c.utilidad_neta,1],['(+) Resultados Acumulados',c.superavit_acumulado_anterior||c.resultados_acumulados||0],['Superávit',c.superavit,1]];
  const esf=[['Total Activo Corriente',c.total_activo_corriente],['Anticipo a Proveedores',c.anticipo_a_proveedores],['Total Activo No Corriente',c.total_activo_no_corriente],['TOTAL ACTIVOS',c.total_activos,1],['Total Pasivo Corriente',c.total_pasivo_corriente],['Capital Neto',c.capital_neto],['TOTAL PASIVO Y PATRIMONIO',c.total_pasivo_patrimonio,1]];
  const w=window.open('','_blank'); if(!w) return;
  w.document.write(`<html><head><meta charset=utf-8><title>${c.empresa}</title><style>body{font-family:Georgia,serif;color:#16222B;max-width:840px;margin:28px auto;padding:0 26px}.hd{text-align:center;border-bottom:2px solid #C9A24B;padding-bottom:10px;margin-bottom:16px}.hd b{font-size:13px}h1{font-size:19px;margin:0}.sub{text-align:center;color:#6b6450;font-size:12px;margin:2px 0 16px}table{width:100%;border-collapse:collapse;font-family:Arial;font-size:12px}td{padding:5px 6px;border-bottom:1px dotted #d8d2c2}.n{text-align:right;font-variant-numeric:tabular-nums}tr.b td{font-weight:700;border-bottom:1px solid #0C1822}.cols{display:flex;gap:34px}.cols>div{flex:1}.sec{font-weight:700;color:#0C1822;border-bottom:2px solid #C9A24B;margin:12px 0 4px;font-size:12px;text-transform:uppercase;letter-spacing:1px}.firma{text-align:center;margin-top:46px;font-size:12px}.cuadre{text-align:center;margin-top:16px;padding:9px;background:#E6F3ED;color:#0C6A5A;font-weight:700;border-radius:7px}</style></head><body>
  <div class="hd"><b>LCDA. ROSELYS SALAZAR</b><br>CONTADOR PÚBLICO COLEGIADO · C.P.C N° 78.838</div>
  <h1>${c.empresa}</h1><div class="sub">${c.tipo} · ${c.fecha_cierre||''} · RIF ${c.rif||''}</div>
  <div class="cols"><div><div class="sec">Estado de Resultados</div><table>${eerr.map(r=>fila(r[0],r[1],r[2])).join('')}</table></div>
  <div><div class="sec">Estado de Situación Financiera</div><table>${esf.map(r=>fila(r[0],r[1],r[2])).join('')}</table></div></div>
  <div class="cuadre">${Math.abs(c.descuadre)<1?'Balance cuadrado ✓':'Descuadre: '+fmt(c.descuadre)}</div>
  <div class="firma">________________________________<br><b>LCDA. ROSELYS SALAZAR</b><br>C.P.C N° 78.838</div>
  <script>window.onload=()=>setTimeout(()=>window.print(),300)<\/script></body></html>`);
  w.document.close();
}
function imprimirPersona(c){
  const d=(v)=>c.tasa?v/c.tasa:0;
  const fila=(l,bs,b)=>`<tr class="${b?'b':''}"><td>${l}</td><td class="n">${fmt(bs)}</td><td class="n">${fmt(d(bs))}</td></tr>`;
  const w=window.open('','_blank'); if(!w) return;
  w.document.write(`<html><head><meta charset=utf-8><title>${c.nombre}</title><style>body{font-family:Georgia,serif;color:#16222B;max-width:680px;margin:28px auto;padding:0 26px}.hd{text-align:center;border-bottom:2px solid #C9A24B;padding-bottom:10px;margin-bottom:16px;font-size:13px}h1{font-size:19px;text-align:center;margin:0}.sub{text-align:center;color:#6b6450;font-size:12px;margin:2px 0 16px}table{width:100%;border-collapse:collapse;font-family:Arial;font-size:12px}th{text-align:right;font-size:10px;color:#8a8064;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid #C9A24B}th:first-child{text-align:left}td{padding:5px 6px;border-bottom:1px dotted #d8d2c2}.n{text-align:right;font-variant-numeric:tabular-nums}tr.b td{font-weight:700;border-bottom:1px solid #0C1822}.sec{font-weight:700;color:#0C1822;margin:14px 0 4px;font-size:12px;text-transform:uppercase;letter-spacing:1px}.firma{text-align:center;margin-top:46px;font-size:12px}</style></head><body>
  <div class="hd"><b>LCDA. ROSELYS SALAZAR</b><br>CONTADOR PÚBLICO COLEGIADO · C.P.C N° 78.838</div>
  <h1>Balance Personal — ${c.nombre}</h1><div class="sub">Cédula ${c.cedula||''} · AL ${c.fecha_cierre||''} · Tasa BCV ${fmt(c.tasa)}</div>
  <table><tr><th>Cuenta</th><th>Bs.</th><th>US$</th></tr>
  <tr><td class="sec" colspan=3>Activo</td></tr>
  ${(c.propiedades||[]).map(p=>fila(p.nombre,+p.valor||0)).join('')}
  ${fila('Cuentas bancarias',+c.cuentas_bancarias||0)}${fila('Cuentas por cobrar',+c.cuentas_por_cobrar||0)}${fila('TOTAL ACTIVO',c.activos,1)}
  <tr><td class="sec" colspan=3>Pasivo</td></tr>
  ${(c.prestamos||[]).map(p=>fila(p.nombre,+p.monto||0)).join('')}${fila('Cuentas por pagar',+c.cuentas_por_pagar||0)}${fila('TOTAL PASIVO',c.pasivos,1)}
  <tr><td class="sec" colspan=3>Patrimonio</td></tr>${fila('PATRIMONIO NETO',c.patrimonio,1)}</table>
  <div class="firma">________________________________<br><b>LCDA. ROSELYS SALAZAR</b><br>C.P.C N° 78.838</div>
  <script>window.onload=()=>setTimeout(()=>window.print(),300)<\/script></body></html>`);
  w.document.close();
}
const TOOLS=[
  {name:"consultar_tasa_bcv",description:"Tasa BCV guardada.",input_schema:{type:"object",properties:{}}},
  {name:"listar_empresas",description:"Lista empresas.",input_schema:{type:"object",properties:{}}},
  {name:"generar_comprobacion",description:"Balance de comprobación real/fiscal; anticipo cuadra; % en decimal.",input_schema:{type:"object",properties:{empresa:{type:"string"},fecha_cierre:{type:"string"},fiscal:{type:"boolean"},ingresos:{type:"number"},inventario_inicial:{type:"number"},pct_compras:{type:"number"},pct_inventario_final:{type:"number"},pct_utilidad_neta:{type:"number"},superavit_acumulado_anterior:{type:"number"},capital:{type:"number"},reserva_legal:{type:"number"},mejoras:{type:"number"},mobiliario:{type:"number"},depreciacion:{type:"number"},efectivo:{type:"number"},cuentas_por_cobrar:{type:"number"},obligaciones:{type:"array",items:{type:"object",properties:{nombre:{type:"string"},monto:{type:"number"}}}}},required:["empresa","fecha_cierre","ingresos"]}},
  {name:"generar_desagregacion",description:"Desagregación por rubro + $ con proyección opcional (meses,crecimiento,meta).",input_schema:{type:"object",properties:{empresa:{type:"string"},mes:{type:"string"},ano:{type:"number"},total_ventas_bs:{type:"number"},total_unidades:{type:"number"},proyeccion:{type:"object",properties:{meses:{type:"number"},crecimiento:{type:"number"},meta:{type:"number"}}}},required:["empresa","mes","ano"]}},
  {name:"generar_informe",description:"Constancia (rif) o certificación (cedula).",input_schema:{type:"object",properties:{nombre:{type:"string"},rif:{type:"string"},cedula:{type:"string"},dirigido_a:{type:"string"},tramite:{type:"string"},fecha_cierre:{type:"string"}},required:["nombre","dirigido_a","tramite"]}},
];
const SISTEMA="Eres el agente del Sistema Contable Margarita (Isla Margarita, Venezuela). Trabajas en Bs y $ a la tasa BCV. Usa las herramientas para desagregaciones, balances de comprobación e informes. Porcentajes en decimal (0.82=82%). Si falta un dato clave, pídelo breve. Confirma cada resultado (cuadre, archivo). Responde en español, claro.";

function ChatAgente({tasa,onLog}){
  const [msgs,setMsgs]=useState([{role:"assistant",text:"Soy el agente. Dime en lenguaje natural qué necesitas: «desagrega Opus de mayo 2026 por 5000 millones, proyecta 12 meses con meta 10% arriba» o «constancia de Canaima para Banesco por apertura de cuenta»."}]);
  const [txt,setTxt]=useState("");const [load,setLoad]=useState(false);const ref=useRef(null);
  useEffect(()=>{ref.current?.scrollIntoView({behavior:"smooth"});},[msgs,load]);
  const exec=async(name,a)=>{
    const e=EMPRESAS[a.empresa];
    if(name==="consultar_tasa_bcv")return `Tasa BCV: ${fmt(tasa)} Bs/$`;
    if(name==="listar_empresas")return Object.keys(EMPRESAS).join(" · ");
    if(name==="generar_comprobacion"){const c=calcComprobacion({...a,rif:a.rif||e?.rif||""});excelBalance(c);onLog(`Comprobación ${a.empresa}`);return `Listo. Cuadre=${fmt(c.descuadre)} (0=OK). Anticipo (cuadre)=${fmt(c.anticipo_a_proveedores)}. Excel descargado.`;}
    if(name==="generar_desagregacion"){if(!e)return"Empresa no configurada.";const regs=serieDesag(e,a.ano,a.mes,a.total_ventas_bs||0,a.total_unidades||0,tasa,a.proyeccion);excelDesag(a.empresa,e,regs);onLog(`Desagregación ${a.empresa}`);return `Listo. ${regs.length} mes(es). Excel descargado.`;}
    if(name==="generar_informe"){descargarInforme(a);onLog(`Informe ${a.nombre}`);return `Listo. Documento de ${a.nombre} descargado.`;}
    return "Herramienta desconocida.";
  };
  const enviar=async()=>{
    const t=txt.trim();if(!t||load)return;setTxt("");setMsgs(m=>[...m,{role:"user",text:t}]);setLoad(true);
    const api=[];[...msgs,{role:"user",text:t}].forEach(m=>{if(m.role==="user")api.push({role:"user",content:m.text});else if(m.role==="assistant"&&m.text)api.push({role:"assistant",content:m.text});});
    try{for(let i=0;i<6;i++){
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1400,system:SISTEMA,tools:TOOLS,messages:api})});
      const data=await r.json();if(!data.content)throw new Error(data.error?.message||"Sin respuesta");
      api.push({role:"assistant",content:data.content});
      const tx=data.content.filter(b=>b.type==="text").map(b=>b.text).join("\n").trim();if(tx)setMsgs(m=>[...m,{role:"assistant",text:tx}]);
      const us=data.content.filter(b=>b.type==="tool_use");if(!us.length)break;
      const res=[];for(const u of us){const out=await exec(u.name,u.input||{});setMsgs(m=>[...m,{role:"tool",text:`${u.name} → ${out}`}]);res.push({type:"tool_result",tool_use_id:u.id,content:out});}
      api.push({role:"user",content:res});
    }}catch(e){setMsgs(m=>[...m,{role:"assistant",text:"No pude conectar: "+e.message}]);}
    setLoad(false);
  };
  return (<div className="view chatview">
    <div className="chat">{msgs.map((m,i)=>(<div className={`row ${m.role}`} key={i}><div><div className="who">{m.role==="user"?"Tú":m.role==="tool"?"ejecución":"Agente"}</div><div className="bub">{m.text}</div></div></div>))}
      {load&&<div className="row assistant"><div><div className="who">Agente</div><div className="bub"><span className="dots"><i/><i/><i/></span></div></div></div>}<div ref={ref}/></div>
    <div className="bar"><textarea value={txt} placeholder="Escribe una instrucción…  (Enter envía)" onChange={e=>setTxt(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();enviar();}}}/><button onClick={enviar} disabled={load||!txt.trim()}>Enviar</button></div>
  </div>);
}

function Historial({historial,onExport,onImport}){
  const SUB={balance:"Balance",persona:"Balance personal",desag:"Desagregación",informe:"Informe"};
  const rebajar=(h)=>{const p=h.payload;
    if(h.kind==="balance")excelBalance(p);else if(h.kind==="persona")excelPersona(p);
    else if(h.kind==="desag")excelDesag(p.empresa,EMPRESAS[p.empresa],p.regs);else if(h.kind==="informe")descargarInforme(p);};
  const pdf=(h)=>{if(h.kind==="balance")imprimirEstado(h.payload);else if(h.kind==="persona")imprimirPersona(h.payload);};
  return (<div className="view"><div className="panelbox">
    <div className="fh"><h3 style={{margin:0}}>Historial y copia de seguridad</h3>
      <div className="bkp"><button onClick={onExport}><Download size={13}/> Exportar respaldo</button>
      <label className="imp"><Upload size={13}/> Restaurar<input type="file" accept="application/json" hidden onChange={e=>e.target.files[0]&&onImport(e.target.files[0])}/></label></div></div>
    <p className="muted hsub">Cada documento que generas queda guardado con todos sus datos. Puedes volver a generarlo cuando quieras —aunque cierres el navegador— por si necesitas recuperar algo o corregir un error.</p>
    {historial.length===0?<p className="muted">Todavía no has generado documentos. Cuando lo hagas, aparecerán aquí para respaldo.</p>
    :<table className="htbl"><thead><tr><th>Fecha</th><th>Tipo</th><th>Documento</th><th>Acciones</th></tr></thead><tbody>
      {historial.map((h)=>(<tr key={h.id}>
        <td className="mono hf">{h.fecha}</td><td><span className={"htag k-"+h.kind}>{SUB[h.kind]||h.kind}</span></td><td>{h.titulo}</td>
        <td className="hact"><button onClick={()=>rebajar(h)} title="Volver a generar / descargar"><RotateCcw size={14}/></button>
        {(h.kind==="balance"||h.kind==="persona")&&<button onClick={()=>pdf(h)} title="Imprimir PDF"><Printer size={14}/></button>}</td></tr>))}
    </tbody></table>}
  </div></div>);
}

/* ---------------- APP ---------------- */
export default function App(){
  const [vista,setVista]=useState("panel");
  const [tasa,setTasa]=useState(623.02);
  const [bitacora,setBitacora]=useState([]);
  const [historial,setHistorial]=useState([]);
  useEffect(()=>{(async()=>{try{const t=await window.storage.get("tasa_bcv");if(t)setTasa(JSON.parse(t.value).valor);const b=await window.storage.get("bitacora");if(b)setBitacora(JSON.parse(b.value));const h=await window.storage.get("historial");if(h)setHistorial(JSON.parse(h.value));}catch(e){}})();},[]);
  useEffect(()=>{
    const jalar=async()=>{
      const fuentes=[
        async()=>{const r=await fetch("https://ve.dolarapi.com/v1/dolares/oficial");const d=await r.json();return d.promedio||d.price;},
        async()=>{const r=await fetch("https://pydolarve.org/api/v1/dollar?page=bcv");const d=await r.json();return d?.monitors?.bcv?.price||d?.bcv?.price||d?.price;},
      ];
      for(const f of fuentes){try{const v=await f();if(v&&Number(v)>0){setTasa(Number(v));const iso=new Date().toISOString().slice(0,10);try{await window.storage.set("tasa_bcv",JSON.stringify({valor:Number(v),fecha:new Date().toISOString()}));await window.storage.set("tasa:"+iso,JSON.stringify({valor:Number(v)}));}catch(e){}return;}}catch(e){}}
    };
    jalar(); const id=setInterval(jalar,60000); return ()=>clearInterval(id);
  },[]);
  const log=async(txt,kind,payload)=>{
    const ahora=new Date();
    const n=[{t:ahora.toLocaleTimeString("es-VE",{hour:"2-digit",minute:"2-digit"}),txt},...bitacora].slice(0,40);
    setBitacora(n);try{await window.storage.set("bitacora",JSON.stringify(n));}catch(e){}
    if(kind&&payload){
      const e2={id:ahora.getTime(),fecha:ahora.toLocaleString("es-VE",{day:"2-digit",month:"2-digit",year:"2-digit",hour:"2-digit",minute:"2-digit"}),kind,titulo:txt,payload};
      const nh=[e2,...historial].slice(0,80);setHistorial(nh);
      try{await window.storage.set("historial",JSON.stringify(nh));}catch(e){}
    }
  };
  const exportarTodo=()=>{const data={tasa,bitacora,historial,exportado:new Date().toISOString(),sistema:"Contable Margarita"};
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:"application/json"}));a.download="respaldo_contable_margarita.json";a.click();};
  const importarTodo=(file)=>{const r=new FileReader();r.onload=async()=>{try{const d=JSON.parse(r.result);
    if(d.tasa){setTasa(d.tasa);await window.storage.set("tasa_bcv",JSON.stringify({valor:d.tasa,fecha:new Date().toISOString()}));}
    if(d.bitacora){setBitacora(d.bitacora);await window.storage.set("bitacora",JSON.stringify(d.bitacora));}
    if(d.historial){setHistorial(d.historial);await window.storage.set("historial",JSON.stringify(d.historial));}
    }catch(e){}};r.readAsText(file);};
  const nav=[["panel","Panel",LayoutDashboard],["comprobacion","Comprobación",Scale],["cierre","Cierre",FileSpreadsheet],["persona","Persona natural",User],["desag","Desagregación",BarChart3],["informe","Informes",FileText],["agente","Agente",MessageSquare],["historial","Historial",History]];
  const titulos={panel:"Panel",comprobacion:"Balance de Comprobación",cierre:"Balance de Cierre",persona:"Balance Personal",desag:"Desagregación",informe:"Informes",agente:"Agente Contable",historial:"Historial y Copia de Seguridad"};
  return (
    <div className="app">
      <style>{CSS}</style>
      <aside className="sb">
        <div className="brand">Contable<span>Margarita</span></div>
        <nav>{nav.map(([k,t,Ic])=>(<button key={k} className={"nav"+(vista===k?" on":"")} onClick={()=>setVista(k)}><Ic size={17}/>{t}</button>))}</nav>
        <div className="sbtasa"><span>Tasa BCV <i className="live"/>en vivo</span><b className="mono">{fmt(tasa)}</b></div>
      </aside>
      <main>
        <header><h1>{titulos[vista]}</h1><div className="hchip mono">Bs/$ {fmt(tasa)}</div></header>
        <div className="content">
          {vista==="panel"&&<Panel tasa={tasa} bitacora={bitacora} ir={setVista}/>}
          {vista==="comprobacion"&&<FormComprobacion tasa={tasa} onDone={log}/>}
          {vista==="cierre"&&<FormCierre onDone={log}/>}
          {vista==="persona"&&<FormPersona tasa={tasa} onDone={log}/>}
          {vista==="desag"&&<FormDesag tasa={tasa} onDone={log}/>}
          {vista==="informe"&&<FormInforme onDone={log}/>}
          {vista==="agente"&&<ChatAgente tasa={tasa} onLog={log}/>}
          {vista==="historial"&&<Historial historial={historial} onExport={exportarTodo} onImport={importarTodo}/>}
        </div>
      </main>
    </div>
  );
}

const CSS=`
.app{display:grid;grid-template-columns:248px 1fr;height:100vh;margin:0;font-family:ui-sans-serif,system-ui,"Segoe UI",sans-serif;color:#1B2630;background:#F4F0E6}
.app *{box-sizing:border-box}
.mono{font-family:ui-monospace,"SF Mono",Menlo,monospace;font-variant-numeric:tabular-nums;letter-spacing:-.2px}
.serif{font-family:Georgia,"Times New Roman",serif}
/* ---- sidebar ---- */
.sb{background:linear-gradient(180deg,#0C1822 0%,#0B1A20 100%);color:#BCC9CC;display:flex;flex-direction:column;padding:24px 16px;gap:5px;border-right:1px solid #1A2C36;position:relative}
.sb:after{content:"";position:absolute;top:0;right:-1px;width:2px;height:100%;background:linear-gradient(180deg,transparent,#C9A24B55,transparent)}
.brand{font-family:Georgia,serif;font-weight:700;font-size:19px;color:#F3ECDC;padding:4px 10px 18px;line-height:1.1;letter-spacing:.3px}
.brand span{display:block;font-size:13px;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:#C9A24B;margin-top:5px}
.nav{display:flex;align-items:center;gap:12px;background:none;border:none;color:#8FA3A8;font-size:13.5px;padding:11px 13px;border-radius:10px;cursor:pointer;text-align:left;font-family:inherit;transition:all .15s;position:relative}
.nav:hover{background:#13252E;color:#F3ECDC}
.nav.on{background:linear-gradient(90deg,#1A2C2A,#15252C);color:#F3ECDC;font-weight:600}
.nav.on:before{content:"";position:absolute;left:0;top:18%;height:64%;width:3px;border-radius:3px;background:#C9A24B}
.nav.on svg{color:#C9A24B}
.sbtasa{margin-top:auto;background:linear-gradient(135deg,#10241F,#0C1C22);border:1px solid #233E37;border-radius:12px;padding:13px 15px;display:flex;flex-direction:column;box-shadow:inset 0 1px 0 #2a4a40}
.sbtasa span{font-size:10px;letter-spacing:1.6px;text-transform:uppercase;color:#7FA6A0}
.sbtasa b{font-size:23px;color:#E9C877;font-family:ui-monospace,monospace;margin-top:3px}
.live{display:inline-block;width:7px;height:7px;border-radius:50%;background:#48C7AE;margin:0 4px 0 3px;box-shadow:0 0 0 0 #48C7AE;animation:pulse 1.6s infinite}
@keyframes pulse{0%{box-shadow:0 0 0 0 rgba(72,199,174,.6)}70%{box-shadow:0 0 0 6px rgba(72,199,174,0)}100%{box-shadow:0 0 0 0 rgba(72,199,174,0)}}
/* ---- shell ---- */
main{display:flex;flex-direction:column;height:100vh;overflow:hidden}
header{display:flex;justify-content:space-between;align-items:center;padding:18px 32px;border-bottom:1px solid #E4DDC9;background:#FBF8F0}
header h1{margin:0;font-family:Georgia,serif;font-size:21px;font-weight:700;letter-spacing:.2px;color:#16222B}
.hchip{font-size:12px;background:#0C1822;color:#E9C877;padding:6px 13px;border-radius:30px;border:1px solid #C9A24B44;letter-spacing:.3px}
.content{flex:1;overflow:auto;padding:28px 32px;background:radial-gradient(120% 60% at 50% 0%,#F8F4EB,#F1ECDF)}
.view{max-width:1120px;margin:0 auto;animation:rise .45s cubic-bezier(.22,1,.36,1) both}
@keyframes rise{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
.dlrow{display:flex;gap:11px;margin-top:16px}.dlrow .dl{margin-top:0;flex:1}
.dl2{flex:1;display:flex;align-items:center;justify-content:center;gap:8px;background:#0C1822;color:#E9C877;border:1px solid #C9A24B66;border-radius:11px;padding:13px;font-weight:700;font-size:14px;cursor:pointer;font-family:inherit;letter-spacing:.3px;transition:all .15s}
.dl2:hover{background:#13242e;box-shadow:0 6px 18px rgba(12,24,34,.3);transform:translateY(-1px)}
.card,.panelbox{transition:transform .16s ease,box-shadow .16s ease}
.card:hover,.panelbox:hover{transform:translateY(-2px);box-shadow:0 10px 26px rgba(20,30,40,.10)}
/* ---- cards ---- */
.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-bottom:20px}
.card{background:#fff;border:1px solid #ECE5D3;border-radius:16px;padding:20px 22px;box-shadow:0 2px 14px rgba(20,30,40,.05);position:relative;overflow:hidden}
.card:not(.hero):before{content:"";position:absolute;top:0;left:22px;right:22px;height:2px;background:linear-gradient(90deg,#C9A24B,transparent)}
.card.hero{background:linear-gradient(135deg,#0C1822 0%,#123F38 100%);color:#fff;border:none;box-shadow:0 8px 28px rgba(12,24,34,.25)}
.card.hero:after{content:"";position:absolute;right:-40px;top:-40px;width:140px;height:140px;border-radius:50%;background:radial-gradient(#C9A24B33,transparent 70%)}
.card .lbl{font-size:10.5px;letter-spacing:1.6px;text-transform:uppercase;color:#9B9277}
.card.hero .lbl{color:#9FC6BD}
.card .big{font-family:Georgia,serif;font-size:34px;font-weight:700;margin:8px 0 2px;color:#16222B}
.card.hero .big{font-family:ui-monospace,monospace;color:#E9C877;font-weight:600}
.card .sub{font-size:11.5px;color:#A39A82}.card.hero .sub{color:#9FC6BD}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:20px}
.panelbox{background:#fff;border:1px solid #ECE5D3;border-radius:16px;padding:20px 22px;box-shadow:0 2px 14px rgba(20,30,40,.04)}
.panelbox h3{margin:0 0 16px;font-family:Georgia,serif;font-size:15px;font-weight:700;color:#1B2630}
.qa{display:grid;grid-template-columns:1fr 1fr;gap:11px}
.qbtn{display:flex;align-items:center;gap:10px;background:#FAF7EF;border:1px solid #EAE2CE;border-radius:11px;padding:13px;font-size:13px;cursor:pointer;font-family:inherit;color:#2A3641;transition:all .15s}
.qbtn:hover{background:#0C1822;color:#fff;border-color:#0C1822;transform:translateY(-1px);box-shadow:0 4px 14px rgba(12,24,34,.18)}
.qbtn:hover svg{color:#E9C877}
.leg{display:flex;flex-wrap:wrap;gap:11px;margin-top:10px;font-size:11.5px;color:#5C5443}.leg span{display:flex;align-items:center;gap:5px}.leg i{width:10px;height:10px;border-radius:3px;display:inline-block}
.actrow{display:flex;gap:13px;padding:8px 0;border-bottom:1px solid #F1ECDD;font-size:12.5px}.actrow time{color:#A39A82;flex:none;font-family:ui-monospace,monospace}
.muted{color:#A39A82;font-size:13px}
.hsub{margin:-6px 0 16px}
.htbl{width:100%;border-collapse:collapse;font-size:13px}
.htbl th{text-align:left;font-size:10px;letter-spacing:1px;text-transform:uppercase;color:#9B9277;border-bottom:1px solid #C9A24B;padding:7px 8px}
.htbl td{padding:9px 8px;border-bottom:1px solid #F1ECDD;vertical-align:middle}
.hf{color:#8A8064;white-space:nowrap}
.htag{font-size:10.5px;font-weight:700;padding:3px 9px;border-radius:20px;background:#EEE7D6;color:#6b6450}
.htag.k-balance{background:#E6F3ED;color:#0C6A5A}.htag.k-persona{background:#EDE7F6;color:#6A4BA0}.htag.k-desag{background:#FBF1DD;color:#9A7415}.htag.k-informe{background:#E7EEF5;color:#3E6B8C}
.hact{display:flex;gap:7px}
.hact button{background:#0C1822;color:#E9C877;border:1px solid #C9A24B55;border-radius:8px;width:31px;height:31px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center}
.hact button:hover{background:#13242e}
/* ---- forms / preview ---- */
.two{display:grid;grid-template-columns:minmax(330px,1fr) minmax(370px,1.05fr);gap:24px;max-width:1140px;margin:0 auto}
.form{background:#fff;border:1px solid #ECE5D3;border-radius:16px;padding:22px 24px;box-shadow:0 2px 14px rgba(20,30,40,.04)}
.form h2{margin:0;font-family:Georgia,serif;font-size:18px;font-weight:700}
.fh{display:flex;justify-content:space-between;align-items:center;margin-bottom:18px;gap:12px}
.undobar{display:flex;gap:6px}
.undobar button{width:32px;height:32px;border-radius:8px;border:1px solid #DCD4C0;background:#FAF7EF;color:#0C1822;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;transition:all .12s}
.undobar button:hover:not(:disabled){background:#0C1822;color:#E9C877;border-color:#0C1822}
.undobar button:disabled{opacity:.32;cursor:default}
.bkp{display:flex;gap:8px}
.bkp button,.imp{display:inline-flex;align-items:center;gap:6px;font-size:12px;font-weight:600;padding:8px 13px;border-radius:9px;border:1px solid #DCD4C0;background:#FAF7EF;color:#2A3641;cursor:pointer;font-family:inherit;transition:all .12s}
.bkp button:hover,.imp:hover{background:#0C1822;color:#E9C877;border-color:#0C1822}
.fld{display:flex;flex-direction:column;gap:5px;margin-bottom:12px;font-size:11.5px;color:#5C5443}
.fld span{font-weight:600;letter-spacing:.3px}
.fld input,.fld select{border:1px solid #DCD4C0;border-radius:9px;padding:10px 12px;font-size:13.5px;font-family:inherit;background:#FCFAF4;transition:all .15s}
.fld input:focus,.fld select:focus{outline:none;border-color:#C9A24B;box-shadow:0 0 0 3px #C9A24B22;background:#fff}
.r2{display:grid;grid-template-columns:1fr 1fr;gap:11px}.r3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
.chk{display:flex;align-items:center;gap:8px;font-size:13px;margin:6px 0 12px;color:#3A4651}
.mini{width:100%;margin:2px 0 12px;padding:9px;border-radius:9px;border:1px dashed #C9A24B;background:#FBF6EA;color:#8A6515;font-weight:700;font-size:12.5px;cursor:pointer;font-family:inherit;transition:all .12s}
.mini:hover{background:#0C1822;color:#E9C877;border-color:#0C1822}
.lista{margin:8px 0 14px;border:1px solid #EDE6D4;border-radius:11px;padding:12px;background:#FAF7EF}
.lhdr{display:flex;justify-content:space-between;align-items:center;font-size:11.5px;font-weight:700;letter-spacing:.4px;text-transform:uppercase;color:#8A8064;margin-bottom:9px}
.lhdr button,.lrow button{background:#0C1822;color:#E9C877;border:none;border-radius:7px;width:25px;height:25px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center}
.lrow{display:grid;grid-template-columns:1fr 132px 30px;gap:7px;margin-bottom:7px}
.lrow input{border:1px solid #E0D8C4;border-radius:7px;padding:7px 9px;font-size:12.5px;font-family:inherit;background:#fff}
.lrow button{background:#A6483E;width:30px;height:auto;color:#fff}
.preview{background:#fff;border:1px solid #ECE5D3;border-radius:16px;padding:22px 24px;align-self:start;position:sticky;top:0;box-shadow:0 4px 22px rgba(20,30,40,.07)}
.pvhead{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;font-family:Georgia,serif;font-weight:700;font-size:15px;padding-bottom:12px;border-bottom:2px solid #0C1822}
.badge{display:inline-flex;align-items:center;gap:5px;font-size:11.5px;font-weight:700;padding:5px 11px;border-radius:30px;letter-spacing:.3px}
.badge.ok{background:#E6F3ED;color:#0C6A5A;border:1px solid #B6DECF}.badge.no{background:#FBEAE7;color:#A6483E;border:1px solid #ECC9C2}
.alerts{display:flex;flex-direction:column;gap:7px;margin-top:14px}
.alert{font-size:12px;line-height:1.4;padding:9px 12px;border-radius:9px;border-left:3px solid}
.alert.a-error{background:#FBE9E6;color:#9B3F35;border-color:#A6483E;font-weight:600}
.alert.a-warn{background:#FBF3DE;color:#7C5A12;border-color:#C8892B}
.alert.a-info{background:#EAF0F6;color:#3A6486;border-color:#3E6B8C}
.alert.a-ok{background:#E6F3ED;color:#0C6A5A;border-color:#0C6A5A;font-weight:600}
.stmt{font-size:12.5px;margin-bottom:14px}
.sec{font-family:Georgia,serif;font-weight:700;color:#0C1822;border-bottom:1px solid #C9A24B;padding:10px 0 5px;margin-top:8px;font-size:11px;letter-spacing:1.4px;text-transform:uppercase}
.ln{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px dotted #EAE3D3}
.ln span:first-child{color:#5C5443}.ln .mono{color:#26333F}
.ln.b{font-weight:700}.ln.b span:first-child,.ln.b .mono{color:#16222B}
.ln.t{background:linear-gradient(90deg,#0C1822,#163a35);color:#E9C877;font-weight:700;border-radius:8px;padding:8px 11px;margin-top:6px;border:none}
.ln.t span:first-child{color:#C9D6CF}.ln.t .mono{color:#E9C877}
.pvsub{font-size:10.5px;letter-spacing:.8px;text-transform:uppercase;color:#9B9277;margin:12px 0 4px;display:flex;align-items:center;gap:5px;font-weight:600}
.dl{margin-top:16px;width:100%;display:flex;align-items:center;justify-content:center;gap:9px;background:linear-gradient(135deg,#CDA64E,#B8923D);color:#1A1206;border:none;border-radius:11px;padding:13px;font-weight:700;font-size:14px;cursor:pointer;font-family:inherit;letter-spacing:.3px;box-shadow:0 4px 14px rgba(184,146,61,.3);transition:all .15s}
.dl:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(184,146,61,.42)}
.carta .doc{font-size:11.5px;color:#2A3640;border:1px solid #EDE7D8;border-radius:10px;padding:18px;background:#FCFAF4;box-shadow:inset 0 0 0 1px #fff,0 1px 0 #EDE7D8}
.doc .ctr{text-align:center}.doc .rgt{text-align:right}.doc .just{text-align:justify}.doc .b{font-weight:700}.doc hr{border:none;border-top:1px solid #C9A24B66;margin:9px 0}
/* ---- chat ---- */
.chatview{display:flex;flex-direction:column;height:calc(100vh - 142px);max-width:900px;margin:0 auto}
.chat{flex:1;overflow:auto;display:flex;flex-direction:column;gap:14px;padding:6px 2px 16px}
.row{display:flex;max-width:90%}.row.user{align-self:flex-end}
.bub{padding:12px 16px;border-radius:16px;font-size:14px;line-height:1.55;white-space:pre-wrap;box-shadow:0 1px 4px rgba(20,30,40,.05)}
.row.assistant .bub{background:#fff;border:1px solid #EAE3D2}
.row.user .bub{background:#0C1822;color:#F3ECDC}
.row.tool .bub{background:#FBF3DF;border:1px solid #E9D9AE;color:#7A5B16;font-family:ui-monospace,monospace;font-size:12.5px}
.who{font-size:10px;letter-spacing:1.2px;text-transform:uppercase;color:#A39A82;margin-bottom:3px;font-weight:600}
.bar{display:flex;gap:11px;padding-top:14px;border-top:1px solid #E4DDC9}
.bar textarea{flex:1;resize:none;border:1px solid #DCD4C0;border-radius:12px;padding:12px 14px;font:inherit;font-size:14px;height:48px;background:#FCFAF4}
.bar textarea:focus{outline:none;border-color:#C9A24B;box-shadow:0 0 0 3px #C9A24B22}
.bar button{border:none;background:linear-gradient(135deg,#CDA64E,#B8923D);color:#1A1206;font-weight:700;padding:0 24px;border-radius:12px;cursor:pointer;box-shadow:0 4px 14px rgba(184,146,61,.3)}
.bar button:disabled{background:#D8CFB6;color:#8A8064;box-shadow:none}
.dots{display:inline-flex;gap:4px}.dots i{width:6px;height:6px;border-radius:50%;background:#C9A24B;animation:b 1s infinite}
.dots i:nth-child(2){animation-delay:.15s}.dots i:nth-child(3){animation-delay:.3s}@keyframes b{0%,80%,100%{opacity:.3}40%{opacity:1}}
@media(max-width:860px){.app{grid-template-columns:1fr}.sb{display:none}.two{grid-template-columns:1fr}.cards,.grid2{grid-template-columns:1fr}}
`;
