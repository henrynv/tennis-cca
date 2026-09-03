/* =====================================================================
   BANC D'ESSAI DE L'EXPORT « GROUPES ET EDT »

   Le classeur que sort le bouton « ⬇︎ Excel » ne se relit pas a l'oeil dans
   le navigateur : il faut cliquer, ouvrir Excel, et recommencer a chaque
   retouche. Ce banc fait la meme chose en une commande, sur l'etat reel du
   club, et sait en plus rendre l'onglet « edt » en HTML pour juger la mise en
   page sans ouvrir Excel.

   Le code teste n'est pas duplique : le script lit `../index.html`, en extrait
   le dernier bloc <script> — celui de l'ecran Groupes et EDT — le deballe de
   son IIFE et l'execute sous node avec un DOM en carton. Les fonctions
   `feuilleEdt`, `feuilleProfs`, `feuilleTerrains`, `feuilleJoueurs` et
   `feuillesDetail` sont alors appelables telles quelles. Si l'une d'elles est
   renommee, le banc s'arrete en le disant : c'est voulu.

       cd site/banc && npm install
       node banc-edt.js                     # ecrit sortie/groupes-cca.xlsx
       node banc-edt.js --apercu            # + sortie/apercu-edt.html
       node banc-edt.js --sans-contraintes  # le chemin sans contraintes-tennis.xlsx
       node banc-edt.js --etat=/chemin/vers/un/autre/etat-ecran.json

   ===================================================================== */
"use strict";
const fs = require("fs"), vm = require("vm"), path = require("path");

const ICI = __dirname;
const arg = (n, d) => {
  const a = process.argv.find(x => x.startsWith("--" + n + "="));
  return a ? a.split("=").slice(1).join("=") : d;
};
const drapeau = n => process.argv.indexOf("--" + n) >= 0;

const F_HTML   = path.resolve(ICI, "..", "index.html");
const F_ETAT   = path.resolve(ICI, arg("etat", path.join("..", "..", "solveur", "etat-ecran.json")));
const D_SORTIE = path.resolve(ICI, arg("sortie", "sortie"));

let ExcelJS;
try { ExcelJS = require("exceljs"); }
catch(e){ arret("exceljs n'est pas installé. Lancez `npm install` dans site/banc."); }

function arret(msg){ console.error("⛔ " + msg); process.exit(1); }

/* ---------------- le DOM en carton ----------------
   Tout ce que le bloc « Groupes et EDT » touche au chargement : il cable des
   boutons, interroge l'agent, lit le stockage. Rien de tout ca n'existe ici,
   et rien de tout ca ne doit faire echouer le chargement — on repond « oui »
   a tout, sans rien faire. */
const rien = new Proxy(function(){}, {
  get: () => rien, set: () => true, apply: () => rien, construct: () => rien,
});
Object.assign(global, {
  window: global, addEventListener(){}, removeEventListener(){}, indexedDB: rien,
  matchMedia: () => ({matches:false, addEventListener(){}}),
  requestAnimationFrame: () => 0, getComputedStyle: () => rien, setInterval: () => 0,
  document: {
    createElement: () => rien, getElementById: () => rien, querySelectorAll: () => [],
    querySelector: () => null, addEventListener(){}, head: rien, body: rien,
    documentElement: rien,
  },
  localStorage: {getItem: () => null, setItem(){}, removeItem(){}},
  location: {port:"", href:"", search:"", hash:""},
  fetch: () => Promise.reject(new Error("le banc d'essai est hors ligne")),
  XLSX: rien, supabase: rien, pdfjsLib: rien, ExcelJS,
  alert: m => console.log("  (alerte de l'écran :", m + ")"),
  // ce que ce bloc attend des scripts precedents de la page
  DAYS: ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi"],
  PROFS: ["Loïc","Seb","Titi","Max","Valentin","Lucas"],
  ALIAS_PROFS: {}, F: () => rien, dl(){}, showScreen(){}, sb: rien,
});
process.on("unhandledRejection", () => {});   // pingAgent() et consorts, sans reseau

/* ---------------- charger l'ecran ---------------- */
if(!fs.existsSync(F_HTML)) arret("introuvable : " + F_HTML);
const html = fs.readFileSync(F_HTML, "utf8");
const blocs = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const bloc = blocs[blocs.length - 1];
if(!bloc || !/GROUPES ET EMPLOI DU TEMPS/.test(bloc))
  arret("le dernier bloc <script> de index.html n'est plus celui de l'écran Groupes et EDT.");
// le bloc est une IIFE : on la deballe pour atteindre ses fonctions d'ici
const source = bloc.replace("(function(){", "").replace(/\}\)\(\);\s*$/, "");
try { vm.runInThisContext(source, {filename:"groupes-edt.js"}); }
catch(e){ arret("le bloc « Groupes et EDT » n'a pas pu s'exécuter : " + e.message
                + "\n   " + String(e.stack).split("\n")[1]); }
["feuilleEdt","feuilleProfs","feuilleTerrains","feuilleJoueurs","feuillesDetail"]
  .forEach(n => { if(typeof global[n] !== "function" && vm.runInThisContext("typeof "+n) !== "function")
    arret("fonction « " + n + " » introuvable — a-t-elle été renommée dans index.html ?"); });

/* ---------------- l'etat du club ---------------- */
if(!fs.existsSync(F_ETAT))
  arret("état introuvable : " + F_ETAT + "\n   Lancez l'agent au moins une fois, ou passez --etat=…");
const brut = JSON.parse(fs.readFileSync(F_ETAT, "utf8"));
global.__etat = brut.etat || brut;          // l'agent écrit {etat:…}, le navigateur l'état nu
if(!Array.isArray(global.__etat.cours) || !global.__etat.cours.length)
  arret("cet état ne contient aucun cours : il n'y a rien à exporter.");
vm.runInThisContext(`
  cours   = __etat.cours;
  attente = __etat.attente || [];
  annexes = __etat.annexes || {enPlace:[], hors:[]};
  ctr     = ${drapeau("sans-contraintes") ? "null" : "__etat.ctr || null"};
  resFic  = __etat.resFic || null;
  scen    = __etat.scen || [];
  _teintes = null;
`);

/* ---------------- l'apercu HTML ----------------
   Une table qui rejoue les fonds, les cadres, les fusions, les largeurs et les
   hauteurs de la feuille. Ce n'est pas Excel, mais ca suffit a voir si un pave
   tombe juste et si le planning se lit. */
function apercu(ws, titre){
  const merges = {}, masques = new Set();
  (ws.model.merges || []).forEach(r => {
    const m = String(r).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/); if(!m) return;
    const col = s => [...s].reduce((n,ch)=>n*26 + ch.charCodeAt(0)-64, 0);
    const c1 = col(m[1]), r1 = +m[2], c2 = col(m[3]), r2 = +m[4];
    merges[r1+":"+c1] = [r2-r1+1, c2-c1+1];
    for(let r=r1; r<=r2; r++) for(let c=c1; c<=c2; c++)
      if(r!==r1 || c!==c1) masques.add(r+":"+c);
  });
  const coul = o => (o && typeof o.argb === "string" && o.argb.length === 8) ? "#" + o.argb.slice(2) : null;
  const ech = s => String(s).replace(/[&<>]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));
  const out = [`<!doctype html><meta charset="utf-8"><title>${ech(titre)}</title>`,
    "<style>body{background:#fff;font-family:Calibri,Arial,sans-serif;margin:12px}",
    "table{border-collapse:collapse}td{font-size:9px;padding:0 2px;overflow:hidden;white-space:nowrap}",
    "</style><table>"];
  for(let r=1; r<=ws.rowCount; r++){
    const h = ws.getRow(r).height || 15;
    out.push(`<tr style="height:${Math.round(h*1.05)}px">`);
    for(let c=1; c<=ws.columnCount; c++){
      if(masques.has(r+":"+c)) continue;
      const cell = ws.getCell(r,c), [rs,cs] = merges[r+":"+c] || [1,1], st = [];
      const f = cell.fill && cell.fill.type === "pattern" ? coul(cell.fill.fgColor) : null;
      if(f) st.push("background:" + f);
      const fo = cell.font || {};
      if(fo.bold)   st.push("font-weight:700");
      if(fo.italic) st.push("font-style:italic");
      if(fo.size)   st.push(`font-size:${Math.round(fo.size*1.15)}px`);
      if(coul(fo.color)) st.push("color:" + coul(fo.color));
      const b = cell.border || {};
      [["top","border-top"],["bottom","border-bottom"],["left","border-left"],["right","border-right"]]
        .forEach(([k,css]) => { const s = b[k];
          if(s && s.style) st.push(`${css}:${s.style==="medium"?2:1}px solid ${coul(s.color)||"#999"}`); });
      if(cell.alignment && cell.alignment.horizontal) st.push("text-align:" + cell.alignment.horizontal);
      let l = 0;
      for(let x=c; x<c+cs; x++) l += (ws.getColumn(x).width || 8);
      st.push(`width:${Math.round(l*7)}px;max-width:${Math.round(l*7)}px`);
      const v = cell.value;
      out.push(`<td rowspan="${rs}" colspan="${cs}" style="${st.join(";")}">${v==null?"":ech(v)}</td>`);
    }
    out.push("</tr>");
  }
  return out.join("") + "</table>";
}

/* ---------------- on y va ---------------- */
(async () => {
  fs.mkdirSync(D_SORTIE, {recursive:true});
  const wb = new ExcelJS.Workbook();
  const appel = n => vm.runInThisContext(n + "(__wb)");
  global.__wb = wb;
  ["feuilleEdt","feuilleProfs","feuilleTerrains","feuilleJoueurs","feuillesDetail"].forEach(appel);

  const f = path.join(D_SORTIE, "groupes-cca.xlsx");
  await wb.xlsx.writeFile(f);
  console.log("état  : " + F_ETAT + (drapeau("sans-contraintes") ? "  (contraintes ignorées)" : ""));
  console.log("écrit : " + f);
  wb.worksheets.forEach(w => console.log(`        ${w.name} — ${w.rowCount} lignes × ${w.columnCount} colonnes`));

  if(drapeau("apercu")){
    const g = path.join(D_SORTIE, "apercu-edt.html");
    fs.writeFileSync(g, apercu(wb.getWorksheet("edt"), "Aperçu de l'onglet edt"));
    console.log("aperçu: " + g);
  }
})().catch(e => arret(String(e && e.stack || e)));
