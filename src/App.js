import React, { useState, useEffect, useCallback, useRef } from 'react';

// ─────────────────────────────────────────────────────────────────────────────
// 🔧 FIREBASE CONFIG — replace these values with your own project credentials
//    from: Firebase Console → Your Project → Project Settings → Your Apps
// ─────────────────────────────────────────────────────────────────────────────
const FIREBASE_DB_URL = 'https://digital-loyalty-calendar-default-rtdb.asia-southeast1.firebasedatabase.app/';
//  ↑ e.g. 'https://campaign-calendar-abc12-default-rtdb.firebaseio.com'
//  Find it in Firebase Console → Realtime Database → copy the URL shown there

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const CAMPAIGN_COLORS = [
  { label: 'Mint',       value: '#7DD4B8', text: '#0f3327' },
  { label: 'Rose',       value: '#E8A0B4', text: '#4a1a2a' },
  { label: 'Mauve',      value: '#C17BAD', text: '#ffffff' },
  { label: 'Sand',       value: '#E8D5B0', text: '#3a2a08' },
  { label: 'Sky Blue',   value: '#A8D4E8', text: '#0e2a3a' },
  { label: 'Terracotta', value: '#C89080', text: '#ffffff' },
  { label: 'Lavender',   value: '#B0A8D4', text: '#1a1a4a' },
  { label: 'Sage',       value: '#A8C4A0', text: '#0f2e0f' },
  { label: 'Peach',      value: '#F0C0A0', text: '#4a2a10' },
  { label: 'Steel',      value: '#A0B4C8', text: '#ffffff' },
  { label: 'Coral',      value: '#F4877A', text: '#ffffff' },
  { label: 'Amber',      value: '#F5C842', text: '#3a2a00' },
];

const CATEGORIES = [
  'All Categories',
  'IQOSPHERE', 'TX', 'Category Motivation', 'True Stories', 'UGC/UIC',
  'Promo', 'Lending', 'Lottery', 'Exchange', 'Hajimetewari', 'Tier Discount',
  'Earn Qoins', 'Device', 'Accessories', 'Device & Acc. Bundles', 'Terea',
  'Sentia', 'Axia', 'FSS', 'FSF', 'SIS', 'HORECA', 'LAMP', 'CVS', 'Marlboro',
  'Lark', 'CC', 'Lil Hybrid', 'eVoucher', 'Sustainability', 'Device Registration',
  'Questionnaire', 'ZYN', 'Bonds', 'Blends', 'Other', 'Consumables', 'O2O',
];

const MONTH_NAMES = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December',
];
const DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

const getDaysInMonth    = (y, m) => new Date(y, m + 1, 0).getDate();
const getFirstDayOfWeek = (y, m) => new Date(y, m, 1).getDay();

function buildCalendarDays(year, month) {
  const days = [];
  const first = getFirstDayOfWeek(year, month);
  for (let i = 0; i < first; i++) days.push(null);
  for (let d = 1; d <= getDaysInMonth(year, month); d++) days.push(d);
  while (days.length % 7 !== 0) days.push(null);
  return days;
}

function hexToObj(hex) {
  const h = hex.replace('#', '');
  return { r: parseInt(h.slice(0,2),16), g: parseInt(h.slice(2,4),16), b: parseInt(h.slice(4,6),16) };
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
    const s = document.createElement('script');
    s.src = src; s.onload = resolve; s.onerror = reject;
    document.head.appendChild(s);
  });
}

let _localId = 200;
const uid = () => String(++_localId) + '_' + Date.now();

const today = new Date();

const SEED_DATA = [
  { id:'s1', name:'TX Wave 4',    subtitle:'Prize Refresh',     notes:'Confirm assets by Mon', category:'TX',           color:'#7DD4B8', textColor:'#0f3327', date:{ year:today.getFullYear(), month:today.getMonth(), day:3  } },
  { id:'s2', name:'IQOSPHERE',    subtitle:'BAU Earn & Burn',   notes:'',                      category:'IQOSPHERE',    color:'#A8C4A0', textColor:'#0f2e0f', date:{ year:today.getFullYear(), month:today.getMonth(), day:7  } },
  { id:'s3', name:'Hajimetewari', subtitle:'Bold Ruby',         notes:'Check brief',           category:'Hajimetewari', color:'#E8D5B0', textColor:'#3a2a08', date:{ year:today.getFullYear(), month:today.getMonth(), day:10 } },
  { id:'s4', name:'ZYN',          subtitle:'FSS Multicategory', notes:'',                      category:'ZYN',          color:'#A8D4E8', textColor:'#0e2a3a', date:{ year:today.getFullYear(), month:today.getMonth(), day:14 } },
  { id:'s5', name:'Promo',        subtitle:'Grand Opening',     notes:'Assets from design',    category:'Promo',        color:'#C17BAD', textColor:'#ffffff', date:{ year:today.getFullYear(), month:today.getMonth(), day:18 } },
];

// ─────────────────────────────────────────────────────────────────────────────
// Firebase helpers — simple REST API calls, no SDK needed
// ─────────────────────────────────────────────────────────────────────────────

const DB_PATH = `${FIREBASE_DB_URL}/campaigns.json`;

async function firebaseRead() {
  const res = await fetch(DB_PATH);
  if (!res.ok) throw new Error('Firebase read failed');
  const data = await res.json();
  if (!data) return null;
  // Firebase stores arrays as objects keyed by index — convert back to array
  return Array.isArray(data) ? data : Object.values(data);
}

async function firebaseWrite(campaigns) {
  const res = await fetch(DB_PATH, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(campaigns),
  });
  if (!res.ok) throw new Error('Firebase write failed');
  return res.json();
}

// ─────────────────────────────────────────────────────────────────────────────
// Main Component
// ─────────────────────────────────────────────────────────────────────────────

export default function App() {
  const [campaigns,    setCampaigns]    = useState([]);
  const [dbStatus,     setDbStatus]     = useState('loading'); // 'loading' | 'ready' | 'error' | 'saving' | 'saved'
  const [dbError,      setDbError]      = useState('');
  const [hasUnsaved,   setHasUnsaved]   = useState(false);

  const [viewYear,     setViewYear]     = useState(today.getFullYear());
  const [viewMonth,    setViewMonth]    = useState(today.getMonth());
  const [filterCat,    setFilterCat]    = useState('All Categories');
  const [modal,        setModal]        = useState(false);
  const [editId,       setEditId]       = useState(null);
  const [form,         setForm]         = useState({
    name:'', subtitle:'', notes:'', category:CATEGORIES[1],
    color:CAMPAIGN_COLORS[0].value, textColor:CAMPAIGN_COLORS[0].text, day:'',
  });
  const [dragId,       setDragId]       = useState(null);
  const [dragOverDay,  setDragOverDay]  = useState(null);
  const [exporting,    setExporting]    = useState(null);

  // ── Load from Firebase on mount ───────────────────────────────────────────
  useEffect(() => {
    setDbStatus('loading');
    firebaseRead()
      .then(data => {
        setCampaigns(data && data.length ? data : SEED_DATA);
        setDbStatus('ready');
      })
      .catch(err => {
        console.error('Firebase load error:', err);
        setCampaigns(SEED_DATA);
        setDbStatus('error');
        setDbError('Could not connect to database. Changes will not be saved until connection is restored.');
      });
  }, []);

  // ── Save to Firebase ──────────────────────────────────────────────────────
  const saveToDatabase = useCallback(async (listToSave) => {
    const list = listToSave || campaigns;
    setDbStatus('saving');
    try {
      await firebaseWrite(list);
      setDbStatus('saved');
      setHasUnsaved(false);
      setTimeout(() => setDbStatus('ready'), 2500);
    } catch (err) {
      console.error('Save error:', err);
      setDbStatus('error');
      setDbError('Save failed. Please check your internet connection and try again.');
    }
  }, [campaigns]);

  // ── Local state helpers (do NOT auto-save — user clicks Save) ────────────
  const updateLocal = useCallback((newList) => {
    setCampaigns(newList);
    setHasUnsaved(true);
    if (dbStatus === 'saved') setDbStatus('ready');
  }, [dbStatus]);

  const calDays     = buildCalendarDays(viewYear, viewMonth);
  const daysInMonth = getDaysInMonth(viewYear, viewMonth);

  const visible = campaigns.filter(c =>
    c.date.year === viewYear &&
    c.date.month === viewMonth &&
    (filterCat === 'All Categories' || c.category === filterCat)
  );
  const forDay = (day) => visible.filter(c => c.date.day === day);

  const prevMonth = () => viewMonth === 0 ? (setViewMonth(11), setViewYear(y=>y-1)) : setViewMonth(m=>m-1);
  const nextMonth = () => viewMonth === 11 ? (setViewMonth(0),  setViewYear(y=>y+1)) : setViewMonth(m=>m+1);

  // ── Modal ─────────────────────────────────────────────────────────────────
  const openNew = (day='') => {
    setEditId(null);
    setForm({ name:'', subtitle:'', notes:'', category:CATEGORIES[1], color:CAMPAIGN_COLORS[0].value, textColor:CAMPAIGN_COLORS[0].text, day:day||'' });
    setModal(true);
  };
  const openEdit = (c, e) => {
    e.stopPropagation();
    setEditId(c.id);
    setForm({ name:c.name, subtitle:c.subtitle||'', notes:c.notes||'', category:c.category||CATEGORIES[1], color:c.color, textColor:c.textColor, day:c.date.day });
    setModal(true);
  };
  const pickColor = (v) => {
    const match = CAMPAIGN_COLORS.find(c=>c.value===v);
    setForm(f=>({...f, color:v, textColor:match?match.text:'#ffffff'}));
  };

  const handleSave = () => {
    if (!form.name.trim() || !form.day) return;
    const day = parseInt(form.day, 10);
    let next;
    if (editId) {
      next = campaigns.map(c => c.id===editId ? {...c,...form, date:{year:viewYear,month:viewMonth,day}} : c);
    } else {
      next = [...campaigns, { id:uid(), ...form, date:{year:viewYear,month:viewMonth,day} }];
    }
    updateLocal(next);
    setModal(false);
  };

  const handleDelete = () => {
    updateLocal(campaigns.filter(c=>c.id!==editId));
    setModal(false);
  };

  // ── Drag & drop ───────────────────────────────────────────────────────────
  const onDragStart = (e, id) => { setDragId(id); e.dataTransfer.effectAllowed='move'; };
  const onDragOver  = (e, day) => { if(!day) return; e.preventDefault(); setDragOverDay(day); };
  const onDrop      = (e, day) => {
    e.preventDefault();
    if (dragId && day) {
      updateLocal(campaigns.map(c => c.id===dragId ? {...c, date:{year:viewYear,month:viewMonth,day}} : c));
    }
    setDragId(null); setDragOverDay(null);
  };
  const onDragEnd = () => { setDragId(null); setDragOverDay(null); };

  // ── PPTX Export ───────────────────────────────────────────────────────────
  const exportPptx = async () => {
    setExporting('pptx');
    try {
      await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js');
      const pptx = new window.PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';
      const W=13.33,H=7.5,MAR=0.25,HDR=0.6,DLBL=0.27;
      const GRID_TOP=MAR+HDR+DLBL, COLS=7;
      const first=getFirstDayOfWeek(viewYear,viewMonth), total=getDaysInMonth(viewYear,viewMonth);
      const cells=Math.ceil((first+total)/7)*7, ROWS=cells/7;
      const CW=(W-MAR*2)/COLS, CH=(H-GRID_TOP-MAR)/ROWS;
      const exportList = filterCat==='All Categories'
        ? campaigns.filter(c=>c.date.year===viewYear&&c.date.month===viewMonth)
        : campaigns.filter(c=>c.date.year===viewYear&&c.date.month===viewMonth&&c.category===filterCat);
      const slide=pptx.addSlide();
      slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:'F7F8FA'},line:{color:'F7F8FA'}});
      slide.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:HDR,fill:{color:'1a1d2e'},line:{color:'1a1d2e'}});
      slide.addText(`${MONTH_NAMES[viewMonth]} ${viewYear}  ·  Campaign Calendar`,{x:MAR,y:0.08,w:9,h:HDR-0.16,fontSize:20,bold:true,color:'FFFFFF',fontFace:'Calibri',valign:'middle'});
      if(filterCat!=='All Categories') slide.addText(`Category: ${filterCat}`,{x:W-3.5,y:0.1,w:3.2,h:HDR-0.2,fontSize:11,color:'A8D4E8',fontFace:'Calibri',align:'right',valign:'middle'});
      DAY_NAMES.forEach((d,i)=>{
        const x=MAR+i*CW;
        slide.addShape(pptx.ShapeType.rect,{x,y:MAR+HDR,w:CW,h:DLBL,fill:{color:'2d3250'},line:{color:'2d3250'}});
        slide.addText(d,{x,y:MAR+HDR,w:CW,h:DLBL,fontSize:8.5,bold:true,color:'FFFFFF',fontFace:'Calibri',align:'center',valign:'middle'});
      });
      for(let i=0;i<cells;i++){
        const col=i%COLS,row=Math.floor(i/COLS),day=i-first+1,valid=day>=1&&day<=total;
        const cx=MAR+col*CW,cy=GRID_TOP+row*CH;
        slide.addShape(pptx.ShapeType.rect,{x:cx,y:cy,w:CW,h:CH,fill:{color:valid?'FFFFFF':'F3F4F8'},line:{color:'E0E3EC',pt:0.5}});
        if(!valid) continue;
        const isTdy=viewYear===today.getFullYear()&&viewMonth===today.getMonth()&&day===today.getDate();
        if(isTdy){
          slide.addShape(pptx.ShapeType.ellipse,{x:cx+0.05,y:cy+0.04,w:0.2,h:0.2,fill:{color:'3B5BDB'},line:{color:'3B5BDB'}});
          slide.addText(String(day),{x:cx+0.05,y:cy+0.04,w:0.2,h:0.2,fontSize:8,bold:true,color:'FFFFFF',fontFace:'Calibri',align:'center',valign:'middle'});
        } else {
          slide.addText(String(day),{x:cx+0.06,y:cy+0.04,w:0.2,h:0.18,fontSize:8,bold:true,color:'4a5068',fontFace:'Calibri'});
        }
        const dayCamps=exportList.filter(c=>c.date.day===day);
        const BOX_H=0.3,GAP=0.03;
        dayCamps.forEach((c,ci)=>{
          const by=cy+0.26+ci*(BOX_H+GAP);
          if(by+BOX_H>cy+CH-0.03) return;
          const bg=c.color.replace('#',''),fg=c.textColor.replace('#','');
          slide.addShape(pptx.ShapeType.roundRect,{x:cx+0.05,y:by,w:CW-0.1,h:BOX_H,fill:{color:bg},line:{color:bg},rectRadius:0.04});
          slide.addText(c.name,{x:cx+0.08,y:by+0.01,w:CW-0.16,h:0.13,fontSize:7.5,bold:true,color:fg,fontFace:'Calibri',valign:'middle'});
          if(c.subtitle) slide.addText(c.subtitle,{x:cx+0.08,y:by+0.13,w:CW-0.16,h:0.1,fontSize:6,color:fg,fontFace:'Calibri'});
          if(c.notes)    slide.addText(`📝 ${c.notes}`,{x:cx+0.08,y:by+0.21,w:CW-0.16,h:0.09,fontSize:5.5,color:fg,fontFace:'Calibri',italic:true});
        });
      }
      slide.addText(`Exported ${new Date().toLocaleDateString()}  ·  Campaign Calendar`,{x:MAR,y:H-0.22,w:W-MAR*2,h:0.18,fontSize:7,color:'8a90a0',fontFace:'Calibri',align:'right'});
      await pptx.writeFile({fileName:`Campaign_Calendar_${MONTH_NAMES[viewMonth]}_${viewYear}`});
    } catch(err){ console.error('PPTX error:',err); alert('PPTX export failed.'); }
    setExporting(null);
  };

  // ── PDF Export ────────────────────────────────────────────────────────────
  const exportPdf = async () => {
    setExporting('pdf');
    try {
      await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({orientation:'landscape',unit:'pt',format:'a4'});
      const PW=841.89,PH=595.28,MAR=18,HDR=38,DLBL=18;
      const GRID_TOP=MAR+HDR+DLBL, COLS=7;
      const first=getFirstDayOfWeek(viewYear,viewMonth), total=getDaysInMonth(viewYear,viewMonth);
      const cells=Math.ceil((first+total)/7)*7, ROWS=cells/7;
      const CW=(PW-MAR*2)/COLS, CH=(PH-GRID_TOP-MAR)/ROWS;
      const exportList = filterCat==='All Categories'
        ? campaigns.filter(c=>c.date.year===viewYear&&c.date.month===viewMonth)
        : campaigns.filter(c=>c.date.year===viewYear&&c.date.month===viewMonth&&c.category===filterCat);
      doc.setFillColor(26,29,46); doc.rect(0,0,PW,HDR,'F');
      doc.setTextColor(255,255,255); doc.setFontSize(16); doc.setFont('helvetica','bold');
      doc.text(`${MONTH_NAMES[viewMonth]} ${viewYear}  ·  Campaign Calendar`,MAR,HDR*0.68);
      if(filterCat!=='All Categories'){
        doc.setFontSize(9); doc.setFont('helvetica','normal'); doc.setTextColor(168,212,232);
        doc.text(`Category: ${filterCat}`,PW-MAR,HDR*0.68,{align:'right'});
      }
      doc.setFillColor(45,50,80); doc.rect(MAR,MAR+HDR,PW-MAR*2,DLBL,'F');
      DAY_NAMES.forEach((d,i)=>{
        doc.setFontSize(8); doc.setFont('helvetica','bold'); doc.setTextColor(255,255,255);
        doc.text(d,MAR+i*CW+CW/2,MAR+HDR+DLBL*0.72,{align:'center'});
      });
      for(let i=0;i<cells;i++){
        const col=i%COLS,row=Math.floor(i/COLS),day=i-first+1,valid=day>=1&&day<=total;
        const cx=MAR+col*CW,cy=GRID_TOP+row*CH;
        doc.setFillColor(...(valid?[255,255,255]:[243,244,248]));
        doc.setDrawColor(224,227,236); doc.setLineWidth(0.4);
        doc.rect(cx,cy,CW,CH,'FD');
        if(!valid) continue;
        const isTdy=viewYear===today.getFullYear()&&viewMonth===today.getMonth()&&day===today.getDate();
        if(isTdy){
          doc.setFillColor(59,91,219); doc.circle(cx+9,cy+9,7,'F');
          doc.setFontSize(7); doc.setFont('helvetica','bold'); doc.setTextColor(255,255,255);
          doc.text(String(day),cx+9,cy+12,{align:'center'});
        } else {
          doc.setFontSize(7.5); doc.setFont('helvetica','bold'); doc.setTextColor(74,80,104);
          doc.text(String(day),cx+5,cy+12);
        }
        const dayCamps=exportList.filter(c=>c.date.day===day);
        const BH=18,GAP=2;
        dayCamps.forEach((c,ci)=>{
          const by=cy+20+ci*(BH+GAP);
          if(by+BH>cy+CH-2) return;
          const {r,g,b}=hexToObj(c.color);
          doc.setFillColor(r,g,b); doc.setDrawColor(r,g,b);
          doc.roundedRect(cx+3,by,CW-6,BH,2,2,'FD');
          const {r:tr,g:tg,b:tb}=hexToObj(c.textColor);
          doc.setTextColor(tr,tg,tb);
          doc.setFontSize(7); doc.setFont('helvetica','bold');
          doc.text(c.name.slice(0,18),cx+6,by+8);
          if(c.subtitle){doc.setFontSize(5.5);doc.setFont('helvetica','normal');doc.text(c.subtitle.slice(0,22),cx+6,by+14);}
          if(c.notes){doc.setFontSize(5);doc.setFont('helvetica','italic');doc.text(c.notes.slice(0,24),cx+6,by+BH-2);}
        });
      }
      doc.setFontSize(6.5); doc.setFont('helvetica','normal'); doc.setTextColor(138,144,160);
      doc.text(`Exported ${new Date().toLocaleDateString()}  ·  Campaign Calendar`,PW-MAR,PH-6,{align:'right'});
      doc.save(`Campaign_Calendar_${MONTH_NAMES[viewMonth]}_${viewYear}.pdf`);
    } catch(err){ console.error('PDF error:',err); alert('PDF export failed.'); }
    setExporting(null);
  };

  // ─────────────────────────────────────────────────────────────────────────
  // Status bar config
  // ─────────────────────────────────────────────────────────────────────────
  const statusConfig = {
    loading: { bg:'#f0f4ff', border:'#c7d4f8', color:'#3B5BDB', text:'⏳ Loading calendar data…', showSave:false },
    ready:   { bg: hasUnsaved ? '#fff8e1':'#f0fdf4', border: hasUnsaved ? '#fde68a':'#bbf7d0', color: hasUnsaved ? '#92400e':'#166534', text: hasUnsaved ? '● Unsaved changes — click Save to publish for all users':'✓ All changes saved', showSave: hasUnsaved },
    saving:  { bg:'#f0f4ff', border:'#c7d4f8', color:'#3B5BDB', text:'💾 Saving to database…', showSave:false },
    saved:   { bg:'#f0fdf4', border:'#bbf7d0', color:'#166534', text:'✅ Saved! All users will see these changes.', showSave:false },
    error:   { bg:'#fef2f2', border:'#fecaca', color:'#991b1b', text:`⚠ ${dbError}`, showSave:true },
  };
  const status = statusConfig[dbStatus] || statusConfig.ready;

  // ─────────────────────────────────────────────────────────────────────────
  // Render
  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div style={{fontFamily:"'Inter','Segoe UI',sans-serif",minHeight:'100vh',background:'#F7F8FA',padding:20}}>
      <style>{`
        @keyframes spin{to{transform:rotate(360deg)}}
        .camp-card{transition:transform 0.1s,filter 0.1s}
        .camp-card:hover{filter:brightness(0.91);transform:translateY(-1px)}
        .cal-cell:hover{background:#F3F5FF!important}
      `}</style>

      <div style={{maxWidth:1300,margin:'0 auto'}}>

        {/* ── Status / Save bar ── */}
        <div style={{
          display:'flex', alignItems:'center', justifyContent:'space-between',
          background:status.bg, border:`1px solid ${status.border}`,
          borderRadius:10, padding:'10px 16px', marginBottom:14,
          color:status.color, fontSize:13, fontWeight:600, gap:12, flexWrap:'wrap',
        }}>
          <span>{status.text}</span>
          {status.showSave && (
            <button
              onClick={() => saveToDatabase(campaigns)}
              disabled={dbStatus === 'saving'}
              style={{
                background:'#1a1d2e', color:'#fff', border:'none',
                borderRadius:7, padding:'8px 22px', fontSize:13,
                fontWeight:700, cursor:'pointer', letterSpacing:'0.2px',
                boxShadow:'0 2px 8px rgba(0,0,0,0.18)',
              }}
            >
              💾 Save Changes
            </button>
          )}
        </div>

        {/* ── Top bar ── */}
        <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:14,flexWrap:'wrap',gap:10}}>
          <div>
            <h1 style={{margin:0,fontSize:21,fontWeight:800,color:'#1a1d2e',letterSpacing:'-0.4px'}}>Campaign Calendar</h1>
            <p style={{margin:'3px 0 0',fontSize:12,color:'#8a90a0'}}>Add or drag campaigns, then click <strong>Save Changes</strong> to publish for all users</p>
          </div>
          <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center'}}>
            <select value={filterCat} onChange={e=>setFilterCat(e.target.value)} style={selectSt}>
              {CATEGORIES.map(c=><option key={c}>{c}</option>)}
            </select>
            <button onClick={exportPdf} disabled={!!exporting} style={{...btnSt,background:'#c0392b',color:'#fff',opacity:exporting?0.55:1}}>
              {exporting==='pdf'?'Exporting…':'⬇ PDF'}
            </button>
            <button onClick={exportPptx} disabled={!!exporting} style={{...btnSt,background:'#b7550a',color:'#fff',opacity:exporting?0.55:1}}>
              {exporting==='pptx'?'Exporting…':'⬇ PPTX'}
            </button>
            <button onClick={()=>openNew()} style={{...btnSt,background:'#3B5BDB',color:'#fff'}}>
              + Add Campaign
            </button>
          </div>
        </div>

        {/* ── Month nav ── */}
        <div style={{display:'flex',alignItems:'center',gap:14,marginBottom:12}}>
          <button onClick={prevMonth} style={navBtnSt}>‹</button>
          <h2 style={{margin:0,fontSize:17,fontWeight:700,color:'#1a1d2e',minWidth:195,textAlign:'center'}}>
            {MONTH_NAMES[viewMonth]} {viewYear}
          </h2>
          <button onClick={nextMonth} style={navBtnSt}>›</button>
          {filterCat!=='All Categories'&&(
            <span style={{fontSize:12,background:'#EEF2FF',color:'#3B5BDB',borderRadius:20,padding:'3px 11px',fontWeight:600,display:'flex',alignItems:'center',gap:6}}>
              {filterCat}
              <span style={{cursor:'pointer',fontWeight:800}} onClick={()=>setFilterCat('All Categories')}>✕</span>
            </span>
          )}
        </div>

        {/* ── Calendar ── */}
        <div style={{background:'#fff',borderRadius:14,boxShadow:'0 2px 16px rgba(0,0,0,0.07)',overflow:'hidden'}}>
          <div style={{display:'grid',gridTemplateColumns:'repeat(7,1fr)',background:'#2d3250'}}>
            {DAY_NAMES.map(d=>(
              <div key={d} style={{padding:'9px 0',textAlign:'center',fontSize:11.5,fontWeight:700,color:'#ffffff',letterSpacing:'0.6px',textTransform:'uppercase'}}>{d}</div>
            ))}
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(7,1fr)'}}>
            {calDays.map((day,i)=>{
              const isToday=day&&viewYear===today.getFullYear()&&viewMonth===today.getMonth()&&day===today.getDate();
              const isOver=dragOverDay===day&&day!==null;
              const dayCamps=day?forDay(day):[];
              return (
                <div key={i} className={day?'cal-cell':''}
                  onClick={()=>day&&openNew(day)}
                  onDragOver={e=>onDragOver(e,day)}
                  onDrop={e=>onDrop(e,day)}
                  style={{
                    minHeight:118,
                    borderRight:(i+1)%7===0?'none':'1px solid #eef0f5',
                    borderBottom:i<calDays.length-7?'1px solid #eef0f5':'none',
                    padding:'6px 5px 5px',
                    cursor:day?'pointer':'default',
                    background:isOver?'#EEF2FF':day?'#fff':'#FAFBFC',
                    transition:'background 0.12s',
                  }}>
                  {day&&(
                    <>
                      <div style={{fontSize:11.5,fontWeight:700,color:isToday?'#fff':'#4a5068',background:isToday?'#3B5BDB':'transparent',borderRadius:'50%',width:22,height:22,display:'flex',alignItems:'center',justifyContent:'center',marginBottom:4}}>
                        {day}
                      </div>
                      <div style={{display:'flex',flexDirection:'column',gap:3}}>
                        {dayCamps.map(c=>(
                          <div key={c.id} className="camp-card" draggable
                            onDragStart={e=>{e.stopPropagation();onDragStart(e,c.id);}}
                            onDragEnd={onDragEnd}
                            onClick={e=>openEdit(c,e)}
                            style={{background:c.color,color:c.textColor,borderRadius:5,padding:'4px 6px',fontSize:10.5,fontWeight:700,cursor:'grab',opacity:dragId===c.id?0.3:1,boxShadow:'0 1px 3px rgba(0,0,0,0.13)',userSelect:'none',lineHeight:1.35}}>
                            <div>{c.name}</div>
                            {c.subtitle&&<div style={{fontWeight:400,fontSize:9.5,opacity:0.85}}>{c.subtitle}</div>}
                            {c.notes&&<div style={{fontWeight:400,fontSize:8.5,opacity:0.7,marginTop:1,fontStyle:'italic'}}>📝 {c.notes}</div>}
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* ── Legend ── */}
        <div style={{marginTop:12,display:'flex',flexWrap:'wrap',gap:8}}>
          {CAMPAIGN_COLORS.map(c=>(
            <div key={c.value} style={{display:'flex',alignItems:'center',gap:5,fontSize:11,color:'#6a7080'}}>
              <div style={{width:11,height:11,borderRadius:3,background:c.value,flexShrink:0}}/>
              {c.label}
            </div>
          ))}
        </div>
      </div>

      {/* ── Modal ── */}
      {modal&&(
        <div style={{position:'fixed',inset:0,background:'rgba(10,15,40,0.48)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:200}}
          onClick={()=>setModal(false)}>
          <div style={{background:'#fff',borderRadius:14,padding:26,width:440,maxWidth:'93vw',boxShadow:'0 24px 70px rgba(0,0,0,0.22)',maxHeight:'92vh',overflowY:'auto'}}
            onClick={e=>e.stopPropagation()}>
            <h3 style={{margin:'0 0 16px',fontSize:16,fontWeight:700,color:'#1a1d2e'}}>{editId?'Edit Campaign':'New Campaign'}</h3>

            <Lbl>Campaign Name *</Lbl>
            <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="e.g. TX Wave 4" style={inpSt}/>

            <Lbl>Subtitle</Lbl>
            <input value={form.subtitle} onChange={e=>setForm(f=>({...f,subtitle:e.target.value}))} placeholder="e.g. Prize Refresh" style={inpSt}/>

            <Lbl>Notes <span style={{fontWeight:400,color:'#aaa',fontSize:11}}>(small italic on card)</span></Lbl>
            <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))}
              placeholder="e.g. Confirm assets by Monday" rows={2}
              style={{...inpSt,resize:'vertical',minHeight:52}}/>

            <Lbl>Category</Lbl>
            <select value={form.category} onChange={e=>setForm(f=>({...f,category:e.target.value}))} style={{...inpSt,cursor:'pointer'}}>
              {CATEGORIES.filter(c=>c!=='All Categories').map(c=><option key={c}>{c}</option>)}
            </select>

            <Lbl>Day of Month *</Lbl>
            <input type="number" min={1} max={daysInMonth} value={form.day}
              onChange={e=>setForm(f=>({...f,day:e.target.value}))}
              placeholder={`1–${daysInMonth}`} style={{...inpSt,width:80}}/>

            <Lbl>Color</Lbl>
            <div style={{display:'flex',flexWrap:'wrap',gap:7,marginBottom:14}}>
              {CAMPAIGN_COLORS.map(c=>(
                <button key={c.value} onClick={()=>pickColor(c.value)} title={c.label}
                  style={{width:30,height:30,borderRadius:7,background:c.value,border:'none',cursor:'pointer',outline:form.color===c.value?'3px solid #3B5BDB':'2px solid transparent',outlineOffset:2}}/>
              ))}
            </div>

            <Lbl>Preview</Lbl>
            <div style={{marginBottom:16}}>
              <div style={{display:'inline-block',background:form.color,color:form.textColor,borderRadius:6,padding:'6px 11px',fontSize:11.5,fontWeight:700,lineHeight:1.4}}>
                {form.name||'Campaign Name'}
                {form.subtitle&&<div style={{fontWeight:400,fontSize:10.5}}>{form.subtitle}</div>}
                {form.notes&&<div style={{fontWeight:400,fontSize:9,opacity:0.75,marginTop:1,fontStyle:'italic'}}>📝 {form.notes}</div>}
              </div>
            </div>

            <div style={{display:'flex',gap:8,justifyContent:'flex-end'}}>
              {editId&&<button onClick={handleDelete} style={{...btnSt,background:'#fee2e2',color:'#c0392b'}}>Delete</button>}
              <button onClick={()=>setModal(false)} style={{...btnSt,background:'#f0f1f5',color:'#4a5068'}}>Cancel</button>
              <button onClick={handleSave} style={{...btnSt,background:'#3B5BDB',color:'#fff'}}>{editId?'Save Changes':'Add to Calendar'}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function Lbl({children}){
  return <label style={{display:'block',fontSize:12,fontWeight:600,color:'#4a5068',marginBottom:5}}>{children}</label>;
}

const navBtnSt={background:'#fff',border:'1px solid #e0e3ec',borderRadius:8,width:34,height:34,fontSize:17,cursor:'pointer',color:'#4a5068',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:700};
const btnSt={border:'none',borderRadius:7,padding:'8px 15px',fontSize:13,fontWeight:600,cursor:'pointer'};
const selectSt={border:'1px solid #e0e3ec',borderRadius:7,padding:'7px 10px',fontSize:13,color:'#1a1d2e',cursor:'pointer',background:'#fff',fontFamily:'inherit',maxWidth:190};
const inpSt={display:'block',width:'100%',boxSizing:'border-box',border:'1px solid #e0e3ec',borderRadius:7,padding:'8px 11px',fontSize:13.5,color:'#1a1d2e',marginBottom:13,outline:'none',fontFamily:'inherit'};