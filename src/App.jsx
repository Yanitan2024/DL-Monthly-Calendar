import React, { useState, useEffect, useCallback, useRef } from 'react';
import PptxGenJS from 'pptxgenjs';

// ─────────────────────────────────────────────────────────────────────────────
// 🔧 FIREBASE CONFIG — paste your Realtime Database URL here
// ─────────────────────────────────────────────────────────────────────────────
const FIREBASE_DB_URL = 'https://YOUR-PROJECT-ID-default-rtdb.firebaseio.com';

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const CAMPAIGN_COLORS = [
  { label: 'Mint',        value: '#7DD4B8', text: '#0f3327' },
  { label: 'Turquoise',   value: '#2EC4C4', text: '#ffffff' },
  { label: 'Sky Blue',    value: '#A8D4E8', text: '#0e2a3a' },
  { label: 'Ocean Blue',  value: '#3A86D4', text: '#ffffff' },
  { label: 'Navy',        value: '#2D4A8A', text: '#ffffff' },
  { label: 'Lavender',    value: '#B0A8D4', text: '#1a1a4a' },
  { label: 'Mauve',       value: '#C17BAD', text: '#ffffff' },
  { label: 'Rose',        value: '#E8A0B4', text: '#4a1a2a' },
  { label: 'Crimson',     value: '#D64045', text: '#ffffff' },
  { label: 'Coral',       value: '#F4877A', text: '#ffffff' },
  { label: 'Terracotta',  value: '#C89080', text: '#ffffff' },
  { label: 'Peach',       value: '#F0C0A0', text: '#4a2a10' },
  { label: 'Amber',       value: '#F5C842', text: '#3a2a00' },
  { label: 'Sand',        value: '#E8D5B0', text: '#3a2a08' },
  { label: 'Sage',        value: '#A8C4A0', text: '#0f2e0f' },
  { label: 'Forest',      value: '#4A8C6A', text: '#ffffff' },
  { label: 'Slate',       value: '#7A8FA6', text: '#ffffff' },
  { label: 'Steel',       value: '#A0B4C8', text: '#1a2a3a' },
  { label: 'Silver',      value: '#C8CDD6', text: '#2a2a2a' },
  { label: 'Charcoal',    value: '#5A6272', text: '#ffffff' },
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
const DAY_NAMES_ALL = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const DAY_NAMES_WD  = ['Mon','Tue','Wed','Thu','Fri'];
// day-of-week indices that are weekdays
const WEEKDAY_COLS  = [1,2,3,4,5]; // Mon=1 ... Fri=5

const getDaysInMonth    = (y, m) => new Date(y, m + 1, 0).getDate();
const getFirstDayOfWeek = (y, m) => new Date(y, m, 1).getDay();

// Returns array of {day, dow} for the grid.
// showWeekends=true → 7-col grid; false → 5-col Mon-Fri grid (weekend days skipped)
function buildCalendarDays(year, month, showWeekends) {
  const daysInMo = getDaysInMonth(year, month);
  const all = [];
  for (let d = 1; d <= daysInMo; d++) {
    const dow = new Date(year, month, d).getDay(); // 0=Sun
    if (!showWeekends && (dow === 0 || dow === 6)) continue;
    all.push({ day: d, dow });
  }
  if (showWeekends) {
    // Pad start so day 1 lands on its correct column (Sun=0)
    const firstDow = new Date(year, month, 1).getDay();
    const padded = [];
    for (let i = 0; i < firstDow; i++) padded.push(null);
    all.forEach(x => padded.push(x));
    while (padded.length % 7 !== 0) padded.push(null);
    return padded;
  } else {
    // Pad start so day 1 lands on its Mon-Fri column
    const firstDow = new Date(year, month, 1).getDay();
    // firstDow: 0=Sun(skip), 1=Mon(col0), ... 5=Fri(col4), 6=Sat(skip)
    // Count how many Mon-Fri days precede first weekday
    let pad = 0;
    if (firstDow >= 1 && firstDow <= 5) pad = firstDow - 1;
    const padded = [];
    for (let i = 0; i < pad; i++) padded.push(null);
    all.forEach(x => padded.push(x));
    while (padded.length % 5 !== 0) padded.push(null);
    return padded;
  }
}

function hexToObj(hex) {
  const h = hex.replace('#', '');
  return { r: parseInt(h.slice(0,2),16), g: parseInt(h.slice(2,4),16), b: parseInt(h.slice(4,6),16) };
}

// Robustly load a script and wait for a named global to appear on window
function loadScriptWithGlobal(src, globalName, timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    // If already available, resolve immediately
    if (window[globalName]) { resolve(window[globalName]); return; }

    // Remove any previously failed script tag for this src so it re-fetches
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) existing.remove();

    const s = document.createElement('script');
    s.src = src;
    s.async = true;

    const timer = setTimeout(() => {
      reject(new Error(`Timeout waiting for ${globalName} from ${src}`));
    }, timeoutMs);

    s.onload = () => {
      clearTimeout(timer);
      // Poll briefly — some bundles set the global slightly after onload
      let attempts = 0;
      const poll = setInterval(() => {
        if (window[globalName]) {
          clearInterval(poll);
          resolve(window[globalName]);
        } else if (++attempts > 20) {
          clearInterval(poll);
          reject(new Error(`${globalName} not found on window after script loaded`));
        }
      }, 50);
    };
    s.onerror = () => { clearTimeout(timer); reject(new Error(`Failed to load script: ${src}`)); };
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
// Firebase REST helpers
// ─────────────────────────────────────────────────────────────────────────────

const DB_PATH = `${FIREBASE_DB_URL}/campaigns.json`;

async function firebaseRead() {
  const res = await fetch(DB_PATH);
  if (!res.ok) throw new Error('Firebase read failed');
  const data = await res.json();
  if (!data) return null;
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
// Touch drag-and-drop hook — works on iOS Safari + Android + desktop
// ─────────────────────────────────────────────────────────────────────────────

function useTouchDrag({ campaigns, viewYear, viewMonth, onDropComplete }) {
  // Refs so touch handlers always see fresh values without re-registering
  const dragIdRef      = useRef(null);
  const ghostRef       = useRef(null);
  const dragOverDayRef = useRef(null);

  const [dragId,      setDragId]      = useState(null);
  const [dragOverDay, setDragOverDay] = useState(null);

  // Find which calendar day cell is under a touch point
  const getDayUnderPoint = useCallback((x, y) => {
    // Temporarily hide ghost so it doesn't interfere with elementFromPoint
    if (ghostRef.current) ghostRef.current.style.display = 'none';
    const el = document.elementFromPoint(x, y);
    if (ghostRef.current) ghostRef.current.style.display = '';
    if (!el) return null;
    // Walk up DOM to find an element with data-day attribute
    const cell = el.closest('[data-day]');
    return cell ? parseInt(cell.getAttribute('data-day'), 10) : null;
  }, []);

  // Create floating ghost element that follows the finger
  const createGhost = useCallback((sourceEl, x, y) => {
    const rect = sourceEl.getBoundingClientRect();
    const ghost = sourceEl.cloneNode(true);
    ghost.style.cssText = `
      position: fixed;
      left: ${rect.left}px;
      top: ${rect.top}px;
      width: ${rect.width}px;
      pointer-events: none;
      opacity: 0.85;
      z-index: 9999;
      transform: scale(1.06);
      box-shadow: 0 8px 24px rgba(0,0,0,0.25);
      border-radius: 6px;
      transition: transform 0.1s;
    `;
    document.body.appendChild(ghost);
    ghostRef.current = ghost;
  }, []);

  const removeGhost = useCallback(() => {
    if (ghostRef.current) {
      ghostRef.current.remove();
      ghostRef.current = null;
    }
  }, []);

  const onTouchStart = useCallback((e, campaignId, sourceEl) => {
    // Long-press detection: only start drag after 300ms hold
    const timer = setTimeout(() => {
      dragIdRef.current = campaignId;
      setDragId(campaignId);
      const touch = e.touches[0];
      createGhost(sourceEl, touch.clientX, touch.clientY);
      // Vibrate on supported devices (Android) to confirm drag started
      if (navigator.vibrate) navigator.vibrate(40);
    }, 300);

    // If user moves finger before 300ms, cancel — it's a scroll
    const cancelTimer = () => clearTimeout(timer);
    sourceEl.addEventListener('touchmove', cancelTimer, { once: true });
    sourceEl.addEventListener('touchend', cancelTimer, { once: true });
  }, [createGhost]);

  const onTouchMove = useCallback((e) => {
    if (!dragIdRef.current || !ghostRef.current) return;
    e.preventDefault(); // prevent page scroll while dragging
    const touch = e.touches[0];

    // Move ghost to follow finger
    const ghost = ghostRef.current;
    const w = parseFloat(ghost.style.width);
    ghost.style.left = `${touch.clientX - w / 2}px`;
    ghost.style.top  = `${touch.clientY - 20}px`;

    // Highlight target cell
    const day = getDayUnderPoint(touch.clientX, touch.clientY);
    if (day !== dragOverDayRef.current) {
      dragOverDayRef.current = day;
      setDragOverDay(day);
    }
  }, [getDayUnderPoint]);

  const onTouchEnd = useCallback((e) => {
    if (!dragIdRef.current) return;
    const touch = e.changedTouches[0];
    const day = getDayUnderPoint(touch.clientX, touch.clientY);

    if (day && day >= 1 && day <= getDaysInMonth(viewYear, viewMonth)) {
      const updated = campaigns.map(c =>
        c.id === dragIdRef.current
          ? { ...c, date: { year: viewYear, month: viewMonth, day } }
          : c
      );
      onDropComplete(updated);
    }

    dragIdRef.current = null;
    dragOverDayRef.current = null;
    setDragId(null);
    setDragOverDay(null);
    removeGhost();
  }, [campaigns, viewYear, viewMonth, getDayUnderPoint, onDropComplete, removeGhost]);

  // Attach touchmove/touchend to window so they fire even if finger leaves the card
  useEffect(() => {
    window.addEventListener('touchmove', onTouchMove, { passive: false });
    window.addEventListener('touchend',  onTouchEnd,  { passive: false });
    return () => {
      window.removeEventListener('touchmove', onTouchMove);
      window.removeEventListener('touchend',  onTouchEnd);
    };
  }, [onTouchMove, onTouchEnd]);

  return { dragId, dragOverDay, onTouchStart };
}

// ─────────────────────────────────────────────────────────────────────────────
// Main Component
// ─────────────────────────────────────────────────────────────────────────────

export default function App() {
  const [campaigns,   setCampaigns]   = useState([]);
  const [dbStatus,    setDbStatus]    = useState('loading');
  const [dbError,     setDbError]     = useState('');
  const [hasUnsaved,  setHasUnsaved]  = useState(false);

  const [viewYear,    setViewYear]    = useState(today.getFullYear());
  const [viewMonth,   setViewMonth]   = useState(today.getMonth());
  const [filterCat,   setFilterCat]   = useState('All Categories');
  const [modal,       setModal]       = useState(false);
  const [editId,      setEditId]      = useState(null);
  const [form,        setForm]        = useState({
    name:'', subtitle:'', notes:'', category:CATEGORIES[1],
    color:CAMPAIGN_COLORS[0].value, textColor:CAMPAIGN_COLORS[0].text, day:'',
  });

  // Desktop drag state
  const [desktopDragId,      setDesktopDragId]      = useState(null);
  const [desktopDragOverDay, setDesktopDragOverDay] = useState(null);

  const [exporting,   setExporting]   = useState(null);
  const [showWeekends, setShowWeekends] = useState(true);

  // ── Firebase load ─────────────────────────────────────────────────────────
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

  // ── Firebase save ─────────────────────────────────────────────────────────
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

  const updateLocal = useCallback((newList) => {
    setCampaigns(newList);
    setHasUnsaved(true);
    if (dbStatus === 'saved') setDbStatus('ready');
  }, [dbStatus]);

  // ── Touch drag-and-drop ───────────────────────────────────────────────────
  const { dragId: touchDragId, dragOverDay: touchDragOverDay, onTouchStart } = useTouchDrag({
    campaigns,
    viewYear,
    viewMonth,
    onDropComplete: updateLocal,
  });

  // Merge desktop + touch drag state for UI highlights
  const dragId      = desktopDragId      || touchDragId;
  const dragOverDay = desktopDragOverDay || touchDragOverDay;

  const COLS     = showWeekends ? 7 : 5;
  const DAY_NAMES = showWeekends ? DAY_NAMES_ALL : DAY_NAMES_WD;
  const calDays     = buildCalendarDays(viewYear, viewMonth, showWeekends);
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

  // ── Desktop drag & drop ───────────────────────────────────────────────────
  const onDragStart = (e, id) => { setDesktopDragId(id); e.dataTransfer.effectAllowed='move'; };
  const onDragOver  = (e, day) => { if(!day) return; e.preventDefault(); setDesktopDragOverDay(day); };
  const onDrop      = (e, day) => {
    e.preventDefault();
    if (desktopDragId && day) {
      updateLocal(campaigns.map(c => c.id===desktopDragId ? {...c, date:{year:viewYear,month:viewMonth,day}} : c));
    }
    setDesktopDragId(null); setDesktopDragOverDay(null);
  };
  const onDragEnd = () => { setDesktopDragId(null); setDesktopDragOverDay(null); };

  // ── Shared export grid builder ────────────────────────────────────────────
  // Returns the same cell array that the screen renders, so exports always match.
  // Each entry: { day: number, col: number, row: number } or null (empty pad cell)
  const buildExportGrid = () => {
    // Re-use the same buildCalendarDays logic but return positional info
    const cells = buildCalendarDays(viewYear, viewMonth, showWeekends);
    const numCols = showWeekends ? 7 : 5;
    return {
      cells,         // array of {day,dow}|null  — same as calDays
      numCols,
      numRows: cells.length / numCols,
      dayHeaders: showWeekends ? DAY_NAMES_ALL : DAY_NAMES_WD,
      exportList: filterCat === 'All Categories'
        ? campaigns.filter(c => c.date.year===viewYear && c.date.month===viewMonth)
        : campaigns.filter(c => c.date.year===viewYear && c.date.month===viewMonth && c.category===filterCat),
    };
  };

  // ── PPTX Export ───────────────────────────────────────────────────────────
  const exportPptx = async () => {
    setExporting('pptx');
    try {
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';

      const { cells, numCols, numRows, dayHeaders, exportList } = buildExportGrid();
      const W=13.33, H=7.5, MAR=0.25, HDR=0.55, DLBL=0.25;
      const GRID_TOP = MAR+HDR+DLBL;
      const CW = (W-MAR*2) / numCols;
      const CH = (H-GRID_TOP-MAR) / numRows;

      const slide = pptx.addSlide();

      // Background
      slide.addShape(pptx.ShapeType.rect, {x:0,y:0,w:W,h:H, fill:{color:'F7F8FA'}, line:{color:'F7F8FA'}});

      // Header bar
      slide.addShape(pptx.ShapeType.rect, {x:0,y:0,w:W,h:HDR, fill:{color:'1a1d2e'}, line:{color:'1a1d2e'}});
      slide.addText(`${MONTH_NAMES[viewMonth]} ${viewYear}  ·  Campaign Calendar`, {
        x:MAR, y:0.07, w:9, h:HDR-0.14, fontSize:19, bold:true, color:'FFFFFF', fontFace:'Calibri', valign:'middle',
      });
      if (filterCat !== 'All Categories') {
        slide.addText(`Category: ${filterCat}`, {
          x:W-3.6, y:0.07, w:3.3, h:HDR-0.14, fontSize:10, color:'A8D4E8', fontFace:'Calibri', align:'right', valign:'middle',
        });
      }

      // Day header row
      dayHeaders.forEach((d, i) => {
        const x = MAR + i * CW;
        slide.addShape(pptx.ShapeType.rect, {x, y:MAR+HDR, w:CW, h:DLBL, fill:{color:'2d3250'}, line:{color:'2d3250'}});
        slide.addText(d, {x, y:MAR+HDR, w:CW, h:DLBL, fontSize:8, bold:true, color:'FFFFFF', fontFace:'Calibri', align:'center', valign:'middle'});
      });

      // Grid cells — iterate using the same cells array as the screen
      cells.forEach((cell, i) => {
        const col = i % numCols;
        const row = Math.floor(i / numCols);
        const cx = MAR + col * CW;
        const cy = GRID_TOP + row * CH;
        const day = cell ? cell.day : null;

        slide.addShape(pptx.ShapeType.rect, {
          x:cx, y:cy, w:CW, h:CH,
          fill:{color: day ? 'FFFFFF' : 'F3F4F8'},
          line:{color:'E0E3EC', pt:0.5},
        });

        if (!day) return;

        const isTdy = viewYear===today.getFullYear() && viewMonth===today.getMonth() && day===today.getDate();
        if (isTdy) {
          slide.addShape(pptx.ShapeType.ellipse, {x:cx+0.05,y:cy+0.04,w:0.19,h:0.19, fill:{color:'3B5BDB'}, line:{color:'3B5BDB'}});
          slide.addText(String(day), {x:cx+0.05,y:cy+0.04,w:0.19,h:0.19, fontSize:7.5, bold:true, color:'FFFFFF', fontFace:'Calibri', align:'center', valign:'middle'});
        } else {
          slide.addText(String(day), {x:cx+0.06,y:cy+0.03,w:0.19,h:0.17, fontSize:7.5, bold:true, color:'4a5068', fontFace:'Calibri'});
        }

        const dayCamps = exportList.filter(c => c.date.day === day);
        const BOX_H=0.28, GAP=0.03;
        dayCamps.forEach((c, ci) => {
          const by = cy + 0.24 + ci*(BOX_H+GAP);
          if (by + BOX_H > cy + CH - 0.03) return;
          const bg = c.color.replace('#','');
          const fg = c.textColor.replace('#','');
          // Use plain rect — roundRectangle has inconsistent support across pptxgenjs versions
          slide.addShape(pptx.ShapeType.rect, {x:cx+0.04,y:by,w:CW-0.08,h:BOX_H, fill:{color:bg}, line:{color:bg}});
          slide.addText(c.name, {x:cx+0.07,y:by+0.01,w:CW-0.14,h:0.13, fontSize:7, bold:true, color:fg, fontFace:'Calibri', valign:'middle'});
          if (c.subtitle) slide.addText(c.subtitle, {x:cx+0.07,y:by+0.13,w:CW-0.14,h:0.09, fontSize:5.5, color:fg, fontFace:'Calibri'});
          if (c.notes)    slide.addText(c.notes,    {x:cx+0.07,y:by+0.20,w:CW-0.14,h:0.08, fontSize:5,   color:fg, fontFace:'Calibri', italic:true});
        });
      });

      // Footer
      slide.addText(`Exported ${new Date().toLocaleDateString()}  ·  Campaign Calendar`, {
        x:MAR, y:H-0.2, w:W-MAR*2, h:0.17, fontSize:6.5, color:'8a90a0', fontFace:'Calibri', align:'right',
      });

      await pptx.writeFile({fileName:`Campaign_Calendar_${MONTH_NAMES[viewMonth]}_${viewYear}`});
    } catch(err) {
      console.error('PPTX error:', err);
      alert(`PPTX export failed: ${err.message}`);
    }
    setExporting(null);
  };

  // ── PDF Export ────────────────────────────────────────────────────────────
  const exportPdf = async () => {
    setExporting('pdf');
    try {
      await loadScriptWithGlobal('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js', 'jspdf');
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({orientation:'landscape', unit:'pt', format:'a4'});

      const { cells, numCols, numRows, dayHeaders, exportList } = buildExportGrid();
      const PW=841.89, PH=595.28, MAR=18, HDR=38, DLBL=18;
      const GRID_TOP = MAR+HDR+DLBL;
      const CW = (PW-MAR*2) / numCols;
      const CH = (PH-GRID_TOP-MAR) / numRows;

      // Header bar
      doc.setFillColor(26,29,46);
      doc.rect(0, 0, PW, HDR, 'F');
      doc.setTextColor(255,255,255);
      doc.setFontSize(16); doc.setFont('helvetica','bold');
      doc.text(`${MONTH_NAMES[viewMonth]} ${viewYear}  ·  Campaign Calendar`, MAR, HDR*0.68);
      if (filterCat !== 'All Categories') {
        doc.setFontSize(9); doc.setFont('helvetica','normal'); doc.setTextColor(168,212,232);
        doc.text(`Category: ${filterCat}`, PW-MAR, HDR*0.68, {align:'right'});
      }

      // Day header row
      doc.setFillColor(45,50,80);
      doc.rect(MAR, MAR+HDR, PW-MAR*2, DLBL, 'F');
      dayHeaders.forEach((d, i) => {
        doc.setFontSize(8); doc.setFont('helvetica','bold'); doc.setTextColor(255,255,255);
        doc.text(d, MAR + i*CW + CW/2, MAR+HDR+DLBL*0.72, {align:'center'});
      });

      // Grid cells — same cells array as screen & PPTX
      cells.forEach((cell, i) => {
        const col = i % numCols;
        const row = Math.floor(i / numCols);
        const cx = MAR + col * CW;
        const cy = GRID_TOP + row * CH;
        const day = cell ? cell.day : null;

        doc.setFillColor(...(day ? [255,255,255] : [243,244,248]));
        doc.setDrawColor(224,227,236); doc.setLineWidth(0.4);
        doc.rect(cx, cy, CW, CH, 'FD');

        if (!day) return;

        const isTdy = viewYear===today.getFullYear() && viewMonth===today.getMonth() && day===today.getDate();
        if (isTdy) {
          doc.setFillColor(59,91,219);
          doc.circle(cx+9, cy+9, 7, 'F');
          doc.setFontSize(7); doc.setFont('helvetica','bold'); doc.setTextColor(255,255,255);
          doc.text(String(day), cx+9, cy+12, {align:'center'});
        } else {
          doc.setFontSize(7.5); doc.setFont('helvetica','bold'); doc.setTextColor(74,80,104);
          doc.text(String(day), cx+5, cy+12);
        }

        const dayCamps = exportList.filter(c => c.date.day === day);
        const LINE_NAME=8, LINE_SUB=7, LINE_NOTES=6.5, PAD_TOP=3, PAD_BOT=3, GAP=2;
        let cardY = cy + 16;
        dayCamps.forEach((c) => {
          // Calculate card height based on how many lines this card has
          const lines = 1 + (c.subtitle ? 1 : 0) + (c.notes ? 1 : 0);
          const BH = PAD_TOP + LINE_NAME + (c.subtitle ? LINE_SUB : 0) + (c.notes ? LINE_NOTES : 0) + PAD_BOT;
          if (cardY + BH > cy + CH - 2) return; // overflow guard
          const {r,g,b} = hexToObj(c.color);
          doc.setFillColor(r,g,b); doc.setDrawColor(r,g,b);
          doc.roundedRect(cx+3, cardY, CW-6, BH, 2, 2, 'FD');
          const {r:tr,g:tg,b:tb} = hexToObj(c.textColor);
          doc.setTextColor(tr,tg,tb);
          let textY = cardY + PAD_TOP + LINE_NAME - 1;
          doc.setFontSize(7); doc.setFont('helvetica','bold');
          doc.text(c.name.slice(0,22), cx+5, textY);
          if (c.subtitle) {
            textY += LINE_SUB;
            doc.setFontSize(5.5); doc.setFont('helvetica','normal');
            doc.text(c.subtitle.slice(0,26), cx+5, textY);
          }
          if (c.notes) {
            textY += LINE_NOTES;
            doc.setFontSize(5); doc.setFont('helvetica','italic');
            doc.text(c.notes.slice(0,28), cx+5, textY);
          }
          cardY += BH + GAP;
        });
      });

      // Footer
      doc.setFontSize(6.5); doc.setFont('helvetica','normal'); doc.setTextColor(138,144,160);
      doc.text(`Exported ${new Date().toLocaleDateString()}  ·  Campaign Calendar`, PW-MAR, PH-6, {align:'right'});
      doc.save(`Campaign_Calendar_${MONTH_NAMES[viewMonth]}_${viewYear}.pdf`);
    } catch(err) {
      console.error('PDF error:', err);
      alert(`PDF export failed: ${err.message}`);
    }
    setExporting(null);
  };

  // ── Status bar ────────────────────────────────────────────────────────────
  const statusConfig = {
    loading: { bg:'#f0f4ff', border:'#c7d4f8', color:'#3B5BDB', text:'⏳ Loading calendar data…',                                       showSave:false },
    ready:   { bg: hasUnsaved?'#fff8e1':'#f0fdf4', border: hasUnsaved?'#fde68a':'#bbf7d0', color: hasUnsaved?'#92400e':'#166534',
               text: hasUnsaved?'● Unsaved changes — click Save to publish for all users':'✓ All changes saved',                         showSave: hasUnsaved },
    saving:  { bg:'#f0f4ff', border:'#c7d4f8', color:'#3B5BDB', text:'💾 Saving to database…',                                          showSave:false },
    saved:   { bg:'#f0fdf4', border:'#bbf7d0', color:'#166534', text:'✅ Saved! All users will see these changes.',                      showSave:false },
    error:   { bg:'#fef2f2', border:'#fecaca', color:'#991b1b', text:`⚠ ${dbError}`,                                                   showSave:true  },
  };
  const status = statusConfig[dbStatus] || statusConfig.ready;

  // ─────────────────────────────────────────────────────────────────────────
  // Render
  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div style={{fontFamily:"'Inter','Segoe UI',sans-serif",minHeight:'100vh',background:'#F7F8FA',padding:16}}>
      <style>{`
        .camp-card { transition: transform 0.1s, filter 0.1s; touch-action: none; }
        .camp-card:hover { filter: brightness(0.91); transform: translateY(-1px); }
        .cal-cell:hover { background: #F3F5FF !important; }
        /* Prevent iOS Safari rubber-band scroll while dragging */
        body.is-dragging { overflow: hidden; }
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
            <button onClick={()=>saveToDatabase(campaigns)} disabled={dbStatus==='saving'}
              style={{background:'#1a1d2e',color:'#fff',border:'none',borderRadius:7,padding:'8px 22px',fontSize:13,fontWeight:700,cursor:'pointer',boxShadow:'0 2px 8px rgba(0,0,0,0.18)'}}>
              💾 Save Changes
            </button>
          )}
        </div>

        {/* ── Top bar ── */}
        <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',marginBottom:14,flexWrap:'wrap',gap:10}}>
          <div>
            <h1 style={{margin:0,fontSize:21,fontWeight:800,color:'#1a1d2e',letterSpacing:'-0.4px'}}>Campaign Calendar</h1>
            <p style={{margin:'3px 0 0',fontSize:12,color:'#8a90a0'}}>
              Desktop: drag cards · Mobile: <strong>hold &amp; drag</strong> · Then hit Save
            </p>
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
            <button
              onClick={()=>setShowWeekends(s=>!s)}
              style={{...btnSt, background: showWeekends?'#f0f1f5':'#e0e7ff', color: showWeekends?'#4a5068':'#3B5BDB', border: showWeekends?'1.5px solid #e0e3ec':'1.5px solid #3B5BDB'}}>
              {showWeekends ? '5-Day View' : '7-Day View'}
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

        {/* ── Calendar grid ── */}
        <div style={{background:'#fff',borderRadius:14,boxShadow:'0 2px 16px rgba(0,0,0,0.07)',overflow:'hidden'}}>
          {/* Day headers */}
          <div style={{display:'grid',gridTemplateColumns:`repeat(${COLS},1fr)`,background:'#2d3250'}}>
            {DAY_NAMES.map(d=>(
              <div key={d} style={{padding:'9px 0',textAlign:'center',fontSize:11.5,fontWeight:700,color:'#ffffff',letterSpacing:'0.6px',textTransform:'uppercase'}}>{d}</div>
            ))}
          </div>

          {/* Cells */}
          <div style={{display:'grid',gridTemplateColumns:`repeat(${COLS},1fr)`}}>
            {calDays.map((cell,i)=>{
              const day      = cell ? cell.day : null;
              const isToday  = day&&viewYear===today.getFullYear()&&viewMonth===today.getMonth()&&day===today.getDate();
              const isOver   = dragOverDay===day&&day!==null;
              const dayCamps = day ? forDay(day) : [];
              return (
                <div
                  key={i}
                  data-day={day||''}
                  className={day?'cal-cell':''}
                  onClick={()=>day&&!dragId&&openNew(day)}
                  onDragOver={e=>onDragOver(e,day)}
                  onDrop={e=>onDrop(e,day)}
                  style={{
                    minHeight:118,
                    borderRight:(i+1)%COLS===0?'none':'1px solid #eef0f5',
                    borderBottom:i<calDays.length-COLS?'1px solid #eef0f5':'none',
                    padding:'6px 5px 5px',
                    cursor:day?'pointer':'default',
                    background:isOver?'#DBEAFE':day?'#fff':'#FAFBFC',
                    transition:'background 0.12s',
                    outline:isOver?'2px solid #3B5BDB':'none',
                  }}>
                  {day&&(
                    <>
                      <div style={{fontSize:11.5,fontWeight:700,color:isToday?'#fff':'#4a5068',background:isToday?'#3B5BDB':'transparent',borderRadius:'50%',width:22,height:22,display:'flex',alignItems:'center',justifyContent:'center',marginBottom:4}}>
                        {day}
                      </div>
                      <div style={{display:'flex',flexDirection:'column',gap:3}}>
                        {dayCamps.map(c=>(
                          <div
                            key={c.id}
                            className="camp-card"
                            draggable
                            onDragStart={e=>{e.stopPropagation();onDragStart(e,c.id);}}
                            onDragEnd={onDragEnd}
                            onTouchStart={e=>{e.stopPropagation();onTouchStart(e,c.id,e.currentTarget);}}
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

      </div>

      {/* ── Modal ── */}
      {modal&&(
        <div style={{position:'fixed',inset:0,background:'rgba(10,15,40,0.48)',display:'flex',alignItems:'center',justifyContent:'center',zIndex:200}}
          onClick={()=>setModal(false)}>
          <div style={{background:'#fff',borderRadius:14,padding:26,width:440,maxWidth:'93vw',boxShadow:'0 24px 70px rgba(0,0,0,0.22)',maxHeight:'92vh',overflowY:'auto'}}
            onClick={e=>e.stopPropagation()}>
            <h3 style={{margin:'0 0 16px',fontSize:16,fontWeight:700,color:'#1a1d2e'}}>{editId?'Edit Campaign':'New Campaign'}</h3>

            <Lbl>Campaign Name *</Lbl>
            <input value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} style={inpSt}/>

            <Lbl>Subtitle</Lbl>
            <input value={form.subtitle} onChange={e=>setForm(f=>({...f,subtitle:e.target.value}))} style={inpSt}/>

            <Lbl>Notes</Lbl>
            <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))}
              rows={2}
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
