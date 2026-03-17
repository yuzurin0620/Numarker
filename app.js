// ─── Constants & State ────────────────────────────────────────────────────────
const STAGE_PAD = 300;
const COLORS = ['#ef4444','#f97316','#eab308','#22c55e','#3b82f6','#a855f7','#ec4899','#14b8a6','#ffffff','#1f2937'];

let annotations = [];
let history = [];
let historyIndex = -1;
let nextId = 1;
let nextNum = 1;
let activeTool = 'select';

const colorByType = { number: '#ef4444', rect: '#ef4444', arrow: '#ef4444' };
let activeColorType = 'number';
let selectedIds = new Set();
let imageLoaded = false;
let baseStageW = 800, baseStageH = 600;
let viewScale = 1.0;

let defaultRadius = 20;
let defaultRectStroke = 3;
let defaultArrowStroke = 3;
let defaultArrowType = 'arrow';
let defaultFillOpacitySolid = 0;
let defaultFillOpacityHatch = 0.25;
let defaultFillOpacity = 0;
let defaultFillHatch = false;
let bgColor = 'transparent';
let bgRect = null;
let defaultNumStyle = 'circle';

let stage, imageLayer, annotationLayer, drawLayer, transformer;
let isDrawing = false, drawStart = { x: 0, y: 0 }, tempShape = null;
let marqueeActive = false, marqueeStart = { x: 0, y: 0 }, marqueeRect = null;
let multiDragAnchorId = null, multiDragStartPos = {};
let arrowHandleNodes = [];
let notesVisible = false;
let shiftPressed = false;
let isPanning = false, panStartX = 0, panStartY = 0, panScrollLeft = 0, panScrollTop = 0;

// Cursor ghost for number tool
let cursorGhost = null;
let cursorGhostNeedsRebuild = true;
let lastStagePointer = null;

// Internal clipboard for copy/paste
let clipboard = [];

// Edge alignment state
let edgeSelections = {}; // { annId: 'top'|'bottom'|'left'|'right' }
let edgeHandleNodes = [];

// Prevent re-entrant history saves during undo/redo
let isRestoringHistory = false;

// ─── Helpers ──────────────────────────────────────────────────────────────────
function hexToRgba(hex, alpha) {
  const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function makeHatchCanvas(color, opacity) {
  const c = document.createElement('canvas');
  c.width = 8; c.height = 8;
  const ctx = c.getContext('2d');
  ctx.strokeStyle = hexToRgba(color, opacity != null ? opacity : 1); ctx.lineWidth = 1.5;
  ctx.beginPath(); ctx.moveTo(-1,9); ctx.lineTo(9,-1); ctx.moveTo(-1,17); ctx.lineTo(17,-1); ctx.stroke();
  return c;
}

function isDarkColor(hex) { return hex !== '#ffffff' && hex !== '#eab308'; }
function textColor(color) { return isDarkColor(color) ? '#ffffff' : '#111827'; }
function findAnn(id) { return annotations.find(a => a.id === id); }

function normalizeNumericInput(val) {
  let s = String(val).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  s = s.replace(/[^0-9\-]/g, '');
  if (s === '' || s === '-') return null;
  const n = parseInt(s);
  return isNaN(n) ? null : n;
}

// ─── Zoom ────────────────────────────────────────────────────────────────────
function setZoom(newScale) {
  if (Math.abs(newScale - 1.0) < 0.05) newScale = 1.0;
  viewScale = Math.max(0.1, Math.min(4.0, newScale));
  if (!imageLoaded) {
    document.getElementById('zoom-slider').value = Math.round(viewScale * 100);
    document.getElementById('zoom-label').textContent = Math.round(viewScale * 100) + '%';
    return;
  }
  const totalW = Math.round((baseStageW + 2 * STAGE_PAD) * viewScale);
  const totalH = Math.round((baseStageH + 2 * STAGE_PAD) * viewScale);
  stage.width(totalW); stage.height(totalH);
  stage.scale({ x: viewScale, y: viewScale });
  const container = document.getElementById('konva-container');
  container.style.width = totalW + 'px';
  container.style.height = totalH + 'px';
  document.getElementById('zoom-slider').value = Math.round(viewScale * 100);
  document.getElementById('zoom-label').textContent = Math.round(viewScale * 100) + '%';
  annotationLayer.batchDraw();
}

// ─── Color Type Tabs ──────────────────────────────────────────────────────────
function setActiveColorTab(type) {
  activeColorType = type;
  document.querySelectorAll('.color-type-tab').forEach(t => {
    const isA = t.dataset.type === type;
    t.classList.toggle('active', isA);
    t.style.borderColor = isA ? '#ef4444' : '#374151';
    t.style.color = isA ? '#e5e7eb' : '#9ca3af';
  });
  syncPaletteToColor(colorByType[type]);
  if (/^#[0-9a-fA-F]{6}$/.test(colorByType[type])) {
    document.getElementById('custom-color').value = colorByType[type];
  }
}

function updateColorDots() {
  document.getElementById('dot-number').style.background = colorByType.number;
  document.getElementById('dot-rect').style.background = colorByType.rect;
  document.getElementById('dot-arrow').style.background = colorByType.arrow;
}

// ─── Init Konva ──────────────────────────────────────────────────────────────
function initKonva() {
  const w = (baseStageW + 2 * STAGE_PAD) * viewScale;
  const h = (baseStageH + 2 * STAGE_PAD) * viewScale;
  stage = new Konva.Stage({ container: 'konva-container', width: w, height: h });
  stage.scale({ x: viewScale, y: viewScale });
  const container = document.getElementById('konva-container');
  container.style.width = w + 'px'; container.style.height = h + 'px';
  imageLayer = new Konva.Layer();
  annotationLayer = new Konva.Layer();
  drawLayer = new Konva.Layer();
  stage.add(imageLayer, annotationLayer, drawLayer);

  transformer = new Konva.Transformer({
    rotateEnabled: false, keepRatio: false,
    boundBoxFunc: (o, n) => (Math.abs(n.width) < 5 || Math.abs(n.height) < 5) ? o : n
  });
  annotationLayer.add(transformer);

  transformer.on('transformstart', () => { hideEdgeHandles(); });
  transformer.on('transformend', () => {
    transformer.nodes().forEach(node => {
      const ann = findAnn(node.id());
      if (ann && ann.type === 'rect') {
        ann.x = node.x() - STAGE_PAD; ann.y = node.y() - STAGE_PAD;
        ann.width = Math.abs(node.width() * node.scaleX());
        ann.height = Math.abs(node.height() * node.scaleY());
        node.width(ann.width); node.height(ann.height);
        node.scaleX(1); node.scaleY(1);
      }
    });
    saveHistory(); annotationLayer.batchDraw(); showEdgeHandles();
  });

  stage.on('mousedown touchstart', onStageDown);
  stage.on('mousemove touchmove', onStageMove);
  stage.on('mouseup touchend', onStageUp);
  stage.on('mouseleave', () => {
    if (cursorGhost) { cursorGhost.destroy(); cursorGhost = null; drawLayer.batchDraw(); }
  });
}

// ─── Pointer helpers ─────────────────────────────────────────────────────────
function getStagePointer() {
  const raw = stage.getPointerPosition() || { x: 0, y: 0 };
  return { x: raw.x / viewScale, y: raw.y / viewScale };
}
function getPointer() {
  const sp = getStagePointer();
  return { x: sp.x - STAGE_PAD, y: sp.y - STAGE_PAD };
}

// ─── Image Loading ────────────────────────────────────────────────────────────
function loadImage(src) {
  const img = new Image();
  img.onload = () => {
    imageLayer.destroyChildren();
    const area = document.getElementById('canvas-area');
    const maxW = area.clientWidth - 40, maxH = area.clientHeight - 40;
    let w = img.width, h = img.height;
    const sc = Math.min(maxW / w, maxH / h, 1);
    w = Math.round(w * sc); h = Math.round(h * sc);
    baseStageW = w; baseStageH = h;
    setZoom(viewScale);
    const kImg = new Konva.Image({ image: img, x: STAGE_PAD, y: STAGE_PAD, width: w, height: h });
    imageLayer.add(kImg); imageLayer.batchDraw();
    imageLoaded = true;
    document.getElementById('drop-hint').classList.add('hidden');
    const totalW = (w + 2*STAGE_PAD)*viewScale, totalH = (h + 2*STAGE_PAD)*viewScale;
    const PAD_PX = 500;
    document.getElementById('canvas-pan-pad').style.padding = PAD_PX + 'px';
    area.scrollLeft = (totalW + 2*PAD_PX - area.clientWidth) / 2;
    area.scrollTop = (totalH + 2*PAD_PX - area.clientHeight) / 2;
    renderAnnotations(); renderNotesPanel();
  };
  img.src = src;
}

function handleImageFile(file) {
  if (!file || !file.type.startsWith('image/')) return;
  if (imageLoaded && annotations.length > 0) {
    if (!confirm('編集中のアノテーションがあります。新しい画像を開くと失われます。続けますか？')) return;
    annotations = []; nextNum = 1; nextId = 1; selectedIds = new Set();
    saveHistory(); renderAnnotations(); renderNotesPanel();
  }
  const reader = new FileReader();
  reader.onload = e => loadImage(e.target.result);
  reader.readAsDataURL(file);
}

// ─── Arrow Handles ────────────────────────────────────────────────────────────
function showArrowHandles(ann) {
  hideArrowHandles();
  const makeHandle = (x, y, idx) => {
    const h = new Konva.Circle({ x, y, radius: 7, fill: 'white', stroke: '#ef4444', strokeWidth: 2, draggable: true, name: 'arrow-handle' });
    h.on('click tap', e => e.cancelBubble = true);
    h.on('dragmove', () => {
      ann.points[idx*2] = h.x() - STAGE_PAD; ann.points[idx*2+1] = h.y() - STAGE_PAD;
      const n = annotationLayer.findOne('#' + ann.id);
      if (n) n.points([ann.points[0]+STAGE_PAD, ann.points[1]+STAGE_PAD, ann.points[2]+STAGE_PAD, ann.points[3]+STAGE_PAD]);
      annotationLayer.batchDraw();
    });
    h.on('dragend', () => saveHistory());
    return h;
  };
  const h1 = makeHandle(ann.points[0]+STAGE_PAD, ann.points[1]+STAGE_PAD, 0);
  const h2 = makeHandle(ann.points[2]+STAGE_PAD, ann.points[3]+STAGE_PAD, 1);
  drawLayer.add(h1, h2); drawLayer.batchDraw();
  arrowHandleNodes = [h1, h2];
}

function hideArrowHandles() {
  arrowHandleNodes.forEach(n => n.destroy());
  arrowHandleNodes = []; drawLayer.batchDraw();
}

// ─── Edge Alignment Handles ───────────────────────────────────────────────────
function showEdgeHandles() {
  hideEdgeHandles();
  const selRects = [...selectedIds].map(id => findAnn(id)).filter(a => a && a.type === 'rect');
  if (selRects.length < 2) return;
  selRects.forEach(ann => {
    const cx = ann.x + STAGE_PAD, cy = ann.y + STAGE_PAD;
    const w = ann.width, h = ann.height;
    const handleDefs = [
      { dir: 'top',    x: cx + w/2, y: cy },
      { dir: 'bottom', x: cx + w/2, y: cy + h },
      { dir: 'left',   x: cx,       y: cy + h/2 },
      { dir: 'right',  x: cx + w,   y: cy + h/2 },
    ];
    handleDefs.forEach(({ dir, x, y }) => {
      const isSelected = edgeSelections[ann.id] === dir;
      const isV = dir === 'top' || dir === 'bottom';
      const h_node = new Konva.Circle({
        x, y, radius: 6,
        fill: isSelected ? '#facc15' : '#1f2937',
        stroke: isSelected ? '#facc15' : '#9ca3af',
        strokeWidth: 2,
        name: 'edge-handle',
        draggable: isSelected,
      });
      h_node.on('click tap', e => {
        e.cancelBubble = true;
        toggleEdgeSelection(ann.id, dir);
      });
      if (isSelected) {
        let handleStart = null;
        let rectsStart = {};
        h_node.on('dragstart', e => {
          e.cancelBubble = true;
          handleStart = { x: h_node.x(), y: h_node.y() };
          Object.keys(edgeSelections).forEach(id => {
            const a = findAnn(id);
            if (a) rectsStart[id] = { x: a.x, y: a.y, width: a.width, height: a.height };
          });
        });
        h_node.on('dragmove', () => {
          if (!handleStart) return;
          if (isV) h_node.x(handleStart.x); else h_node.y(handleStart.y);
          const dy = h_node.y() - handleStart.y;
          const dx = h_node.x() - handleStart.x;
          const MIN = 4;
          Object.entries(edgeSelections).forEach(([id, d]) => {
            const a = findAnn(id); const orig = rectsStart[id];
            if (!a || !orig) return;
            const node = annotationLayer.findOne('#' + id);
            if (d === 'top') {
              const newH = Math.max(MIN, orig.height - dy);
              a.y = orig.y + (orig.height - newH); a.height = newH;
              if (node) { node.y(a.y + STAGE_PAD); node.height(a.height); }
            } else if (d === 'bottom') {
              a.height = Math.max(MIN, orig.height + dy);
              if (node) node.height(a.height);
            } else if (d === 'left') {
              const newW = Math.max(MIN, orig.width - dx);
              a.x = orig.x + (orig.width - newW); a.width = newW;
              if (node) { node.x(a.x + STAGE_PAD); node.width(a.width); }
            } else if (d === 'right') {
              a.width = Math.max(MIN, orig.width + dx);
              if (node) node.width(a.width);
            }
          });
          annotationLayer.batchDraw();
        });
        h_node.on('dragend', () => {
          handleStart = null; rectsStart = {};
          saveHistory(); showEdgeHandles();
        });
      }
      drawLayer.add(h_node);
      edgeHandleNodes.push(h_node);
    });
  });
  drawLayer.batchDraw();
}

function hideEdgeHandles() {
  edgeHandleNodes.forEach(n => n.destroy());
  edgeHandleNodes = []; drawLayer.batchDraw();
}

function toggleEdgeSelection(annId, dir) {
  // All selected edges must be the exact same direction
  const existing = Object.values(edgeSelections);
  if (existing.length > 0 && existing[0] !== dir) edgeSelections = {};
  if (edgeSelections[annId] === dir) delete edgeSelections[annId];
  else edgeSelections[annId] = dir;
  showEdgeHandles();
  updatePropsPanel();
}


// ─── createNode ───────────────────────────────────────────────────────────────
function makeDragHandlers(ann, node) {
  node.on('dragstart', () => {
    if (activeTool !== 'select') { node.stopDrag(); return; }
    if (!selectedIds.has(ann.id)) {
      selectedIds = new Set([ann.id]);
      updateSelectionVisual(); updatePropsPanel();
    }
    hideEdgeHandles();
    multiDragAnchorId = ann.id;
    multiDragStartPos = {};
    selectedIds.forEach(id => {
      const n = annotationLayer.findOne('#' + id);
      if (n) multiDragStartPos[id] = { x: n.x(), y: n.y() };
    });
  });
  node.on('dragmove', () => {
    if (multiDragAnchorId !== ann.id) return;
    const startAnchor = multiDragStartPos[ann.id]; if (!startAnchor) return;
    let dx = node.x() - startAnchor.x, dy = node.y() - startAnchor.y;
    if (shiftPressed) {
      if (Math.abs(dx) >= Math.abs(dy)) dy = 0; else dx = 0;
      node.x(startAnchor.x + dx); node.y(startAnchor.y + dy);
    }
    selectedIds.forEach(id => {
      if (id === ann.id) return;
      const n = annotationLayer.findOne('#' + id), s = multiDragStartPos[id];
      if (n && s) { n.x(s.x + dx); n.y(s.y + dy); }
    });
    annotationLayer.batchDraw();
    if (bgColor !== 'transparent' && imageLoaded) {
      const tempUpdates = [];
      selectedIds.forEach(id => {
        const a = findAnn(id), n = annotationLayer.findOne('#' + id);
        if (!a || !n) return;
        if (a.type === 'number' || a.type === 'rect') {
          const orig = { x: a.x, y: a.y };
          a.x = n.x() - STAGE_PAD; a.y = n.y() - STAGE_PAD;
          tempUpdates.push({ a, orig });
        }
      });
      updateBgRect();
      tempUpdates.forEach(({ a, orig }) => { a.x = orig.x; a.y = orig.y; });
    }
  });
  node.on('dragend', () => {
    selectedIds.forEach(id => {
      const a = findAnn(id), n = annotationLayer.findOne('#' + id);
      if (!a || !n) return;
      if (a.type === 'number' || a.type === 'rect') {
        a.x = n.x() - STAGE_PAD; a.y = n.y() - STAGE_PAD;
      } else if (a.type === 'arrow') {
        const ddx = n.x(), ddy = n.y();
        if (ddx !== 0 || ddy !== 0) {
          a.points = [a.points[0]+ddx, a.points[1]+ddy, a.points[2]+ddx, a.points[3]+ddy];
          n.x(0); n.y(0); n.points([a.points[0]+STAGE_PAD,a.points[1]+STAGE_PAD,a.points[2]+STAGE_PAD,a.points[3]+STAGE_PAD]);
        }
      }
    });
    multiDragAnchorId = null; multiDragStartPos = {};
    saveHistory(); annotationLayer.batchDraw(); showEdgeHandles();
  });
  node.on('click tap', e => {
    e.cancelBubble = true;
    if (activeTool !== 'select') return;
    if (e.evt && e.evt.shiftKey) {
      if (selectedIds.has(ann.id)) selectedIds.delete(ann.id); else selectedIds.add(ann.id);
    } else {
      selectedIds = new Set([ann.id]);
    }
    updateSelectionVisual(); updatePropsPanel(); saveHistory();
    if (selectedIds.size === 1 && selectedIds.has(ann.id)) setActiveColorTab(ann.type);
  });
}

function createNode(ann) {
  if (ann.type === 'number') {
    const r = ann.radius || 20, style = ann.style || 'circle';
    const numStr = String(ann.num);
    const fsize = r * (numStr.length > 2 ? 0.7 : numStr.length > 1 ? 0.85 : 1.1);
    const isOutline = style === 'circle-outline';
    const tColor = isOutline ? ann.color : textColor(ann.color);
    const group = new Konva.Group({ id: ann.id, x: ann.x + STAGE_PAD, y: ann.y + STAGE_PAD, draggable: true });
    let shape;
    if (style === 'circle') {
      shape = new Konva.Circle({ x:0,y:0,radius:r,fill:ann.color,name:'shape' });
    } else if (style === 'square') {
      shape = new Konva.Rect({ x:-r,y:-r,width:r*2,height:r*2,fill:ann.color,name:'shape' });
    } else if (style === 'circle-outline') {
      shape = new Konva.Circle({ x:0,y:0,radius:r,fill:'transparent',stroke:ann.color,strokeWidth:2.5,name:'shape' });
    } else if (style === 'rounded-square') {
      shape = new Konva.Rect({ x:-r,y:-r,width:r*2,height:r*2,cornerRadius:r*0.4,fill:ann.color,name:'shape' });
    } else if (style === 'diamond') {
      shape = new Konva.Line({ points:[0,-r,r,0,0,r,-r,0],closed:true,fill:ann.color,name:'shape' });
    } else {
      shape = new Konva.Circle({ x:0,y:0,radius:r,fill:ann.color,name:'shape' });
    }
    const text = new Konva.Text({ x:-r,y:-r,width:r*2,height:r*2,text:numStr,fontSize:fsize,fontStyle:'bold',fill:tColor,align:'center',verticalAlign:'middle',name:'label' });
    group.add(shape, text);
    makeDragHandlers(ann, group);
    return group;
  }

  if (ann.type === 'rect') {
    const fillOpacity = ann.fillOpacity != null ? ann.fillOpacity : 0;
    const fillHatch = !!ann.fillHatch;
    let fillConfig = {};
    if (fillHatch && fillOpacity > 0) {
      const hc = makeHatchCanvas(ann.color, fillOpacity);
      fillConfig = { fillPriority: 'pattern', fillPatternImage: hc, fillPatternRepeat: 'repeat' };
    } else if (!fillHatch && fillOpacity > 0) {
      fillConfig = { fill: hexToRgba(ann.color, fillOpacity) };
    } else {
      fillConfig = { fill: 'transparent' };
    }
    const rect = new Konva.Rect({
      id: ann.id, x: ann.x + STAGE_PAD, y: ann.y + STAGE_PAD,
      width: ann.width, height: ann.height,
      stroke: ann.color, strokeWidth: ann.strokeWidth || defaultRectStroke,
      draggable: true, ...fillConfig
    });
    makeDragHandlers(ann, rect);
    return rect;
  }

  if (ann.type === 'arrow') {
    const arrowType = ann.arrowType || 'arrow';
    const pts = [ann.points[0]+STAGE_PAD, ann.points[1]+STAGE_PAD, ann.points[2]+STAGE_PAD, ann.points[3]+STAGE_PAD];
    const sw = ann.strokeWidth || defaultArrowStroke;
    const shape = new Konva.Arrow({
      id: ann.id, points: pts,
      stroke: ann.color, strokeWidth: sw,
      fill: ann.color,
      pointerLength: arrowType === 'line' ? 0 : Math.max(10, sw * 4),
      pointerWidth: arrowType === 'line' ? 0 : Math.max(8, sw * 3),
      draggable: true
    });
    makeDragHandlers(ann, shape);
    return shape;
  }
  return null;
}

// ─── Render ───────────────────────────────────────────────────────────────────
function renderAnnotations() {
  hideArrowHandles();
  annotationLayer.getChildren(n => n !== transformer).forEach(n => n.destroy());
  // Draw order: arrows → rects → numbers (numbers on top)
  const arrows = annotations.filter(a => a.type === 'arrow');
  const rects = annotations.filter(a => a.type === 'rect');
  const numbers = annotations.filter(a => a.type === 'number');
  [...arrows, ...rects, ...numbers].forEach(ann => {
    const node = createNode(ann);
    if (node) annotationLayer.add(node);
  });
  transformer.moveToTop();
  annotationLayer.batchDraw();
  updateSelectionVisual();
  updateBgRect();
}

function updateSelectionVisual() {
  transformer.nodes([]);
  hideArrowHandles();
  annotationLayer.find('Circle').forEach(c => {
    if (c.name() !== 'arrow-handle') {
      const p = c.getParent(); const pann = p ? findAnn(p.id()) : null;
      if (pann?.style === 'circle-outline') return;
      c.stroke(null); c.strokeWidth(0);
    }
  });
  annotationLayer.find('Rect').forEach(r => { r.dash([]); r.shadowEnabled(false); if(r.name()==='shape'){r.stroke(null);r.strokeWidth(0);} });
  annotationLayer.find('Line').forEach(l => { l.stroke(null); l.strokeWidth(0); });
  annotationLayer.find('Arrow').forEach(a => { a.shadowEnabled(false); });

  selectedIds.forEach(id => {
    const node = annotationLayer.findOne('#' + id);
    if (!node) return;
    const ann = findAnn(id); if (!ann) return;
    if (ann.type === 'number') {
      const shape = node.findOne('.shape');
      if (shape) {
        shape.stroke(ann.style === 'circle-outline' ? ann.color : '#fff');
        shape.strokeWidth(ann.style === 'circle-outline' ? 3.5 : 2.5);
      }
    } else if (ann.type === 'rect') {
      node.dash([6,3]); node.shadowColor('#ef4444'); node.shadowBlur(8); node.shadowEnabled(true);
    } else if (ann.type === 'arrow') {
      node.shadowColor('#ef4444'); node.shadowBlur(8); node.shadowEnabled(true);
    }
  });

  if (selectedIds.size === 1) {
    const id = [...selectedIds][0], ann = findAnn(id);
    if (ann && ann.type === 'rect') {
      const node = annotationLayer.findOne('#' + ann.id);
      if (node) { transformer.nodes([node]); node.dash([]); node.shadowEnabled(false); }
    } else if (ann && ann.type === 'arrow') {
      showArrowHandles(ann);
    }
  } else if (selectedIds.size >= 2) {
    const selAnns = [...selectedIds].map(id => findAnn(id)).filter(Boolean);
    const allRects = selAnns.length > 0 && selAnns.every(a => a.type === 'rect');
    if (allRects && Object.keys(edgeSelections).length === 0) {
      const nodes = selAnns.map(a => annotationLayer.findOne('#' + a.id)).filter(Boolean);
      nodes.forEach(n => { n.dash([]); n.shadowEnabled(false); });
      transformer.nodes(nodes);
    }
  }
  annotationLayer.batchDraw();
  updateRenumberHint();
  updateColorSectionVisibility();
  showEdgeHandles();
}

function deselectAll() {
  selectedIds = new Set();
  if (marqueeRect) { marqueeRect.destroy(); marqueeRect = null; drawLayer.batchDraw(); }
  marqueeActive = false;
  hideArrowHandles();
  transformer.nodes([]);
  annotationLayer.find('Circle').forEach(c => {
    if (c.name() !== 'arrow-handle') {
      const p = c.getParent(); const pann = p ? findAnn(p.id()) : null;
      if (pann?.style === 'circle-outline') return;
      c.stroke(null); c.strokeWidth(0);
    }
  });
  annotationLayer.find('Rect').forEach(r => { r.dash([]); r.shadowEnabled(false); if(r.name()==='shape'){r.stroke(null);r.strokeWidth(0);} });
  annotationLayer.find('Line').forEach(l => { l.stroke(null); l.strokeWidth(0); });
  annotationLayer.find('Arrow').forEach(a => { a.shadowEnabled(false); });
  annotationLayer.batchDraw();
  edgeSelections = {}; hideEdgeHandles();
  updatePropsPanel(); updateRenumberHint();
  updateColorSectionVisibility();
}

// ─── Properties Panel ─────────────────────────────────────────────────────────
function updatePropsPanel() {
  const show = id => { const el = document.getElementById(id); if (el) el.style.display = ''; };
  const hide = id => { const el = document.getElementById(id); if (el) el.style.display = 'none'; };
  ['prop-num-row','prop-size-row','prop-style-row','prop-fill-opacity-row','prop-fill-hatch-row',
   'prop-rect-stroke-row','prop-arrow-type-row','prop-arrow-stroke-row',
   'prop-align-row','btn-delete'].forEach(hide);

  const selAnns = [...selectedIds].map(id => findAnn(id)).filter(Boolean);
  const hasNums = selAnns.some(a => a.type === 'number');
  const hasRects = selAnns.some(a => a.type === 'rect');
  const hasArrows = selAnns.some(a => a.type === 'arrow');

  if (selectedIds.size >= 2) show('prop-align-row');
  if (selectedIds.size > 0) show('btn-delete');
  // Disable align buttons that don't apply when edges are selected
  const edgeDir = Object.values(edgeSelections)[0];
  document.querySelectorAll('#prop-align-row .align-btn[data-align]').forEach(btn => {
    const a = btn.dataset.align;
    let enabled = true;
    if (edgeDir) {
      const isV = edgeDir === 'top' || edgeDir === 'bottom';
      if (isV) enabled = a === 'top' || a === 'center-v' || a === 'bottom';
      else enabled = a === 'left' || a === 'center-h' || a === 'right';
    }
    btn.disabled = !enabled;
    btn.style.opacity = enabled ? '' : '0.3';
  });

  if (selectedIds.size === 0) {
    if (activeTool === 'number') {
      show('prop-size-row');
      document.getElementById('prop-size').value = defaultRadius;
      document.getElementById('prop-size-val').textContent = defaultRadius;
      document.getElementById('prop-size-num').value = defaultRadius;
    } else if (activeTool === 'rect') {
      show('prop-fill-opacity-row'); show('prop-fill-hatch-row'); show('prop-rect-stroke-row');
      document.getElementById('prop-rect-stroke').value = defaultRectStroke;
      document.getElementById('prop-rect-stroke-val').textContent = defaultRectStroke;
      document.getElementById('prop-rect-stroke-num').value = defaultRectStroke;
      const fo = Math.round(defaultFillOpacity * 100);
      document.getElementById('prop-fill-opacity').value = fo;
      document.getElementById('prop-fill-opacity-val').textContent = fo;
      updateFillHatchToggle(defaultFillHatch);
    } else if (activeTool === 'arrow') {
      show('prop-arrow-type-row'); show('prop-arrow-stroke-row');
      document.getElementById('prop-arrow-stroke').value = defaultArrowStroke;
      document.getElementById('prop-arrow-stroke-val').textContent = defaultArrowStroke;
      document.getElementById('prop-arrow-stroke-num').value = defaultArrowStroke;
      updateArrowTypeToggle(defaultArrowType);
    }
    updateRectArrowPreview();
    return;
  }

  if (selectedIds.size === 1 && selAnns[0]?.type === 'number') {
    show('prop-num-row');
    document.getElementById('prop-num-input').value = selAnns[0].num;
  }

  if (hasNums) {
    show('prop-size-row');
    show('prop-style-row');
    const firstNum = selAnns.find(a => a.type === 'number');
    const sz = firstNum?.radius || defaultRadius;
    const sty = firstNum?.style || 'circle';
    document.getElementById('prop-size').value = sz;
    document.getElementById('prop-size-val').textContent = sz;
    document.getElementById('prop-size-num').value = sz;
    document.querySelectorAll('#prop-style-selector .style-btn').forEach(b => b.classList.toggle('active', b.dataset.style === sty));
  }
  if (hasRects) {
    show('prop-fill-opacity-row'); show('prop-fill-hatch-row'); show('prop-rect-stroke-row');
    const ra = selAnns.find(a => a.type === 'rect');
    if (ra) {
      const sw = ra.strokeWidth || defaultRectStroke;
      document.getElementById('prop-rect-stroke').value = sw;
      document.getElementById('prop-rect-stroke-val').textContent = sw;
      document.getElementById('prop-rect-stroke-num').value = sw;
      const fo = Math.round((ra.fillOpacity || 0) * 100);
      document.getElementById('prop-fill-opacity').value = fo;
      document.getElementById('prop-fill-opacity-val').textContent = fo;
      updateFillHatchToggle(!!ra.fillHatch);
    }
  }
  if (hasArrows) {
    show('prop-arrow-type-row'); show('prop-arrow-stroke-row');
    const aa = selAnns.find(a => a.type === 'arrow');
    if (aa) {
      const sw = aa.strokeWidth || defaultArrowStroke;
      document.getElementById('prop-arrow-stroke').value = sw;
      document.getElementById('prop-arrow-stroke-val').textContent = sw;
      document.getElementById('prop-arrow-stroke-num').value = sw;
      updateArrowTypeToggle(aa.arrowType || 'arrow');
    }
  }
  if (selectedIds.size === 1) syncPaletteToColor(selAnns[0].color);
  updateRectArrowPreview();
}

function updateFillHatchToggle(hatch) {
  document.getElementById('fill-solid-btn').classList.toggle('active', !hatch);
  document.getElementById('fill-hatch-btn').classList.toggle('active', hatch);
}
function updateArrowTypeToggle(type) {
  document.getElementById('arrow-type-arrow-btn').classList.toggle('active', type === 'arrow');
  document.getElementById('arrow-type-line-btn').classList.toggle('active', type === 'line');
}
function syncPaletteToColor(color) {
  document.querySelectorAll('.color-swatch').forEach((sw, i) => sw.classList.toggle('active', COLORS[i] === color));
}

// ─── History ──────────────────────────────────────────────────────────────────
function deepCopy(arr) { return JSON.parse(JSON.stringify(arr)); }

function saveHistory() {
  if (isRestoringHistory) return;
  history = history.slice(0, historyIndex + 1);
  history.push({ annotations: deepCopy(annotations), selection: [...selectedIds], activeTool });
  historyIndex = history.length - 1;
  updateUndoRedoBtns();
}

// Restore tool UI without side effects (used during undo/redo)
function restoreToolUI(tool) {
  if (!tool || tool === activeTool) return;
  activeTool = tool;
  document.querySelectorAll('.tool-btn').forEach(b => b.classList.toggle('active', b.dataset.tool === tool));
  const cursors = { select:'default', number:'crosshair', rect:'crosshair', arrow:'crosshair' };
  document.getElementById('canvas-area').style.cursor = cursors[tool] || 'default';
  document.getElementById('sel-btns').style.display = tool === 'select' ? 'flex' : 'none';
  document.getElementById('stamp-section').style.display = tool === 'select' ? 'none' : '';
  document.getElementById('num-stamp-content').style.display = tool === 'number' ? '' : 'none';
  document.getElementById('rect-preview-content').style.display = tool === 'rect' ? '' : 'none';
  document.getElementById('arrow-preview-content').style.display = tool === 'arrow' ? '' : 'none';
  document.getElementById('color-type-tabs').style.display = tool === 'select' ? 'none' : 'flex';
  if (tool !== 'select') setActiveColorTab(tool);
  updateRectArrowPreview();
  updateRenumberHint();
}

function undo() {
  if (historyIndex <= 0) return;
  historyIndex--;
  const entry = history[historyIndex];
  annotations = deepCopy(entry.annotations);
  selectedIds = new Set(entry.selection || []);
  isRestoringHistory = true;
  restoreToolUI(entry.activeTool || 'select');
  isRestoringHistory = false;
  edgeSelections = {}; hideEdgeHandles();
  renderAnnotations(); renderNotesPanel(); updateUndoRedoBtns(); updatePropsPanel();
  updateColorSectionVisibility();
}

function redo() {
  if (historyIndex >= history.length - 1) return;
  historyIndex++;
  const entry = history[historyIndex];
  annotations = deepCopy(entry.annotations);
  selectedIds = new Set(entry.selection || []);
  isRestoringHistory = true;
  restoreToolUI(entry.activeTool || 'select');
  isRestoringHistory = false;
  edgeSelections = {}; hideEdgeHandles();
  renderAnnotations(); renderNotesPanel(); updateUndoRedoBtns(); updatePropsPanel();
  updateColorSectionVisibility();
}

function updateUndoRedoBtns() {
  document.getElementById('btn-undo').disabled = historyIndex <= 0;
  document.getElementById('btn-redo').disabled = historyIndex >= history.length - 1;
}

// ─── Marquee ──────────────────────────────────────────────────────────────────
function annotationIntersectsRect(ann, mx, my, mw, mh) {
  let ax1, ay1, ax2, ay2;
  if (ann.type === 'number') {
    const r = ann.radius || 20;
    ax1 = ann.x - r; ay1 = ann.y - r; ax2 = ann.x + r; ay2 = ann.y + r;
  } else if (ann.type === 'rect') {
    ax1 = ann.x; ay1 = ann.y; ax2 = ann.x + ann.width; ay2 = ann.y + ann.height;
  } else if (ann.type === 'arrow') {
    ax1 = Math.min(ann.points[0],ann.points[2])-5; ay1 = Math.min(ann.points[1],ann.points[3])-5;
    ax2 = Math.max(ann.points[0],ann.points[2])+5; ay2 = Math.max(ann.points[1],ann.points[3])+5;
  } else return false;
  const rx1 = mw>=0?mx:mx+mw, ry1 = mh>=0?my:my+mh;
  const rx2 = mw>=0?mx+mw:mx, ry2 = mh>=0?my+mh:my;
  return ax1<rx2&&ax2>rx1&&ay1<ry2&&ay2>ry1;
}

// ─── Stage Events ─────────────────────────────────────────────────────────────
function onStageDown(e) {
  if (!imageLoaded) return;
  if (e.evt && e.evt.button === 1) return;
  const sp = getStagePointer(), ip = getPointer();

  if (activeTool === 'select') {
    const isEmptyClick = e.target === stage || e.target.getLayer() === imageLayer;
    if (isEmptyClick) {
      marqueeActive = true; marqueeStart = { ...sp };
      marqueeRect = new Konva.Rect({ x:sp.x,y:sp.y,width:0,height:0,stroke:'#ef4444',strokeWidth:1,dash:[4,3],fill:'#ef444411' });
      drawLayer.add(marqueeRect); drawLayer.batchDraw();
    }
    return;
  }

  if (activeTool === 'number') { placeNumber(ip.x, ip.y); return; }

  if (activeTool === 'rect' || activeTool === 'arrow') {
    isDrawing = true; drawStart = { ...sp };
    if (activeTool === 'rect') {
      tempShape = new Konva.Rect({ x:sp.x,y:sp.y,width:0,height:0,stroke:colorByType.rect,strokeWidth:defaultRectStroke,fill:'transparent',dash:[4,3] });
    } else {
      tempShape = new Konva.Arrow({ points:[sp.x,sp.y,sp.x,sp.y],stroke:colorByType.arrow,strokeWidth:defaultArrowStroke,fill:colorByType.arrow,pointerLength:defaultArrowType==='line'?0:Math.max(10,defaultArrowStroke*4),pointerWidth:defaultArrowType==='line'?0:Math.max(8,defaultArrowStroke*3) });
    }
    drawLayer.add(tempShape); drawLayer.batchDraw();
  }
}

function onStageMove(e) {
  const sp = getStagePointer();
  lastStagePointer = sp;

  // Cursor ghost for number tool
  updateCursorGhost(sp);

  if (marqueeActive && marqueeRect) {
    marqueeRect.x(Math.min(sp.x,marqueeStart.x)); marqueeRect.y(Math.min(sp.y,marqueeStart.y));
    marqueeRect.width(Math.abs(sp.x-marqueeStart.x)); marqueeRect.height(Math.abs(sp.y-marqueeStart.y));
    drawLayer.batchDraw(); return;
  }
  if (!isDrawing || !tempShape) return;
  if (activeTool === 'rect') {
    tempShape.x(Math.min(sp.x,drawStart.x)); tempShape.y(Math.min(sp.y,drawStart.y));
    tempShape.width(Math.abs(sp.x-drawStart.x)); tempShape.height(Math.abs(sp.y-drawStart.y));
  } else if (activeTool === 'arrow') {
    tempShape.points([drawStart.x,drawStart.y,sp.x,sp.y]);
  }
  drawLayer.batchDraw();
}

function onStageUp(e) {
  const sp = getStagePointer(), ip = getPointer();
  if (marqueeActive) {
    marqueeActive = false;
    if (marqueeRect) {
      const mx = marqueeRect.x()-STAGE_PAD, my = marqueeRect.y()-STAGE_PAD;
      const mw = marqueeRect.width(), mh = marqueeRect.height();
      marqueeRect.destroy(); marqueeRect = null; drawLayer.batchDraw();
      if (mw > 2 || mh > 2) {
        const intersecting = annotations.filter(ann => annotationIntersectsRect(ann, mx, my, mw, mh));
        if (e.evt && e.evt.shiftKey) intersecting.forEach(ann => selectedIds.add(ann.id));
        else selectedIds = new Set(intersecting.map(a => a.id));
        updateSelectionVisual(); updatePropsPanel(); saveHistory();
      } else {
        deselectAll();
      }
    }
    return;
  }
  if (!isDrawing || !tempShape) return;
  isDrawing = false;
  if (activeTool === 'rect') {
    const w = Math.abs(sp.x-drawStart.x), h = Math.abs(sp.y-drawStart.y);
    if (w >= 4 && h >= 4) {
      const ann = { id:'ann_'+(nextId++), type:'rect',
        x: Math.min(sp.x,drawStart.x)-STAGE_PAD, y: Math.min(sp.y,drawStart.y)-STAGE_PAD,
        width:w, height:h, color:colorByType.rect, strokeWidth:defaultRectStroke,
        fillOpacity:defaultFillOpacity, fillHatch:defaultFillHatch };
      annotations.push(ann); saveHistory(); renderAnnotations(); renderNotesPanel();
    }
  } else if (activeTool === 'arrow') {
    const dx = sp.x-drawStart.x, dy = sp.y-drawStart.y;
    if (Math.sqrt(dx*dx+dy*dy) >= 4) {
      const ann = { id:'ann_'+(nextId++), type:'arrow',
        points:[drawStart.x-STAGE_PAD,drawStart.y-STAGE_PAD,sp.x-STAGE_PAD,sp.y-STAGE_PAD],
        color:colorByType.arrow, strokeWidth:defaultArrowStroke, arrowType:defaultArrowType };
      annotations.push(ann); saveHistory(); renderAnnotations(); renderNotesPanel();
    }
  }
  tempShape.destroy(); tempShape = null; drawLayer.batchDraw();
}

// ─── Place Number ─────────────────────────────────────────────────────────────
function placeNumber(x, y) {
  const ann = { id:'ann_'+(nextId++), type:'number', x, y, num:nextNum, color:colorByType.number, radius:defaultRadius, style:defaultNumStyle, label:'' };
  annotations.push(ann);
  nextNum++;
  updateStampBadge();
  saveHistory(); renderAnnotations(); renderNotesPanel();
}

function updateStampBadge() {
  const badge = document.getElementById('stamp-badge');
  // Linear scale: radius 6→22px, radius 100→78px — always changes
  const displaySize = Math.round(22 + (defaultRadius - 6) / (100 - 6) * 56);
  badge.style.width = displaySize + 'px'; badge.style.height = displaySize + 'px';
  badge.style.fontSize = Math.round(displaySize * 0.45) + 'px';
  badge.textContent = nextNum;
  badge.style.background = colorByType.number;
  badge.style.color = textColor(colorByType.number);
  document.getElementById('next-num-input').value = nextNum;
  document.getElementById('stamp-ghost').textContent = nextNum;
  document.getElementById('stamp-ghost').style.background = colorByType.number;
  badge.style.clipPath = '';
  const styleRadiusMap = { circle:'50%', square:'0', 'circle-outline':'50%', 'rounded-square':'30%', diamond:'0' };
  badge.style.borderRadius = styleRadiusMap[defaultNumStyle] || '50%';
  if (defaultNumStyle === 'diamond') {
    badge.style.clipPath = 'polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)';
    badge.style.border = '3px solid #fff2';
  } else if (defaultNumStyle === 'circle-outline') {
    badge.style.background = 'transparent';
    badge.style.border = '3px solid ' + colorByType.number;
    badge.style.color = colorByType.number;
  } else {
    badge.style.border = '3px solid #fff2';
  }
  invalidateCursorGhost();
}

function updateStyleSelector() {
  document.querySelectorAll('#style-selector .style-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.style === defaultNumStyle));
}

// ─── Rect/Arrow Preview ───────────────────────────────────────────────────────
function updateRectArrowPreview() {
  const rectSvg = document.getElementById('rect-preview-svg');
  if (rectSvg && activeTool === 'rect') {
    const color = colorByType.rect;
    const sw = defaultRectStroke;
    // Scale stroke for preview: maps 1..20 → 1..11, always changes
    const visualSw = Math.max(1, Math.round(1 + (sw - 1) / 19 * 10));
    const pad = visualSw / 2 + 2;
    const rx = 80 - pad * 2, ry = 60 - pad * 2;
    let inner;
    if (defaultFillHatch && defaultFillOpacity > 0) {
      inner = `<defs><pattern id="hp" width="10" height="10" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><line x1="0" y1="0" x2="0" y2="10" stroke="${color}" stroke-width="1.5" opacity="${defaultFillOpacity}"/></pattern></defs><rect x="${pad}" y="${pad}" width="${rx}" height="${ry}" stroke="${color}" stroke-width="${visualSw}" fill="url(#hp)"/>`;
    } else {
      const fillC = defaultFillOpacity > 0 ? hexToRgba(color, defaultFillOpacity) : 'transparent';
      inner = `<rect x="${pad}" y="${pad}" width="${rx}" height="${ry}" stroke="${color}" stroke-width="${visualSw}" fill="${fillC}"/>`;
    }
    rectSvg.innerHTML = inner;
  }

  const arrowSvg = document.getElementById('arrow-preview-svg');
  if (arrowSvg && activeTool === 'arrow') {
    const color = colorByType.arrow;
    const sw = defaultArrowStroke;
    const visualSw = Math.max(1, Math.round(1 + (sw - 1) / 19 * 10));
    if (defaultArrowType === 'line') {
      arrowSvg.innerHTML = `<line x1="10" y1="50" x2="70" y2="10" stroke="${color}" stroke-width="${visualSw}" stroke-linecap="round"/>`;
    } else {
      // Draw arrowhead as polygon (no SVG markers) to avoid stroke-overlap bug
      const x1=10, y1=50, tipX=70, tipY=10;
      const dx=tipX-x1, dy=tipY-y1, len=Math.sqrt(dx*dx+dy*dy);
      const ux=dx/len, uy=dy/len;
      const mw=Math.max(10, visualSw*4), mh=Math.max(8, visualSw*3);
      const shaftEndX=tipX-mw*ux, shaftEndY=tipY-mw*uy;
      const px=-uy, py=ux;
      const b1x=shaftEndX+(mh/2)*px, b1y=shaftEndY+(mh/2)*py;
      const b2x=shaftEndX-(mh/2)*px, b2y=shaftEndY-(mh/2)*py;
      const f = n => n.toFixed(1);
      arrowSvg.innerHTML = `<line x1="${x1}" y1="${y1}" x2="${f(shaftEndX)}" y2="${f(shaftEndY)}" stroke="${color}" stroke-width="${visualSw}" stroke-linecap="round"/><polygon points="${tipX},${tipY} ${f(b1x)},${f(b1y)} ${f(b2x)},${f(b2y)}" fill="${color}"/>`;
    }
  }
}

// ─── Cursor Ghost (Number Tool) ───────────────────────────────────────────────
function invalidateCursorGhost() {
  cursorGhostNeedsRebuild = true;
  if (lastStagePointer) updateCursorGhost(lastStagePointer);
}

function updateCursorGhost(sp) {
  if (activeTool !== 'number' || !imageLoaded) {
    if (cursorGhost) { cursorGhost.destroy(); cursorGhost = null; if (drawLayer) drawLayer.batchDraw(); }
    return;
  }

  if (cursorGhostNeedsRebuild || !cursorGhost) {
    if (cursorGhost) { cursorGhost.destroy(); cursorGhost = null; }
    const r = defaultRadius, style = defaultNumStyle, color = colorByType.number;
    const numStr = String(nextNum);
    const fsize = r * (numStr.length > 2 ? 0.7 : numStr.length > 1 ? 0.85 : 1.1);
    const isOutline = style === 'circle-outline';
    const tColor = isOutline ? color : textColor(color);
    let shape;
    if (style === 'circle') {
      shape = new Konva.Circle({ x:0,y:0,radius:r,fill:color });
    } else if (style === 'square') {
      shape = new Konva.Rect({ x:-r,y:-r,width:r*2,height:r*2,fill:color });
    } else if (style === 'circle-outline') {
      shape = new Konva.Circle({ x:0,y:0,radius:r,fill:'transparent',stroke:color,strokeWidth:2.5 });
    } else if (style === 'rounded-square') {
      shape = new Konva.Rect({ x:-r,y:-r,width:r*2,height:r*2,cornerRadius:r*0.4,fill:color });
    } else if (style === 'diamond') {
      shape = new Konva.Line({ points:[0,-r,r,0,0,r,-r,0],closed:true,fill:color });
    } else {
      shape = new Konva.Circle({ x:0,y:0,radius:r,fill:color });
    }
    const text = new Konva.Text({ x:-r,y:-r,width:r*2,height:r*2,text:numStr,fontSize:fsize,fontStyle:'bold',fill:tColor,align:'center',verticalAlign:'middle' });
    cursorGhost = new Konva.Group({ x: sp.x, y: sp.y, opacity: 0.4, listening: false });
    cursorGhost.add(shape, text);
    drawLayer.add(cursorGhost);
    cursorGhostNeedsRebuild = false;
  } else {
    cursorGhost.position({ x: sp.x, y: sp.y });
  }
  drawLayer.batchDraw();
}

// ─── Tool Selection ───────────────────────────────────────────────────────────
function setTool(tool) {
  if (tool === activeTool) return;
  if (activeTool === 'select' && tool !== 'select') deselectAll();
  if (tool !== 'select') { edgeSelections = {}; hideEdgeHandles(); }
  if (tool !== 'number' && cursorGhost) { cursorGhost.destroy(); cursorGhost = null; if (drawLayer) drawLayer.batchDraw(); }
  if (tool === 'number') invalidateCursorGhost();
  activeTool = tool;
  document.querySelectorAll('.tool-btn').forEach(b => b.classList.toggle('active', b.dataset.tool === tool));
  const cursors = { select:'default', number:'crosshair', rect:'crosshair', arrow:'crosshair' };
  document.getElementById('canvas-area').style.cursor = cursors[tool] || 'default';
  document.getElementById('sel-btns').style.display = tool === 'select' ? 'flex' : 'none';
  document.getElementById('stamp-section').style.display = tool === 'select' ? 'none' : '';
  document.getElementById('num-stamp-content').style.display = tool === 'number' ? '' : 'none';
  document.getElementById('rect-preview-content').style.display = tool === 'rect' ? '' : 'none';
  document.getElementById('arrow-preview-content').style.display = tool === 'arrow' ? '' : 'none';
  document.getElementById('color-type-tabs').style.display = tool === 'select' ? 'none' : 'flex';
  if (tool !== 'select') setActiveColorTab(tool);
  updatePropsPanel();
  updateRectArrowPreview();
  updateRenumberHint();
  updateColorSectionVisibility();
  saveHistory();
}

// ─── Color Section Visibility ─────────────────────────────────────────────────
function updateColorSectionVisibility() {
  const visible = activeTool !== 'select' || selectedIds.size > 0;
  const el = document.getElementById('color-section');
  if (el) el.style.display = visible ? '' : 'none';
}

// ─── Color Palette ────────────────────────────────────────────────────────────
function buildPalette() {
  const palette = document.getElementById('color-palette');
  COLORS.forEach(c => {
    const sw = document.createElement('div');
    sw.className = 'color-swatch' + (c === colorByType.number ? ' active' : '');
    sw.style.background = c;
    sw.style.border = c === '#1f2937' ? '2px solid #374151' : '2px solid transparent';
    sw.title = c;
    sw.addEventListener('click', () => setColor(c));
    palette.appendChild(sw);
  });
}

function setColor(color) {
  colorByType[activeColorType] = color;
  if (activeColorType === 'number') invalidateCursorGhost();
  updateColorDots();
  document.querySelectorAll('.color-swatch').forEach((sw,i) => sw.classList.toggle('active', COLORS[i] === color));
  document.getElementById('custom-color').value = /^#[0-9a-fA-F]{6}$/.test(color) ? color : '#ef4444';
  updateStampBadge();
  updateRectArrowPreview();
  if (selectedIds.size > 0) {
    const types = new Set([...selectedIds].map(id => findAnn(id)?.type).filter(Boolean));
    const applyAll = types.size > 1;
    let changed = false;
    selectedIds.forEach(id => {
      const ann = findAnn(id);
      if (ann && (applyAll || ann.type === activeColorType)) { ann.color = color; changed = true; }
    });
    if (changed) { saveHistory(); renderAnnotations(); updatePropsPanel(); }
  }
}

// ─── Alignment ────────────────────────────────────────────────────────────────
function getAnnBounds(ann) {
  if (ann.type === 'number') { const r=ann.radius||20; return {x:ann.x-r,y:ann.y-r,w:r*2,h:r*2}; }
  if (ann.type === 'rect') return {x:ann.x,y:ann.y,w:ann.width,h:ann.height};
  if (ann.type === 'arrow') {
    const x=Math.min(ann.points[0],ann.points[2]), y=Math.min(ann.points[1],ann.points[3]);
    return {x,y,w:Math.abs(ann.points[2]-ann.points[0]),h:Math.abs(ann.points[3]-ann.points[1])};
  }
  return {x:0,y:0,w:0,h:0};
}

function moveAnn(ann, newX, newY) {
  if (ann.type === 'number') { const r=ann.radius||20; ann.x=newX+r; ann.y=newY+r; }
  else if (ann.type === 'rect') { ann.x=newX; ann.y=newY; }
  else if (ann.type === 'arrow') { const b=getAnnBounds(ann),dx=newX-b.x,dy=newY-b.y; ann.points=[ann.points[0]+dx,ann.points[1]+dy,ann.points[2]+dx,ann.points[3]+dy]; }
}

function alignSelected(dir) {
  const edgeEntries = Object.entries(edgeSelections);
  if (edgeEntries.length >= 2) {
    const edgeDir = edgeEntries[0][1];
    const isV = edgeDir === 'top' || edgeDir === 'bottom';
    const posOf = (ann, d) => {
      if (d === 'top') return ann.y;
      if (d === 'bottom') return ann.y + ann.height;
      if (d === 'left') return ann.x;
      if (d === 'right') return ann.x + ann.width;
    };
    const positions = edgeEntries.map(([id, d]) => { const a = findAnn(id); return a ? posOf(a, d) : null; }).filter(v => v != null);
    if (positions.length < 2) return;
    let target;
    if ((isV && dir === 'top') || (!isV && dir === 'left')) target = Math.min(...positions);
    else if ((isV && dir === 'bottom') || (!isV && dir === 'right')) target = Math.max(...positions);
    else if ((isV && dir === 'center-v') || (!isV && dir === 'center-h')) target = (Math.min(...positions) + Math.max(...positions)) / 2;
    else return;
    const MIN = 4;
    edgeEntries.forEach(([id, d]) => {
      const a = findAnn(id); if (!a) return;
      if (d === 'top') {
        const bottom = a.y + a.height;
        a.y = target; a.height = Math.max(MIN, bottom - target);
      } else if (d === 'bottom') {
        a.height = Math.max(MIN, target - a.y);
      } else if (d === 'left') {
        const right = a.x + a.width;
        a.x = target; a.width = Math.max(MIN, right - target);
      } else if (d === 'right') {
        a.width = Math.max(MIN, target - a.x);
      }
    });
    edgeSelections = {};
    saveHistory(); renderAnnotations(); renderNotesPanel();
    showEdgeHandles(); updatePropsPanel();
    return;
  }
  const sel = [...selectedIds].map(id=>findAnn(id)).filter(Boolean);
  if (sel.length < 2) return;
  const bounds = sel.map(a => ({a,b:getAnnBounds(a)}));
  if (dir==='left') { const minX=Math.min(...bounds.map(({b})=>b.x)); bounds.forEach(({a,b})=>moveAnn(a,minX,b.y)); }
  else if (dir==='center-h') { const minX=Math.min(...bounds.map(({b})=>b.x)),maxX=Math.max(...bounds.map(({b})=>b.x+b.w)),cx=(minX+maxX)/2; bounds.forEach(({a,b})=>moveAnn(a,cx-b.w/2,b.y)); }
  else if (dir==='right') { const maxX=Math.max(...bounds.map(({b})=>b.x+b.w)); bounds.forEach(({a,b})=>moveAnn(a,maxX-b.w,b.y)); }
  else if (dir==='top') { const minY=Math.min(...bounds.map(({b})=>b.y)); bounds.forEach(({a,b})=>moveAnn(a,b.x,minY)); }
  else if (dir==='center-v') { const minY=Math.min(...bounds.map(({b})=>b.y)),maxY=Math.max(...bounds.map(({b})=>b.y+b.h)),cy=(minY+maxY)/2; bounds.forEach(({a,b})=>moveAnn(a,b.x,cy-b.h/2)); }
  else if (dir==='bottom') { const maxY=Math.max(...bounds.map(({b})=>b.y+b.h)); bounds.forEach(({a,b})=>moveAnn(a,b.x,maxY-b.h)); }
  else if (dir==='distribute-h' && sel.length >= 3) {
    const sorted = bounds.slice().sort((a,b)=>a.b.x-b.b.x);
    const first=sorted[0].b, last=sorted[sorted.length-1].b;
    const innerW = sorted.slice(1,-1).reduce((s,{b})=>s+b.w,0);
    const gap = (last.x - first.x - first.w - innerW) / (sorted.length-1);
    let curX = first.x + first.w + gap;
    for (let i=1;i<sorted.length-1;i++) { moveAnn(sorted[i].a,curX,sorted[i].b.y); curX+=sorted[i].b.w+gap; }
  }
  else if (dir==='distribute-v' && sel.length >= 3) {
    const sorted = bounds.slice().sort((a,b)=>a.b.y-b.b.y);
    const first=sorted[0].b, last=sorted[sorted.length-1].b;
    const innerH = sorted.slice(1,-1).reduce((s,{b})=>s+b.h,0);
    const gap = (last.y - first.y - first.h - innerH) / (sorted.length-1);
    let curY = first.y + first.h + gap;
    for (let i=1;i<sorted.length-1;i++) { moveAnn(sorted[i].a,sorted[i].b.x,curY); curY+=sorted[i].b.h+gap; }
  }
  saveHistory(); renderAnnotations();
}

// ─── Delete ───────────────────────────────────────────────────────────────────
function deleteSelected() {
  if (selectedIds.size === 0) return;
  annotations = annotations.filter(a => !selectedIds.has(a.id));
  selectedIds = new Set();
  saveHistory(); renderAnnotations(); renderNotesPanel(); updatePropsPanel();
}

// ─── Copy / Paste ─────────────────────────────────────────────────────────────
function copySelected() {
  if (selectedIds.size === 0) return;
  clipboard = [...selectedIds].map(id => findAnn(id)).filter(Boolean).map(a => JSON.parse(JSON.stringify(a)));
}

function pasteAnnotations() {
  if (clipboard.length === 0 || !imageLoaded) return;
  const OFFSET = 15;
  const newAnns = clipboard.map(ann => {
    const newAnn = JSON.parse(JSON.stringify(ann));
    newAnn.id = 'ann_' + (nextId++);
    if (newAnn.type === 'number' || newAnn.type === 'rect') {
      newAnn.x += OFFSET; newAnn.y += OFFSET;
    } else if (newAnn.type === 'arrow') {
      newAnn.points = [newAnn.points[0]+OFFSET, newAnn.points[1]+OFFSET, newAnn.points[2]+OFFSET, newAnn.points[3]+OFFSET];
    }
    return newAnn;
  });
  annotations.push(...newAnns);
  selectedIds = new Set(newAnns.map(a => a.id));
  saveHistory(); renderAnnotations(); renderNotesPanel(); updatePropsPanel();
}

// ─── Renumber ─────────────────────────────────────────────────────────────────
function updateRenumberHint() {
  const hint = document.getElementById('renumber-scope-hint');
  if (!hint) return;
  const selNums = activeTool === 'select' ? annotations.filter(a => a.type==='number' && selectedIds.has(a.id)) : [];
  hint.textContent = selNums.length > 0 ? '(選択範囲内)' : '';
}

function renumber() {
  const mode = document.getElementById('renumber-mode').value;
  const selNums = activeTool === 'select' ? annotations.filter(a => a.type==='number' && selectedIds.has(a.id)) : [];
  const scope = selNums.length > 0 ? selNums : annotations.filter(a => a.type==='number');
  if (scope.length === 0) return;
  const allNums = annotations.filter(a => a.type==='number');
  if (mode === 'max-down') {
    const maxNum = allNums.length ? Math.max(...allNums.map(a=>parseInt(a.num)||1)) : scope.length;
    if (maxNum < scope.length) { alert(`最大値(${maxNum})が対象アノテーション数(${scope.length})より小さいため実行できません。`); return; }
    // Sort descending so the largest original number maps to maxNum (preserves the top)
    scope.sort((a,b)=>(parseInt(b.num)||0)-(parseInt(a.num)||0));
    scope.forEach((ann,i) => { ann.num = maxNum - i; });
  } else {
    let startNum;
    if (mode === 'from1') startNum = 1;
    else if (mode === 'min') startNum = allNums.length ? Math.min(...allNums.map(a=>parseInt(a.num)||1)) : 1;
    scope.sort((a,b)=>(parseInt(a.num)||0)-(parseInt(b.num)||0));
    scope.forEach((ann,i) => { ann.num = startNum + i; });
  }
  const updated = annotations.filter(a => a.type==='number');
  nextNum = updated.length ? Math.max(...updated.map(a=>parseInt(a.num)||1)) + 1 : 1;
  updateStampBadge();
  saveHistory(); renderAnnotations(); renderNotesPanel(); updatePropsPanel();
}

// ─── Notes Panel ─────────────────────────────────────────────────────────────
function getSortedNotes() {
  const sortVal = document.getElementById('notes-sort').value;
  let nums = annotations.filter(a => a.type==='number');
  const styleOrder = ['circle','square','circle-outline','rounded-square','diamond'];
  if (sortVal==='asc') nums.sort((a,b)=>(parseInt(a.num)||0)-(parseInt(b.num)||0));
  else if (sortVal==='desc') nums.sort((a,b)=>(parseInt(b.num)||0)-(parseInt(a.num)||0));
  else if (sortVal==='style-asc') nums.sort((a,b)=>{ const sd=styleOrder.indexOf(a.style||'circle')-styleOrder.indexOf(b.style||'circle'); return sd!==0?sd:(parseInt(a.num)||0)-(parseInt(b.num)||0); });
  else if (sortVal==='style-desc') nums.sort((a,b)=>{ const sd=styleOrder.indexOf(a.style||'circle')-styleOrder.indexOf(b.style||'circle'); return sd!==0?sd:(parseInt(b.num)||0)-(parseInt(a.num)||0); });
  return nums;
}

function renderNotesPanel() {
  const body = document.getElementById('notes-body');
  const empty = document.getElementById('notes-empty');
  body.querySelectorAll('.note-row').forEach(r => r.remove());
  const nums = getSortedNotes();
  if (nums.length===0) { empty.style.display=''; return; }
  empty.style.display='none';
  nums.forEach(ann => {
    const row = document.createElement('div'); row.className='note-row';
    const badge = document.createElement('div'); badge.className='note-badge';
    badge.textContent = ann.num; badge.style.background = ann.color; badge.style.color = textColor(ann.color);
    if (ann.style==='circle-outline') { badge.style.background='#ffffff'; badge.style.border='2px solid '+ann.color; badge.style.color=ann.color; }
    if (ann.style==='diamond') {
      badge.style.clipPath='polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)';
      badge.style.borderRadius='0';
    } else {
      badge.style.borderRadius = ann.style==='square'?'0':ann.style==='rounded-square'?'30%':'50%';
    }
    const ta = document.createElement('textarea'); ta.className='note-input';
    ta.dataset.id = ann.id; ta.value = ann.label||''; ta.placeholder='メモを入力...'; ta.rows=1;
    ta.addEventListener('input', () => {
      ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px';
      const a=findAnn(ta.dataset.id); if(a) a.label=ta.value;
    });
    setTimeout(() => { ta.style.height='auto'; ta.style.height=ta.scrollHeight+'px'; }, 0);
    row.appendChild(badge); row.appendChild(ta); body.appendChild(row);
  });
}

// ─── Export ───────────────────────────────────────────────────────────────────
function computeAnnotationBounds() {
  let minX=0,minY=0,maxX=baseStageW,maxY=baseStageH,hasOut=false;
  annotations.forEach(ann => {
    let ax1,ay1,ax2,ay2;
    if (ann.type==='number') { const r=ann.radius||20; ax1=ann.x-r;ay1=ann.y-r;ax2=ann.x+r;ay2=ann.y+r; }
    else if (ann.type==='rect') { ax1=ann.x;ay1=ann.y;ax2=ann.x+(ann.width||0);ay2=ann.y+(ann.height||0); }
    else if (ann.type==='arrow') { const p=ann.points||[]; ax1=Math.min(p[0]||0,p[2]||0);ay1=Math.min(p[1]||0,p[3]||0);ax2=Math.max(p[0]||0,p[2]||0);ay2=Math.max(p[1]||0,p[3]||0); }
    else return;
    if (ax1<0||ay1<0||ax2>baseStageW||ay2>baseStageH) hasOut=true;
    minX=Math.min(minX,ax1);minY=Math.min(minY,ay1);maxX=Math.max(maxX,ax2);maxY=Math.max(maxY,ay2);
  });
  return {minX,minY,maxX,maxY,hasOut};
}

function doExport(bgChoice) {
  deselectAll(); renderAnnotations();
  const bounds = computeAnnotationBounds();
  const PR = 2;

  function cropAndDownload(fullDataURL, srcX, srcY, srcW, srcH, dstW, dstH, bg) {
    const img = new Image();
    img.onload = () => {
      const oc = document.createElement('canvas');
      oc.width = dstW * PR; oc.height = dstH * PR;
      const ctx = oc.getContext('2d');
      if (bg === 'white') { ctx.fillStyle='#ffffff'; ctx.fillRect(0,0,oc.width,oc.height); }
      else if (bg === 'custom') { ctx.fillStyle=document.getElementById('export-custom-color').value; ctx.fillRect(0,0,oc.width,oc.height); }
      else if (bg && bg !== 'transparent') { ctx.fillStyle=bg; ctx.fillRect(0,0,oc.width,oc.height); }
      ctx.drawImage(img, srcX*PR, srcY*PR, srcW*PR, srcH*PR, 0, 0, dstW*PR, dstH*PR);
      triggerDownload(oc.toDataURL('image/png'));
    };
    img.src = fullDataURL;
  }

  setTimeout(() => {
    const stageDataURL = stage.toDataURL({ pixelRatio: PR, mimeType: 'image/png' });
    if (!bounds.hasOut) {
      cropAndDownload(stageDataURL,
        STAGE_PAD * viewScale, STAGE_PAD * viewScale,
        baseStageW * viewScale, baseStageH * viewScale,
        baseStageW, baseStageH, 'transparent');
    } else {
      const padL=Math.max(0,-bounds.minX), padT=Math.max(0,-bounds.minY);
      const padR=Math.max(0,bounds.maxX-baseStageW), padB=Math.max(0,bounds.maxY-baseStageH);
      const totalW=baseStageW+padL+padR, totalH=baseStageH+padT+padB;
      cropAndDownload(stageDataURL,
        (STAGE_PAD - padL) * viewScale, (STAGE_PAD - padT) * viewScale,
        totalW * viewScale, totalH * viewScale,
        totalW, totalH, bgChoice);
    }
  }, 50);
}

// ─── JSON Save / Load ──────────────────────────────────────────────────────────
function saveJSON() {
  if (!imageLoaded) { alert('画像を読み込んでください'); return; }
  const imageDataURL = imageLayer.getChildren(n => n.className === 'Image')[0];
  const imgEl = imageDataURL ? imageDataURL.image() : null;
  const canvas = document.createElement('canvas');
  canvas.width = imgEl ? imgEl.width : 0; canvas.height = imgEl ? imgEl.height : 0;
  if (imgEl) canvas.getContext('2d').drawImage(imgEl, 0, 0);
  const data = {
    version: 1,
    imageDataURL: canvas.toDataURL('image/png'),
    baseStageW, baseStageH,
    annotations: JSON.parse(JSON.stringify(annotations)),
    nextId, nextNum,
    colorByType: { ...colorByType },
    defaultRadius, defaultRectStroke, defaultArrowStroke,
    defaultArrowType, defaultFillOpacity, defaultFillOpacitySolid, defaultFillOpacityHatch, defaultFillHatch, defaultNumStyle, bgColor
  };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = 'numarker_' + Date.now() + '.json'; a.click();
}

function loadJSON(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      if (!data.version || !data.annotations) { alert('無効なJSONファイルです'); return; }
      if (annotations.length > 0 || imageLoaded) {
        if (!confirm('現在の編集内容が失われます。続けますか？')) return;
      }
      annotations = []; selectedIds = new Set(); history = []; historyIndex = -1;
      nextId = data.nextId || 1; nextNum = data.nextNum || 1;
      if (data.colorByType) { colorByType.number = data.colorByType.number || '#ef4444'; colorByType.rect = data.colorByType.rect || '#3b82f6'; colorByType.arrow = data.colorByType.arrow || '#f97316'; }
      defaultRadius = data.defaultRadius || 20; defaultRectStroke = data.defaultRectStroke || 3;
      defaultArrowStroke = data.defaultArrowStroke || 3; defaultArrowType = data.defaultArrowType || 'arrow';
      defaultFillOpacitySolid = data.defaultFillOpacitySolid != null ? data.defaultFillOpacitySolid : 0;
      defaultFillOpacityHatch = data.defaultFillOpacityHatch != null ? data.defaultFillOpacityHatch : 0.25;
      defaultFillHatch = !!data.defaultFillHatch;
      defaultFillOpacity = data.defaultFillOpacity != null ? data.defaultFillOpacity : (defaultFillHatch ? defaultFillOpacityHatch : defaultFillOpacitySolid);
      defaultNumStyle = data.defaultNumStyle || 'circle';
      bgColor = data.bgColor || 'transparent';
      const bgSel = document.getElementById('bg-color-select');
      if (bgColor === 'transparent' || bgColor === '#ffffff') { bgSel.value = bgColor; document.getElementById('bg-custom-color').style.display = 'none'; }
      else { bgSel.value = 'custom'; document.getElementById('bg-custom-color').value = bgColor; document.getElementById('bg-custom-color').style.display = ''; }
      const img = new Image();
      img.onload = () => {
        baseStageW = data.baseStageW || img.width; baseStageH = data.baseStageH || img.height;
        imageLayer.destroyChildren(); bgRect = null;
        const kImg = new Konva.Image({ image: img, x: STAGE_PAD, y: STAGE_PAD, width: baseStageW, height: baseStageH });
        imageLayer.add(kImg); imageLoaded = true;
        document.getElementById('drop-hint').classList.add('hidden');
        setZoom(viewScale);
        const area = document.getElementById('canvas-area');
        const PAD_PX = 500;
        document.getElementById('canvas-pan-pad').style.padding = PAD_PX + 'px';
        const totalW = (baseStageW + 2*STAGE_PAD)*viewScale, totalH = (baseStageH + 2*STAGE_PAD)*viewScale;
        area.scrollLeft = (totalW + 2*PAD_PX - area.clientWidth) / 2;
        area.scrollTop = (totalH + 2*PAD_PX - area.clientHeight) / 2;
        annotations = data.annotations;
        invalidateCursorGhost();
        updateColorDots(); updateStampBadge(); updateStyleSelector();
        saveHistory(); renderAnnotations(); renderNotesPanel(); updatePropsPanel();
      };
      img.src = data.imageDataURL;
    } catch(err) { alert('JSONの読み込みに失敗しました: ' + err.message); }
  };
  reader.readAsText(file);
}

function triggerDownload(dataURL) {
  const a=document.createElement('a'); a.href=dataURL; a.download='numarker_'+Date.now()+'.png'; a.click();
}

function exportPNG() {
  if (!imageLoaded||!stage) return;
  const bounds = computeAnnotationBounds();
  if (bounds.hasOut) {
    if (bgColor === 'transparent') {
      document.querySelector('input[name="export-bg"][value="transparent"]').checked = true;
    } else if (bgColor === '#ffffff') {
      document.querySelector('input[name="export-bg"][value="white"]').checked = true;
    } else {
      document.querySelector('input[name="export-bg"][value="custom"]').checked = true;
      document.getElementById('export-custom-color').value = bgColor;
    }
    document.getElementById('export-panel').classList.remove('hidden');
  } else {
    doExport(bgColor === 'transparent' ? 'transparent' : bgColor === '#ffffff' ? 'white' : bgColor);
  }
}

// ─── Background Color ─────────────────────────────────────────────────────────
function updateBgRect() {
  if (bgRect) { bgRect.destroy(); bgRect = null; }
  if (!imageLoaded || bgColor === 'transparent') { if(annotationLayer) annotationLayer.batchDraw(); return; }
  const bounds = computeAnnotationBounds();
  const x1 = Math.min(0, bounds.minX), y1 = Math.min(0, bounds.minY);
  const x2 = Math.max(baseStageW, bounds.maxX), y2 = Math.max(baseStageH, bounds.maxY);
  bgRect = new Konva.Rect({
    x: x1 + STAGE_PAD, y: y1 + STAGE_PAD,
    width: x2 - x1, height: y2 - y1,
    fill: bgColor, listening: false
  });
  imageLayer.add(bgRect);
  bgRect.moveToBottom();
  imageLayer.batchDraw();
}

document.getElementById('bg-color-select').addEventListener('change', e => {
  const val = e.target.value;
  document.getElementById('bg-custom-color').style.display = val === 'custom' ? '' : 'none';
  bgColor = val === 'custom' ? document.getElementById('bg-custom-color').value : val;
  updateBgRect();
});
document.getElementById('bg-custom-color').addEventListener('input', e => {
  bgColor = e.target.value;
  updateBgRect();
});

// ─── Stamp Badge Drag ─────────────────────────────────────────────────────────
const stampBadge = document.getElementById('stamp-badge');
const stampGhost = document.getElementById('stamp-ghost');

stampBadge.addEventListener('dragstart', e => {
  e.dataTransfer.setData('text/plain','stamp'); e.dataTransfer.effectAllowed='copy';
  const blank=document.createElement('canvas'); blank.width=blank.height=1; e.dataTransfer.setDragImage(blank,0,0);
  stampGhost.style.display='flex'; stampGhost.style.background=colorByType.number; stampGhost.style.color=textColor(colorByType.number);
});
stampBadge.addEventListener('dragend', ()=>{ stampGhost.style.display='none'; });
document.addEventListener('dragover', e=>{ stampGhost.style.left=e.clientX+'px'; stampGhost.style.top=e.clientY+'px'; });

const canvasArea = document.getElementById('canvas-area');
canvasArea.addEventListener('dragover', e=>{ e.preventDefault(); e.dataTransfer.dropEffect='copy'; });
canvasArea.addEventListener('drop', e=>{
  e.preventDefault();
  if (e.dataTransfer.getData('text/plain')==='stamp') {
    stampGhost.style.display='none';
    if (!imageLoaded||!stage) return;
    const rect=document.getElementById('konva-container').getBoundingClientRect();
    const x=(e.clientX-rect.left)/viewScale-STAGE_PAD;
    const y=(e.clientY-rect.top)/viewScale-STAGE_PAD;
    const prev=activeTool; activeTool='number'; placeNumber(x,y); activeTool=prev;
  } else {
    const file=e.dataTransfer.files[0]; if(file) handleImageFile(file);
  }
});

// ─── Middle-click Pan ────────────────────────────────────────────────────────
canvasArea.addEventListener('mousedown', e=>{
  if (e.button===1) {
    e.preventDefault(); isPanning=true;
    panStartX=e.clientX; panStartY=e.clientY;
    panScrollLeft=canvasArea.scrollLeft; panScrollTop=canvasArea.scrollTop;
    canvasArea.style.cursor='grabbing';
  }
});
document.addEventListener('mousemove', e=>{
  if (!isPanning) return;
  canvasArea.scrollLeft=panScrollLeft-(e.clientX-panStartX);
  canvasArea.scrollTop=panScrollTop-(e.clientY-panStartY);
});
document.addEventListener('mouseup', e=>{
  if (e.button===1&&isPanning) { isPanning=false; canvasArea.style.cursor=''; }
});

// Fix: if mouse released outside stage, clean up marquee/drawing state
document.addEventListener('mouseup', e=>{
  if (e.button !== 0) return;
  if (marqueeActive) {
    marqueeActive = false;
    if (marqueeRect) { marqueeRect.destroy(); marqueeRect = null; drawLayer.batchDraw(); }
  }
  if (isDrawing && tempShape) {
    isDrawing = false;
    tempShape.destroy(); tempShape = null; drawLayer.batchDraw();
  }
});

// ─── Scroll Zoom ─────────────────────────────────────────────────────────────
canvasArea.addEventListener('wheel', e=>{
  if (!imageLoaded) return;
  e.preventDefault();
  if (e.altKey) {
    const delta = e.deltaY > 0 ? -1 : 1;
    if (selectedIds.size > 0) {
      let changed = false;
      selectedIds.forEach(id=>{
        const ann=findAnn(id);
        if(ann&&ann.type==='number'){
          ann.radius=Math.max(6,Math.min(100,ann.radius+delta));
          changed=true;
        }
      });
      if(changed){
        renderAnnotations();
        updateStampBadge();
        const firstNum=[...selectedIds].map(id=>findAnn(id)).find(a=>a&&a.type==='number');
        if(firstNum){
          document.getElementById('prop-size').value=firstNum.radius;
          document.getElementById('prop-size-val').textContent=firstNum.radius;
          document.getElementById('prop-size-num').value=firstNum.radius;
        }
      }
    } else {
      defaultRadius=Math.max(6,Math.min(100,defaultRadius+delta));
      document.getElementById('prop-size').value=defaultRadius;
      document.getElementById('prop-size-val').textContent=defaultRadius;
      document.getElementById('prop-size-num').value=defaultRadius;
      updateStampBadge();
    }
    return;
  }
  const PAD_PX = 500;
  const rect = canvasArea.getBoundingClientRect();
  const mouseX = e.clientX - rect.left + canvasArea.scrollLeft;
  const mouseY = e.clientY - rect.top + canvasArea.scrollTop;
  const oldScale = viewScale;
  const delta2 = e.deltaY > 0 ? -0.1 : 0.1;
  setZoom(viewScale + delta2);
  const ratio = viewScale / oldScale;
  canvasArea.scrollLeft = mouseX * ratio + PAD_PX * (1 - ratio) - (e.clientX - rect.left);
  canvasArea.scrollTop = mouseY * ratio + PAD_PX * (1 - ratio) - (e.clientY - rect.top);
}, {passive:false});

// ─── Notes resize ─────────────────────────────────────────────────────────────
let notesResizing=false, notesResizeStartX=0, notesResizeStartW=0;
document.getElementById('notes-resize-handle').addEventListener('mousedown', e=>{
  notesResizing=true; notesResizeStartX=e.clientX; notesResizeStartW=document.getElementById('notes-panel').offsetWidth; e.preventDefault();
});
document.addEventListener('mousemove', e=>{
  if (!notesResizing) return;
  const newW=Math.max(120,Math.min(500,notesResizeStartW-(e.clientX-notesResizeStartX)));
  document.getElementById('notes-panel').style.width=newW+'px';
});
document.addEventListener('mouseup', ()=>{ notesResizing=false; });

// ─── Event Wiring ─────────────────────────────────────────────────────────────
document.querySelectorAll('.tool-btn').forEach(btn=>btn.addEventListener('click',()=>setTool(btn.dataset.tool)));
document.getElementById('btn-undo').addEventListener('click', undo);
document.getElementById('btn-redo').addEventListener('click', redo);
document.getElementById('btn-open').addEventListener('click',()=>document.getElementById('file-input').click());
document.getElementById('file-input').addEventListener('change', e=>{ if(e.target.files[0]) handleImageFile(e.target.files[0]); e.target.value=''; });
document.getElementById('btn-export').addEventListener('click', exportPNG);
document.getElementById('btn-save-json').addEventListener('click', saveJSON);
document.getElementById('btn-load-json').addEventListener('click', () => document.getElementById('json-file-input').click());
document.getElementById('json-file-input').addEventListener('change', e => { if(e.target.files[0]) loadJSON(e.target.files[0]); e.target.value=''; });
document.getElementById('export-cancel-btn').addEventListener('click',()=>document.getElementById('export-panel').classList.add('hidden'));
document.getElementById('export-save-btn').addEventListener('click',()=>{
  const choice=document.querySelector('input[name="export-bg"]:checked').value;
  document.getElementById('export-panel').classList.add('hidden');
  doExport(choice);
});
document.getElementById('btn-delete').addEventListener('click', deleteSelected);
document.getElementById('btn-select-all').addEventListener('click',()=>{ selectedIds=new Set(annotations.map(a=>a.id)); updateSelectionVisual(); updatePropsPanel(); saveHistory(); });
document.getElementById('btn-deselect-all').addEventListener('click', deselectAll);
document.getElementById('btn-plus').addEventListener('click',()=>{ nextNum++; updateStampBadge(); });
document.getElementById('btn-minus').addEventListener('click',()=>{ if(nextNum>1){nextNum--;updateStampBadge();} });

// 次番号を揃える: sets nextNum to max+1; resets to 1 if no annotations
document.getElementById('btn-next-auto').addEventListener('click',()=>{
  const nums = annotations.filter(a => a.type==='number');
  if (nums.length === 0) { nextNum = 1; updateStampBadge(); return; }
  const max = Math.max(...nums.map(a => parseInt(a.num)||1));
  nextNum = max + 1;
  updateStampBadge();
});

document.getElementById('next-num-input').addEventListener('change', e=>{
  const v=normalizeNumericInput(e.target.value);
  if (v!==null) { nextNum=v; updateStampBadge(); } else { e.target.value=nextNum; }
});
document.getElementById('next-num-input').addEventListener('input', e=>{
  e.target.value=e.target.value.replace(/[０-９]/g,c=>String.fromCharCode(c.charCodeAt(0)-0xFEE0));
});

// Style selector
function applyNumStyle(style) {
  defaultNumStyle = style;
  updateStyleSelector(); updateStampBadge();
  if (selectedIds.size > 0) {
    let changed = false;
    selectedIds.forEach(id => { const ann = findAnn(id); if (ann && ann.type === 'number') { ann.style = style; changed = true; } });
    if (changed) { saveHistory(); renderAnnotations(); }
  }
  updatePropsPanel();
}
document.querySelectorAll('#style-selector .style-btn').forEach(btn => {
  btn.addEventListener('click', () => applyNumStyle(btn.dataset.style));
});
document.querySelectorAll('#prop-style-selector .style-btn').forEach(btn => {
  btn.addEventListener('click', () => applyNumStyle(btn.dataset.style));
});

// Color type tabs
document.querySelectorAll('.color-type-tab').forEach(tab=>{
  tab.addEventListener('click',()=>{
    setActiveColorTab(tab.dataset.type);
    if (activeTool!=='select') setTool(tab.dataset.type);
  });
});

// Custom color picker
document.getElementById('custom-color').addEventListener('input', e=>setColor(e.target.value));
document.getElementById('custom-color').addEventListener('click', e=>{
  if (selectedIds.size>0) setColor(e.target.value);
});

// Props: number value
document.getElementById('prop-num-input').addEventListener('change', e=>{
  if (selectedIds.size!==1) return;
  const ann=findAnn([...selectedIds][0]);
  if(ann&&ann.type==='number'){
    const v=normalizeNumericInput(e.target.value);
    if(v!==null){ann.num=v;saveHistory();renderAnnotations();renderNotesPanel();updateSelectionVisual();}
    else e.target.value=ann.num;
  }
});
document.getElementById('prop-num-input').addEventListener('input', e=>{
  e.target.value=e.target.value.replace(/[０-９]/g,c=>String.fromCharCode(c.charCodeAt(0)-0xFEE0));
});
document.getElementById('prop-num-minus').addEventListener('click',()=>{
  if(selectedIds.size!==1)return;
  const ann=findAnn([...selectedIds][0]);
  if(ann&&ann.type==='number'){ann.num=(parseInt(ann.num)||1)-1;saveHistory();renderAnnotations();renderNotesPanel();document.getElementById('prop-num-input').value=ann.num;}
});
document.getElementById('prop-num-plus').addEventListener('click',()=>{
  if(selectedIds.size!==1)return;
  const ann=findAnn([...selectedIds][0]);
  if(ann&&ann.type==='number'){ann.num=(parseInt(ann.num)||1)+1;saveHistory();renderAnnotations();renderNotesPanel();document.getElementById('prop-num-input').value=ann.num;}
});

// Props: size
document.getElementById('prop-size').addEventListener('input', e=>{
  const val=parseInt(e.target.value);
  document.getElementById('prop-size-val').textContent=val;
  document.getElementById('prop-size-num').value=val;
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='number')ann.radius=val;});renderAnnotations();updateStampBadge();}
  else{defaultRadius=val;updateStampBadge();}
});
document.getElementById('prop-size').addEventListener('change',()=>{if(selectedIds.size>0)saveHistory();});
document.getElementById('prop-size-num').addEventListener('change', e=>{
  const val=Math.max(6,Math.min(100,parseInt(e.target.value)||6));
  e.target.value=val;
  document.getElementById('prop-size').value=val;
  document.getElementById('prop-size-val').textContent=val;
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='number')ann.radius=val;});renderAnnotations();updateStampBadge();saveHistory();}
  else{defaultRadius=val;updateStampBadge();}
});

// Props: rect stroke
document.getElementById('prop-rect-stroke').addEventListener('input', e=>{
  const val=parseInt(e.target.value);
  document.getElementById('prop-rect-stroke-val').textContent=val;
  document.getElementById('prop-rect-stroke-num').value=val;
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='rect')ann.strokeWidth=val;});renderAnnotations();}
  else defaultRectStroke=val;
  updateRectArrowPreview();
});
document.getElementById('prop-rect-stroke').addEventListener('change',()=>{if(selectedIds.size>0)saveHistory();});
document.getElementById('prop-rect-stroke-num').addEventListener('change', e=>{
  const val=Math.max(1,Math.min(20,parseInt(e.target.value)||1));
  e.target.value=val;
  document.getElementById('prop-rect-stroke').value=val;
  document.getElementById('prop-rect-stroke-val').textContent=val;
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='rect')ann.strokeWidth=val;});renderAnnotations();saveHistory();}
  else defaultRectStroke=val;
  updateRectArrowPreview();
});

// Props: arrow stroke
document.getElementById('prop-arrow-stroke').addEventListener('input', e=>{
  const val=parseInt(e.target.value);
  document.getElementById('prop-arrow-stroke-val').textContent=val;
  document.getElementById('prop-arrow-stroke-num').value=val;
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='arrow')ann.strokeWidth=val;});renderAnnotations();}
  else defaultArrowStroke=val;
  updateRectArrowPreview();
});
document.getElementById('prop-arrow-stroke').addEventListener('change',()=>{if(selectedIds.size>0)saveHistory();});
document.getElementById('prop-arrow-stroke-num').addEventListener('change', e=>{
  const val=Math.max(1,Math.min(20,parseInt(e.target.value)||1));
  e.target.value=val;
  document.getElementById('prop-arrow-stroke').value=val;
  document.getElementById('prop-arrow-stroke-val').textContent=val;
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='arrow')ann.strokeWidth=val;});renderAnnotations();saveHistory();}
  else defaultArrowStroke=val;
  updateRectArrowPreview();
});

// Props: fill opacity
document.getElementById('prop-fill-opacity').addEventListener('input', e=>{
  const val=parseInt(e.target.value);
  document.getElementById('prop-fill-opacity-val').textContent=val;
  const opacity=val/100;
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='rect')ann.fillOpacity=opacity;});renderAnnotations();}
  else{defaultFillOpacity=opacity;if(defaultFillHatch)defaultFillOpacityHatch=opacity;else defaultFillOpacitySolid=opacity;}
  updateRectArrowPreview();
});
document.getElementById('prop-fill-opacity').addEventListener('change',()=>{if(selectedIds.size>0)saveHistory();});

// Props: fill hatch
document.getElementById('fill-solid-btn').addEventListener('click',()=>{
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='rect')ann.fillHatch=false;});saveHistory();renderAnnotations();updatePropsPanel();}
  else{
    defaultFillHatch=false; updateFillHatchToggle(false);
    defaultFillOpacity=defaultFillOpacitySolid;
    const pct=Math.round(defaultFillOpacitySolid*100);
    document.getElementById('prop-fill-opacity').value=pct;
    document.getElementById('prop-fill-opacity-val').textContent=pct;
  }
  updateRectArrowPreview();
});
document.getElementById('fill-hatch-btn').addEventListener('click',()=>{
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='rect')ann.fillHatch=true;});saveHistory();renderAnnotations();updatePropsPanel();}
  else{
    defaultFillHatch=true; updateFillHatchToggle(true);
    defaultFillOpacity=defaultFillOpacityHatch;
    const pct=Math.round(defaultFillOpacityHatch*100);
    document.getElementById('prop-fill-opacity').value=pct;
    document.getElementById('prop-fill-opacity-val').textContent=pct;
  }
  updateRectArrowPreview();
});

// Props: arrow type
document.getElementById('arrow-type-arrow-btn').addEventListener('click',()=>{
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='arrow')ann.arrowType='arrow';});saveHistory();renderAnnotations();updatePropsPanel();}
  else{defaultArrowType='arrow';updateArrowTypeToggle('arrow');}
  updateRectArrowPreview();
});
document.getElementById('arrow-type-line-btn').addEventListener('click',()=>{
  if(selectedIds.size>0){selectedIds.forEach(id=>{const ann=findAnn(id);if(ann&&ann.type==='arrow')ann.arrowType='line';});saveHistory();renderAnnotations();updatePropsPanel();}
  else{defaultArrowType='line';updateArrowTypeToggle('line');}
  updateRectArrowPreview();
});

// Alignment
document.querySelectorAll('.align-btn').forEach(btn=>btn.addEventListener('click',()=>alignSelected(btn.dataset.align)));

// Zoom controls
document.getElementById('zoom-slider').addEventListener('input', e=>{
  const rect = canvasArea.getBoundingClientRect();
  const cx = canvasArea.scrollLeft + rect.width / 2;
  const cy = canvasArea.scrollTop + rect.height / 2;
  const oldScale = viewScale;
  setZoom(parseInt(e.target.value) / 100);
  const ratio = viewScale / oldScale;
  canvasArea.scrollLeft = cx * ratio - rect.width / 2;
  canvasArea.scrollTop = cy * ratio - rect.height / 2;
});
document.getElementById('btn-zoom-in').addEventListener('click',()=>setZoom(viewScale+0.1));
document.getElementById('btn-zoom-out').addEventListener('click',()=>setZoom(viewScale-0.1));

// Renumber
document.getElementById('btn-renumber').addEventListener('click', renumber);

// Notes
document.getElementById('btn-toggle-notes').addEventListener('click',()=>{
  notesVisible=!notesVisible;
  document.getElementById('notes-panel').classList.toggle('hidden',!notesVisible);
  document.getElementById('btn-toggle-notes').classList.toggle('active',notesVisible);
});
document.getElementById('notes-sort').addEventListener('change', renderNotesPanel);
document.getElementById('btn-copy-notes').addEventListener('click',()=>{
  const nums=getSortedNotes();
  const text=nums.map(a=>`${a.num}. ${a.label||''}`).join('\n');
  navigator.clipboard.writeText(text).catch(()=>{
    const ta=document.createElement('textarea');ta.value=text;document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);
  });
});

// Click empty canvas to open file
canvasArea.addEventListener('click', e=>{
  if (!imageLoaded) document.getElementById('file-input').click();
});

// Drag canvas highlight
canvasArea.addEventListener('dragenter', e=>{ if(e.dataTransfer.types.includes('Files'))canvasArea.classList.add('drag-over'); });
canvasArea.addEventListener('dragleave', e=>{ if(!canvasArea.contains(e.relatedTarget))canvasArea.classList.remove('drag-over'); });
canvasArea.addEventListener('drop',()=>canvasArea.classList.remove('drag-over'));

// Keyboard
document.addEventListener('keydown', e=>{
  if (e.key==='Shift') shiftPressed=true;
  const tag=document.activeElement.tagName.toLowerCase();
  if (tag==='input'||tag==='textarea') return;
  if (e.key==='Escape') deselectAll();
  if (e.key==='v'||e.key==='V') setTool('select');
  if (e.key==='n'||e.key==='N') setTool('number');
  if (e.key==='r'||e.key==='R') setTool('rect');
  if (e.key==='a'||e.key==='A') setTool('arrow');
  if (e.key==='Delete'||e.key==='Backspace') deleteSelected();
  if (e.ctrlKey&&!e.shiftKey&&e.key==='z'){e.preventDefault();undo();}
  if (e.ctrlKey&&!e.shiftKey&&(e.key==='y'||e.key==='Y')){e.preventDefault();redo();}
  if (e.ctrlKey&&e.shiftKey&&(e.key==='z'||e.key==='Z')){e.preventDefault();redo();}
  // Copy / Paste
  if (e.ctrlKey&&!e.shiftKey&&(e.key==='c'||e.key==='C')) { e.preventDefault(); copySelected(); }
  if (e.ctrlKey&&!e.shiftKey&&(e.key==='v'||e.key==='V')) {
    if (clipboard.length > 0 && imageLoaded) { e.preventDefault(); pasteAnnotations(); }
  }
  // Arrow key movement
  if (activeTool==='select'&&selectedIds.size>0) {
    const step=e.shiftKey?10:1;
    let dx=0,dy=0;
    if(e.key==='ArrowLeft')dx=-step; if(e.key==='ArrowRight')dx=step;
    if(e.key==='ArrowUp')dy=-step; if(e.key==='ArrowDown')dy=step;
    if(dx!==0||dy!==0){
      e.preventDefault();
      selectedIds.forEach(id=>{
        const ann=findAnn(id); if(!ann)return;
        if(ann.type==='number'||ann.type==='rect'){ann.x+=dx;ann.y+=dy;}
        else if(ann.type==='arrow'){ann.points=[ann.points[0]+dx,ann.points[1]+dy,ann.points[2]+dx,ann.points[3]+dy];}
      });
      saveHistory(); renderAnnotations();
    }
  }
});
document.addEventListener('keyup', e=>{ if(e.key==='Shift') shiftPressed=false; });

// Ctrl+V: image paste (only when annotation clipboard is empty)
document.addEventListener('paste', async e=>{
  const items=e.clipboardData?.items; if(!items)return;
  // If annotation clipboard is occupied, Ctrl+V was already handled in keydown
  if (clipboard.length > 0 && imageLoaded) return;
  for(const item of items){if(item.type.startsWith('image/')){const file=item.getAsFile();if(file)handleImageFile(file);break;}}
});

// beforeunload
window.addEventListener('beforeunload', e=>{
  if(annotations.length>0){e.preventDefault();e.returnValue='';}
});

// ─── Init ─────────────────────────────────────────────────────────────────────
(function init(){
  buildPalette();
  updateStampBadge();
  updateStyleSelector();
  updateUndoRedoBtns();
  initKonva();
  setActiveColorTab('number');
  updateColorDots();
  setTool('number');
  saveHistory();
})();
