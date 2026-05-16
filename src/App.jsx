import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import {
  BarChart, Bar, PieChart, Pie, LineChart, Line, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell
} from 'recharts';
import * as XLSX from 'xlsx';
import {
  HardHat, Plus, ClipboardList, BarChart3, Home, Search, Eye, Trash2,
  X, Check, AlertTriangle, Pen, FileSpreadsheet, Printer, ChevronLeft,
  Filter, Loader2, Shield, Hash, Calendar, User, Building2, Briefcase,
  Tag, FileText, RotateCcw, Download
} from 'lucide-react';

// ============================================================================
// CONSTANTES
// ============================================================================

const EPP_OPTIONS = [
  'Casco', 'Lentes de seguridad', 'Tapones auditivos', 'Guantes Hyflex',
  'Guantes de nitrilo', 'Mascarilla de media cara', 'Full Face', 'Filtros',
  'Arnés para mascarilla de media cara', 'Mascarilla P95', 'Tyvek', 'Tychem',
  'Playera', 'Pantalón', 'Calzado de seguridad', 'Bata', 'Otro'
];

const REASON_OPTIONS = [
  'Nuevo ingreso',
  'Reposición por daño',
  'Reposición por extravío',
  'Cambio por vencimiento',
  'Actualización por cambio de área o actividad'
];

const CONFORMITY_TEXT =
  'Declaro que he recibido el equipo de protección personal descrito en este registro en buen estado físico y funcional. Asimismo, confirmo que se me brindó la información necesaria sobre su uso correcto, cuidado, conservación y reposición.';

const FONT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,600;12..96,700;12..96,800&family=JetBrains+Mono:wght@400;500;600&family=Outfit:wght@300;400;500;600;700&display=swap');
.font-display{font-family:'Bricolage Grotesque',ui-sans-serif,sans-serif;letter-spacing:-0.02em}
.font-mono{font-family:'JetBrains Mono',ui-monospace,monospace}
.font-body{font-family:'Outfit',ui-sans-serif,sans-serif}
body,#root{font-family:'Outfit',ui-sans-serif,sans-serif}
@media print{
  .no-print{display:none !important}
  .print-page{padding:2rem;background:white !important}
  body{background:white !important}
}
@keyframes slide-up{
  from{opacity:0;transform:translate(-50%, 20px)}
  to{opacity:1;transform:translate(-50%, 0)}
}
.animate-slide-up{animation:slide-up 0.3s ease-out}
`;

// ============================================================
// CAPA DE DATOS (Google Sheets)
// ============================================================

const API_URL = 'https://script.google.com/macros/s/AKfycbyZdQb_QLicuaviucB2s8oWz2t0MaSXjJvY2jHxsfvcAhh0HovCFNSGIHN9ptOziK1wIw/exec';

const db = {
  async getIndex() {
    try {
      const res = await fetch(`${API_URL}?action=getAll`, {
        method: 'GET',
        redirect: 'follow',
      });
      const text = await res.text();
      const data = JSON.parse(text);
      return data.entregas || [];
    } catch(e) { console.error('getIndex error:', e); return []; }
  },
  async saveDelivery({ signature, ...delivery }) {
    try {
      const res = await fetch(API_URL, {
        method: 'POST',
        redirect: 'follow',
        body: JSON.stringify({ action: 'save', delivery, signature })
      });
      const text = await res.text();
      return JSON.parse(text);
    } catch(e) { throw e; }
  },
async deleteDelivery(folio) {
    try {
      const folios = Array.isArray(folio) ? folio : [folio];
      const res = await fetch(API_URL, {
        method: 'POST',
        redirect: 'follow',
        body: JSON.stringify({ action: 'delete', folio: folios })
      });
      const text = await res.text();
      return JSON.parse(text);
    } catch(e) { throw e; }
  },
  async getSignature(folio) {
    try {
      const res = await fetch(`${API_URL}?action=getSignature&folio=${folio}`, {
        redirect: 'follow',
      });
      const text = await res.text();
      const data = JSON.parse(text);
      return data.url || null;
    } catch { return null; }
  },
  async nextFolio() {
    const idx = await this.getIndex();
    const year = new Date().getFullYear();
    const n = (idx.length + 1).toString().padStart(4, '0');
    return `EPP-${year}-${n}`;
  }
};

// ============================================================================
// HELPERS
// ============================================================================

const fmtDate = (iso) => {
  if (!iso) return '—';
  const d = new Date(iso);
  return d.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' });
};

const fmtDateTime = (iso) => {
  if (!iso) return '—';
  const d = new Date(iso);
  return d.toLocaleString('es-MX', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
};

const monthKey = (iso) => {
  const d = new Date(iso);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
};

const monthLabel = (key) => {
  const [y, m] = key.split('-');
  const names = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  return `${names[parseInt(m, 10) - 1]} ${y}`;
};

// ============================================================================
// SIGNATURE PAD
// ============================================================================

function SignaturePad({ onChange, value, error }) {
  const canvasRef = useRef(null);
  const drawingRef = useRef(false);
  const lastRef = useRef({ x: 0, y: 0 });
  const [hasContent, setHasContent] = useState(!!value);

  const setupCanvas = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    canvas.width = rect.width * dpr;
    canvas.height = rect.height * dpr;
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#0c0a09';
    ctx.lineWidth = 2.2;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, rect.width, rect.height);
    if (value) {
      const img = new Image();
      img.onload = () => ctx.drawImage(img, 0, 0, rect.width, rect.height);
      img.src = value;
    }
  }, [value]);

  useEffect(() => {
    setupCanvas();
    const onResize = () => setupCanvas();
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, [setupCanvas]);

  const getPos = (e) => {
    const rect = canvasRef.current.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };

  const start = (e) => {
    e.preventDefault();
    drawingRef.current = true;
    lastRef.current = getPos(e);
    canvasRef.current.setPointerCapture(e.pointerId);
  };

  const move = (e) => {
    if (!drawingRef.current) return;
    e.preventDefault();
    const p = getPos(e);
    const ctx = canvasRef.current.getContext('2d');
    ctx.beginPath();
    ctx.moveTo(lastRef.current.x, lastRef.current.y);
    ctx.lineTo(p.x, p.y);
    ctx.stroke();
    lastRef.current = p;
    if (!hasContent) setHasContent(true);
  };

  const end = (e) => {
    if (!drawingRef.current) return;
    drawingRef.current = false;
    try { canvasRef.current.releasePointerCapture(e.pointerId); } catch {}
    onChange(canvasRef.current.toDataURL('image/png'));
  };

  const clear = () => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, rect.width, rect.height);
    setHasContent(false);
    onChange('');
  };

  return (
    <div>
      <div className={`relative border-2 ${error ? 'border-red-500' : 'border-stone-300'} rounded-lg overflow-hidden bg-white transition-colors`}>
        <canvas
          ref={canvasRef}
          className="w-full block touch-none cursor-crosshair"
          style={{ height: '200px' }}
          onPointerDown={start}
          onPointerMove={move}
          onPointerUp={end}
          onPointerCancel={end}
          onPointerLeave={end}
        />
        {!hasContent && (
          <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
            <div className="flex items-center gap-2 text-stone-400 font-body">
              <Pen size={16} />
              <span className="text-sm">Firma aquí con el dedo, mouse o stylus</span>
            </div>
          </div>
        )}
        <div className="absolute bottom-0 left-0 right-0 border-t border-stone-200 bg-stone-50 px-3 py-1.5 flex items-center justify-between">
          <span className="text-[10px] uppercase tracking-widest text-stone-500 font-mono">x ____________________</span>
          <button
            type="button"
            onClick={clear}
            className="text-xs text-stone-600 hover:text-stone-900 flex items-center gap-1 px-2 py-0.5 rounded hover:bg-stone-200 transition"
          >
            <RotateCcw size={12} /> Limpiar
          </button>
        </div>
      </div>
      {error && <p className="text-red-600 text-xs mt-1.5 flex items-center gap-1"><AlertTriangle size={12} />{error}</p>}
    </div>
  );
}

// ============================================================================
// UI PRIMITIVES
// ============================================================================

const Field = ({ label, icon: Icon, required, error, children, hint }) => (
  <div>
    <label className="block mb-1.5">
      <span className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wider text-stone-700">
        {Icon && <Icon size={12} />}
        {label}
        {required && <span className="text-red-500">*</span>}
      </span>
      {hint && <span className="text-[11px] text-stone-500 font-normal normal-case tracking-normal">{hint}</span>}
    </label>
    {children}
    {error && <p className="text-red-600 text-xs mt-1 flex items-center gap-1"><AlertTriangle size={11} />{error}</p>}
  </div>
);

const Input = ({ error, ...props }) => (
  <input
    {...props}
    className={`w-full px-3 py-2.5 bg-white border ${error ? 'border-red-400' : 'border-stone-300'} rounded-lg text-stone-900 placeholder:text-stone-400 focus:outline-none focus:ring-2 focus:ring-yellow-400 focus:border-stone-900 transition text-sm`}
  />
);

const Select = ({ error, children, ...props }) => (
  <select
    {...props}
    className={`w-full px-3 py-2.5 bg-white border ${error ? 'border-red-400' : 'border-stone-300'} rounded-lg text-stone-900 focus:outline-none focus:ring-2 focus:ring-yellow-400 focus:border-stone-900 transition text-sm`}
  >
    {children}
  </select>
);

const StatCard = ({ label, value, icon: Icon, accent }) => (
  <div className="bg-white border border-stone-200 rounded-xl p-4 sm:p-5 relative overflow-hidden group hover:border-stone-400 transition">
    <div className="flex items-start justify-between mb-3">
      <span className="text-[10px] font-semibold uppercase tracking-widest text-stone-500">{label}</span>
      <div className={`p-1.5 rounded-md ${accent || 'bg-stone-100 text-stone-700'}`}>
        <Icon size={14} />
      </div>
    </div>
    <div className="font-display text-3xl sm:text-4xl font-bold text-stone-900 leading-none">{value}</div>
  </div>
);

// ============================================================================
// FORMULARIO DE ENTREGA
// ============================================================================

function DeliveryForm({ onSuccess, onCancel }) {
  const [folio, setFolio] = useState('—');
  const [form, setForm] = useState({
    empleado_nombre: '',
    empleado_numero: '',
    puesto: '',
    area: '',
    fecha_entrega: new Date().toISOString().slice(0, 10),
    epp: [],
    epp_otro: '',
    motivo: '',
    entrega_nombre: '',
    entrega_cargo: '',
    registrado_por: ''
  });
  const [signature, setSignature] = useState('');
  const [errors, setErrors] = useState({});
  const [saving, setSaving] = useState(false);
  const [savedFolio, setSavedFolio] = useState(null);

  useEffect(() => {
    db.nextFolio().then(setFolio);
  }, []);

  const toggleEpp = (item) => {
    setForm(f => ({
      ...f,
      epp: f.epp.includes(item) ? f.epp.filter(x => x !== item) : [...f.epp, item]
    }));
  };

  const validate = () => {
    const e = {};
    if (!form.empleado_nombre.trim()) e.empleado_nombre = 'Requerido';
    if (!form.empleado_numero.trim()) e.empleado_numero = 'Requerido';
    if (!form.puesto.trim()) e.puesto = 'Requerido';
    if (!form.area.trim()) e.area = 'Requerido';
    if (!form.fecha_entrega) e.fecha_entrega = 'Requerido';
    if (form.epp.length === 0) e.epp = 'Selecciona al menos un EPP';
    if (form.epp.includes('Otro') && !form.epp_otro.trim()) e.epp_otro = 'Especifica qué EPP';
    if (!form.motivo) e.motivo = 'Requerido';
    if (!form.entrega_nombre.trim()) e.entrega_nombre = 'Requerido';
    if (!form.entrega_cargo.trim()) e.entrega_cargo = 'Requerido';
    if (!signature) e.signature = 'La firma del trabajador es obligatoria';
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const submit = async () => {
    if (!validate()) {
      setTimeout(() => {
        const el = document.querySelector('[data-error="true"]');
        el?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }, 50);
      return;
    }
    setSaving(true);
    const newFolio = await db.nextFolio();
    const record = {
      folio: newFolio,
      fecha_registro: new Date().toISOString(),
      conformidad: CONFORMITY_TEXT,
      ...form,
      signature
    };
    await db.saveDelivery(record);
    setSaving(false);
    setSavedFolio(newFolio);
    setTimeout(() => onSuccess(newFolio), 1600);
  };

  if (savedFolio) {
return (
  <div key="success-message" className="bg-white border border-stone-200 rounded-2xl p-8 sm:p-12 text-center">
        <div className="inline-flex items-center justify-center w-16 h-16 rounded-full bg-green-100 text-green-700 mb-4">
          <Check size={32} strokeWidth={3} />
        </div>
        <h2 className="font-display text-2xl sm:text-3xl font-bold text-stone-900 mb-2">Registro guardado</h2>
        <p className="text-stone-600 mb-3">La entrega de EPP fue registrada exitosamente.</p>
        <div className="inline-block bg-stone-900 text-yellow-400 font-mono px-4 py-2 rounded-lg">{savedFolio}</div>
      </div>
    );
  }

return (
  <div key="form-fields" className="space-y-6">
      {/* Header */}
      <div className="bg-stone-950 text-white rounded-2xl p-5 sm:p-6 relative overflow-hidden">
        <div className="absolute top-0 right-0 w-32 h-32 bg-yellow-400 opacity-10 rounded-full -mr-12 -mt-12" />
        <div className="relative flex items-start justify-between gap-4 flex-wrap">
          <div>
            <span className="text-[10px] uppercase tracking-widest text-yellow-400 font-semibold">Nuevo registro</span>
            <h1 className="font-display text-2xl sm:text-3xl font-bold mt-1">Entrega de EPP</h1>
          </div>
          <div className="bg-yellow-400 text-stone-950 px-3 py-2 rounded-lg">
            <div className="text-[9px] uppercase tracking-widest font-semibold opacity-70">Folio</div>
            <div className="font-mono font-bold text-sm sm:text-base">{folio}</div>
          </div>
        </div>
      </div>

      {/* Datos del empleado */}
      <section className="bg-white border border-stone-200 rounded-2xl p-5 sm:p-6">
        <h2 className="font-display text-lg font-bold text-stone-900 mb-4 flex items-center gap-2">
          <span className="w-1 h-5 bg-yellow-400 rounded-full" />
          Datos del trabajador
        </h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div data-error={!!errors.empleado_nombre}>
            <Field label="Nombre del empleado" icon={User} required error={errors.empleado_nombre}>
              <Input
                value={form.empleado_nombre}
                onChange={e => setForm({ ...form, empleado_nombre: e.target.value })}
                placeholder="Ej. María Hernández López"
                error={errors.empleado_nombre}
              />
            </Field>
          </div>
          <div data-error={!!errors.empleado_numero}>
            <Field label="Número de empleado" icon={Hash} required error={errors.empleado_numero}>
              <Input
                value={form.empleado_numero}
                onChange={e => setForm({ ...form, empleado_numero: e.target.value })}
                placeholder="Ej. 04521"
                error={errors.empleado_numero}
              />
            </Field>
          </div>
          <div data-error={!!errors.puesto}>
            <Field label="Puesto" icon={Briefcase} required error={errors.puesto}>
              <Input
                value={form.puesto}
                onChange={e => setForm({ ...form, puesto: e.target.value })}
                placeholder="Ej. Operador de planta"
                error={errors.puesto}
              />
            </Field>
          </div>
          <div data-error={!!errors.area}>
            <Field label="Área o proceso" icon={Building2} required error={errors.area}>
              <Input
                value={form.area}
                onChange={e => setForm({ ...form, area: e.target.value })}
                placeholder="Ej. Producción - Línea 3"
                error={errors.area}
              />
            </Field>
          </div>
          <div data-error={!!errors.fecha_entrega}>
            <Field label="Fecha de entrega" icon={Calendar} required error={errors.fecha_entrega}>
              <Input
                type="date"
                value={form.fecha_entrega}
                onChange={e => setForm({ ...form, fecha_entrega: e.target.value })}
                error={errors.fecha_entrega}
              />
            </Field>
          </div>
          <div data-error={!!errors.motivo}>
            <Field label="Motivo de la entrega" icon={Tag} required error={errors.motivo}>
              <Select
                value={form.motivo}
                onChange={e => setForm({ ...form, motivo: e.target.value })}
                error={errors.motivo}
              >
                <option value="">— Selecciona —</option>
                {REASON_OPTIONS.map(r => <option key={r} value={r}>{r}</option>)}
              </Select>
            </Field>
          </div>
        </div>
      </section>

      {/* EPP entregado */}
      <section className="bg-white border border-stone-200 rounded-2xl p-5 sm:p-6" data-error={!!errors.epp}>
        <h2 className="font-display text-lg font-bold text-stone-900 mb-1 flex items-center gap-2">
          <span className="w-1 h-5 bg-yellow-400 rounded-full" />
          EPP entregado
        </h2>
        <p className="text-xs text-stone-500 mb-4">Puedes seleccionar uno o varios. Marca todos los EPP entregados en esta ocasión.</p>
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-2">
          {EPP_OPTIONS.map(item => {
            const checked = form.epp.includes(item);
            return (
              <label
                key={item}
                className={`flex items-center gap-2 px-3 py-2.5 rounded-lg border cursor-pointer transition text-sm select-none ${
                  checked
                    ? 'bg-stone-950 border-stone-950 text-white shadow-sm'
                    : 'bg-stone-50 border-stone-200 text-stone-700 hover:border-stone-400 hover:bg-stone-100'
                }`}
              >
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => toggleEpp(item)}
                  className="sr-only"
                />
                <div className={`w-4 h-4 rounded border-2 flex items-center justify-center flex-shrink-0 ${
                  checked ? 'bg-yellow-400 border-yellow-400' : 'bg-white border-stone-400'
                }`}>
                  {checked && <Check size={12} strokeWidth={3} className="text-stone-950" />}
                </div>
                <span>{item}</span>
              </label>
            );
          })}
        </div>
        {errors.epp && <p className="text-red-600 text-xs mt-2 flex items-center gap-1"><AlertTriangle size={11} />{errors.epp}</p>}

        {form.epp.includes('Otro') && (
          <div className="mt-4 pt-4 border-t border-stone-200" data-error={!!errors.epp_otro}>
            <Field label='Especifica el EPP "Otro"' required error={errors.epp_otro}>
              <Input
                value={form.epp_otro}
                onChange={e => setForm({ ...form, epp_otro: e.target.value })}
                placeholder="Describe el equipo entregado"
                error={errors.epp_otro}
              />
            </Field>
          </div>
        )}
      </section>

      {/* Conformidad y firma */}
      <section className="bg-white border border-stone-200 rounded-2xl p-5 sm:p-6">
        <h2 className="font-display text-lg font-bold text-stone-900 mb-4 flex items-center gap-2">
          <span className="w-1 h-5 bg-yellow-400 rounded-full" />
          Conformidad del trabajador
        </h2>
        <div className="bg-stone-50 border-l-4 border-yellow-400 rounded-r-lg p-4 mb-5">
          <p className="text-sm text-stone-700 leading-relaxed">{CONFORMITY_TEXT}</p>
        </div>
        <div data-error={!!errors.signature}>
          <Field label="Firma digital del trabajador" icon={Pen} required hint="Use el dedo, mouse o stylus para firmar dentro del recuadro">
            <SignaturePad value={signature} onChange={setSignature} error={errors.signature} />
          </Field>
        </div>
      </section>

      {/* Datos del responsable */}
      <section className="bg-white border border-stone-200 rounded-2xl p-5 sm:p-6">
        <h2 className="font-display text-lg font-bold text-stone-900 mb-4 flex items-center gap-2">
          <span className="w-1 h-5 bg-yellow-400 rounded-full" />
          Quien entrega el equipo
        </h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div data-error={!!errors.entrega_nombre}>
            <Field label="Nombre" icon={User} required error={errors.entrega_nombre}>
              <Input
                value={form.entrega_nombre}
                onChange={e => setForm({ ...form, entrega_nombre: e.target.value })}
                placeholder="Nombre completo"
                error={errors.entrega_nombre}
              />
            </Field>
          </div>
          <div data-error={!!errors.entrega_cargo}>
            <Field label="Cargo" icon={Briefcase} required error={errors.entrega_cargo}>
              <Input
                value={form.entrega_cargo}
                onChange={e => setForm({ ...form, entrega_cargo: e.target.value })}
                placeholder="Ej. Coordinador HSE"
                error={errors.entrega_cargo}
              />
            </Field>
          </div>
          <div className="sm:col-span-2">
            <Field label="Usuario que registra (opcional)" icon={User}>
              <Input
                value={form.registrado_por}
                onChange={e => setForm({ ...form, registrado_por: e.target.value })}
                placeholder="Tu nombre o ID"
              />
            </Field>
          </div>
        </div>
      </section>

      {/* Actions */}
      <div className="sticky bottom-0 -mx-3 sm:mx-0 sm:bottom-4 z-10">
        <div className="bg-white sm:bg-stone-950 border-t sm:border border-stone-200 sm:border-stone-800 sm:rounded-2xl p-4 flex gap-3 shadow-lg">
          <button
            onClick={onCancel}
            disabled={saving}
            className="flex-1 sm:flex-none px-5 py-3 rounded-lg text-stone-700 sm:text-white bg-stone-100 sm:bg-stone-800 hover:bg-stone-200 sm:hover:bg-stone-700 font-semibold transition disabled:opacity-50 text-sm"
          >
            Cancelar
          </button>
          <button
            onClick={submit}
            disabled={saving}
            className="flex-1 px-5 py-3 rounded-lg bg-yellow-400 hover:bg-yellow-300 text-stone-950 font-semibold transition disabled:opacity-50 flex items-center justify-center gap-2 text-sm"
          >
            {saving ? <><Loader2 size={16} className="animate-spin" /> Guardando…</> : <><Check size={16} /> Guardar entrega</>}
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// VISTA DETALLE
// ============================================================================

function DeliveryDetail({ delivery, onClose, onDelete }) {
  const [signature, setSignature] = useState(null);
  const [loading, setLoading] = useState(true);
  const [confirmDel, setConfirmDel] = useState(false);

  useEffect(() => {
    setLoading(true);
    db.getSignature(delivery.folio).then(s => {
      setSignature(s);
      setLoading(false);
    });
  }, [delivery.folio]);

  const print = () => window.print();

  const allEpp = [
    ...delivery.epp.filter(e => e !== 'Otro'),
    ...(delivery.epp.includes('Otro') && delivery.epp_otro ? [`Otro: ${delivery.epp_otro}`] : [])
  ];

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-end sm:items-center justify-center p-0 sm:p-4 no-print">
      <div className="bg-stone-50 w-full max-w-3xl max-h-[95vh] sm:max-h-[90vh] rounded-t-3xl sm:rounded-2xl overflow-hidden flex flex-col shadow-2xl">
        <div className="bg-stone-950 text-white p-5 flex items-start justify-between gap-4 flex-shrink-0">
          <div>
            <span className="text-[10px] uppercase tracking-widest text-yellow-400 font-semibold">Comprobante de entrega</span>
            <div className="font-mono font-bold text-lg sm:text-xl mt-0.5">{delivery.folio}</div>
            <div className="text-xs text-stone-400 mt-1">Registrado el {fmtDateTime(delivery.fecha_registro)}</div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-stone-800 rounded-lg transition">
            <X size={20} />
          </button>
        </div>

        <div className="overflow-y-auto flex-1 print-page" id="print-area">
          <div className="p-5 sm:p-6 space-y-5">
            <DetailGrid title="Trabajador" rows={[
              ['Nombre', delivery.empleado_nombre],
              ['Número', delivery.empleado_numero],
              ['Puesto', delivery.puesto],
              ['Área o proceso', delivery.area]
            ]} />
            <DetailGrid title="Entrega" rows={[
              ['Fecha de entrega', fmtDate(delivery.fecha_entrega)],
              ['Motivo', delivery.motivo]
            ]} />
            <div>
              <h3 className="font-display font-bold text-stone-900 mb-2 text-sm uppercase tracking-wider">EPP entregado</h3>
              <div className="flex flex-wrap gap-1.5">
                {allEpp.map((e, i) => (
                  <span key={i} className="px-3 py-1 bg-stone-900 text-yellow-400 rounded-full text-xs font-medium">{e}</span>
                ))}
              </div>
            </div>
            <div>
              <h3 className="font-display font-bold text-stone-900 mb-2 text-sm uppercase tracking-wider">Conformidad</h3>
              <div className="bg-white border border-stone-200 rounded-lg p-4">
                <p className="text-xs sm:text-sm text-stone-700 leading-relaxed italic">"{delivery.conformidad || CONFORMITY_TEXT}"</p>
              </div>
            </div>
            <div>
              <h3 className="font-display font-bold text-stone-900 mb-2 text-sm uppercase tracking-wider">Firma del trabajador</h3>
              <div className="bg-white border border-stone-200 rounded-lg p-4">
                {loading ? (
                  <div className="h-24 flex items-center justify-center text-stone-400 text-sm">
                    <Loader2 size={16} className="animate-spin mr-2" /> Cargando firma…
                  </div>
                ) : signature ? (
                  <img src={signature} alt="Firma" className="w-full max-h-44 object-contain" />
                ) : (
                  <div className="h-24 flex items-center justify-center text-stone-400 text-sm">Sin firma disponible</div>
                )}
                <div className="mt-2 pt-2 border-t border-stone-200 text-center font-mono text-[10px] text-stone-500 uppercase tracking-widest">
                  {delivery.empleado_nombre}
                </div>
              </div>
            </div>
            <DetailGrid title="Entregado por" rows={[
              ['Nombre', delivery.entrega_nombre],
              ['Cargo', delivery.entrega_cargo],
              ['Registrado por', delivery.registrado_por || '—']
            ]} />
          </div>
        </div>

        <div className="border-t border-stone-200 bg-white p-3 flex gap-2 flex-shrink-0 no-print">
          {confirmDel ? (
            <>
              <button onClick={() => setConfirmDel(false)} className="flex-1 px-4 py-2.5 rounded-lg bg-stone-100 hover:bg-stone-200 text-stone-700 font-semibold text-sm transition">Cancelar</button>
              <button onClick={() => onDelete(delivery.folio)} className="flex-1 px-4 py-2.5 rounded-lg bg-red-600 hover:bg-red-700 text-white font-semibold text-sm transition flex items-center justify-center gap-1.5">
                <Trash2 size={14} /> Confirmar eliminación
              </button>
            </>
          ) : (
            <>
              <button onClick={() => setConfirmDel(true)} className="px-3 py-2.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-700 font-semibold text-sm transition flex items-center gap-1.5">
                <Trash2 size={14} />
              </button>
              <button onClick={print} className="flex-1 px-4 py-2.5 rounded-lg bg-stone-100 hover:bg-stone-200 text-stone-900 font-semibold text-sm transition flex items-center justify-center gap-1.5">
                <Printer size={14} /> Imprimir / PDF
              </button>
              <button onClick={onClose} className="flex-1 px-4 py-2.5 rounded-lg bg-stone-950 hover:bg-stone-800 text-white font-semibold text-sm transition">
                Cerrar
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

const DetailGrid = ({ title, rows }) => (
  <div>
    <h3 className="font-display font-bold text-stone-900 mb-2 text-sm uppercase tracking-wider">{title}</h3>
    <div className="bg-white border border-stone-200 rounded-lg divide-y divide-stone-100">
      {rows.map(([label, value]) => (
        <div key={label} className="px-4 py-2.5 flex justify-between gap-3 text-sm">
          <span className="text-stone-500">{label}</span>
          <span className="text-stone-900 font-medium text-right">{value || '—'}</span>
        </div>
      ))}
    </div>
  </div>
);

// ============================================================================
// HISTORIAL
// ============================================================================

function HistoryView({ deliveries, onView, onDeleteMultiple }) {
  const [filters, setFilters] = useState({
    q: '', fechaIni: '', fechaFin: '', area: '', empleado: '', numero: '', epp: '', motivo: '', entrega: ''
  });
  const [showFilters, setShowFilters] = useState(false);
  const [selected, setSelected] = useState([]);
  const [confirmDeleteMultiple, setConfirmDeleteMultiple] = useState(false);

  const filtered = useMemo(() => {
    return deliveries.filter(d => {
      const q = filters.q.toLowerCase();
      if (q && !`${d.folio} ${d.empleado_nombre} ${d.empleado_numero} ${d.area} ${d.puesto}`.toLowerCase().includes(q)) return false;
      if (filters.fechaIni && d.fecha_entrega < filters.fechaIni) return false;
      if (filters.fechaFin && d.fecha_entrega > filters.fechaFin) return false;
      if (filters.area && !d.area.toLowerCase().includes(filters.area.toLowerCase())) return false;
      if (filters.empleado && !d.empleado_nombre.toLowerCase().includes(filters.empleado.toLowerCase())) return false;
      if (filters.numero && !d.empleado_numero.toLowerCase().includes(filters.numero.toLowerCase())) return false;
      if (filters.epp && !d.epp.includes(filters.epp)) return false;
      if (filters.motivo && d.motivo !== filters.motivo) return false;
      if (filters.entrega && !d.entrega_nombre.toLowerCase().includes(filters.entrega.toLowerCase())) return false;
      return true;
    }).sort((a, b) => b.fecha_registro.localeCompare(a.fecha_registro));
  }, [deliveries, filters]);

  const reset = () => {
    setFilters({ q: '', fechaIni: '', fechaFin: '', area: '', empleado: '', numero: '', epp: '', motivo: '', entrega: '' });
    setSelected([]);
  };

  const activeFilters = Object.values(filters).filter(v => v).length;
  const exportFiltered = () => exportToExcel(filtered, 'entregas_filtradas');

  const toggleSelect = (folio) => {
    setSelected(prev => 
      prev.includes(folio) ? prev.filter(f => f !== folio) : [...prev, folio]
    );
  };

  const toggleSelectAll = () => {
    if (selected.length === filtered.length) {
      setSelected([]);
    } else {
      setSelected(filtered.map(d => d.folio));
    }
  };

  const handleDeleteMultiple = async () => {
    await onDeleteMultiple(selected);
    setSelected([]);
    setConfirmDeleteMultiple(false);
  };

  const isAllSelected = filtered.length > 0 && selected.length === filtered.length;
  const isSomeSelected = selected.length > 0 && selected.length < filtered.length;

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div>
          <h1 className="font-display text-2xl sm:text-3xl font-bold text-stone-900">Historial de entregas</h1>
          <p className="text-sm text-stone-500 mt-0.5">{filtered.length} de {deliveries.length} registros</p>
        </div>
        <div className="flex gap-2">
          <button
            onClick={() => setShowFilters(!showFilters)}
            className={`px-3 py-2 rounded-lg text-sm font-semibold flex items-center gap-1.5 transition ${
              activeFilters > 0 ? 'bg-stone-950 text-yellow-400' : 'bg-white border border-stone-200 text-stone-700 hover:border-stone-400'
            }`}
          >
            <Filter size={14} /> Filtros {activeFilters > 0 && <span className="bg-yellow-400 text-stone-950 px-1.5 rounded-full text-[10px] font-bold">{activeFilters}</span>}
          </button>
          <button
            onClick={exportFiltered}
            disabled={filtered.length === 0}
            className="px-3 py-2 rounded-lg text-sm font-semibold flex items-center gap-1.5 bg-white border border-stone-200 text-stone-700 hover:border-stone-400 transition disabled:opacity-40"
          >
            <FileSpreadsheet size={14} /> Excel
          </button>
        </div>
      </div>

      <div className="relative">
        <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
        <input
          type="text"
          value={filters.q}
          onChange={e => setFilters({ ...filters, q: e.target.value })}
          placeholder="Buscar por folio, nombre, número, área…"
          className="w-full pl-10 pr-4 py-2.5 bg-white border border-stone-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-yellow-400 focus:border-stone-900 transition"
        />
      </div>

      {showFilters && (
        <div className="bg-white border border-stone-200 rounded-xl p-4 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
          <Field label="Desde">
            <Input type="date" value={filters.fechaIni} onChange={e => setFilters({ ...filters, fechaIni: e.target.value })} />
          </Field>
          <Field label="Hasta">
            <Input type="date" value={filters.fechaFin} onChange={e => setFilters({ ...filters, fechaFin: e.target.value })} />
          </Field>
          <Field label="Área o proceso">
            <Input value={filters.area} onChange={e => setFilters({ ...filters, area: e.target.value })} placeholder="Filtrar área" />
          </Field>
          <Field label="Empleado">
            <Input value={filters.empleado} onChange={e => setFilters({ ...filters, empleado: e.target.value })} placeholder="Nombre" />
          </Field>
          <Field label="Número emp.">
            <Input value={filters.numero} onChange={e => setFilters({ ...filters, numero: e.target.value })} placeholder="N° empleado" />
          </Field>
          <Field label="Tipo de EPP">
            <Select value={filters.epp} onChange={e => setFilters({ ...filters, epp: e.target.value })}>
              <option value="">— Todos —</option>
              {EPP_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
            </Select>
          </Field>
          <Field label="Motivo">
            <Select value={filters.motivo} onChange={e => setFilters({ ...filters, motivo: e.target.value })}>
              <option value="">— Todos —</option>
              {REASON_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
            </Select>
          </Field>
          <Field label="Persona que entrega">
            <Input value={filters.entrega} onChange={e => setFilters({ ...filters, entrega: e.target.value })} placeholder="Nombre" />
          </Field>
          <div className="col-span-full flex justify-end">
            <button onClick={reset} className="text-sm text-stone-600 hover:text-stone-900 flex items-center gap-1">
              <RotateCcw size={12} /> Limpiar filtros
            </button>
          </div>
        </div>
      )}

      {filtered.length === 0 ? (
        <div className="bg-white border border-dashed border-stone-300 rounded-2xl p-12 text-center">
          <ClipboardList size={40} className="mx-auto text-stone-300 mb-3" />
          <p className="text-stone-500 font-medium">Sin registros</p>
          <p className="text-stone-400 text-sm">Ajusta los filtros o crea una nueva entrega</p>
        </div>
      ) : (
        <div className="bg-white border border-stone-200 rounded-2xl overflow-hidden">
          <div className="hidden md:block overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-stone-50 border-b border-stone-200">
                <tr>
                  <th className="text-left px-4 py-3 w-12">
                    <label className="flex items-center cursor-pointer">
                      <input
                        type="checkbox"
                        checked={isAllSelected}
                        ref={input => {
                          if (input) input.indeterminate = isSomeSelected;
                        }}
                        onChange={toggleSelectAll}
                        className="w-4 h-4 rounded border-stone-300 text-stone-900 focus:ring-yellow-400 cursor-pointer"
                      />
                    </label>
                  </th>
                  {['Folio', 'Fecha', 'Empleado', 'N°', 'Área', 'EPP', 'Motivo', ''].map(h => (
                    <th key={h} className="text-left px-4 py-3 text-[10px] uppercase tracking-widest font-bold text-stone-600">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-stone-100">
                {filtered.map(d => (
                  <tr 
                    key={d.folio} 
                    className={`hover:bg-stone-50 transition ${selected.includes(d.folio) ? 'bg-yellow-50' : ''}`}
                  >
                    <td className="px-4 py-3">
                      <input
                        type="checkbox"
                        checked={selected.includes(d.folio)}
                        onChange={() => toggleSelect(d.folio)}
                        className="w-4 h-4 rounded border-stone-300 text-stone-900 focus:ring-yellow-400 cursor-pointer"
                        onClick={e => e.stopPropagation()}
                      />
                    </td>
                    <td className="px-4 py-3 font-mono text-xs font-semibold text-stone-900 cursor-pointer" onClick={() => onView(d)}>{d.folio}</td>
                    <td className="px-4 py-3 text-stone-700 cursor-pointer" onClick={() => onView(d)}>{fmtDate(d.fecha_entrega)}</td>
                    <td className="px-4 py-3 text-stone-900 font-medium cursor-pointer" onClick={() => onView(d)}>{d.empleado_nombre}</td>
                    <td className="px-4 py-3 text-stone-500 font-mono text-xs cursor-pointer" onClick={() => onView(d)}>{d.empleado_numero}</td>
                    <td className="px-4 py-3 text-stone-700 cursor-pointer" onClick={() => onView(d)}>{d.area}</td>
                    <td className="px-4 py-3 text-stone-600 cursor-pointer" onClick={() => onView(d)}>
                      <span className="inline-block bg-stone-100 text-stone-700 px-2 py-0.5 rounded text-xs font-medium">{d.epp.length} pieza{d.epp.length !== 1 ? 's' : ''}</span>
                    </td>
                    <td className="px-4 py-3 text-stone-600 text-xs cursor-pointer" onClick={() => onView(d)}>{d.motivo}</td>
                    <td className="px-4 py-3 text-right">
                      <button onClick={() => onView(d)} className="p-1 hover:bg-stone-100 rounded transition">
                        <Eye size={16} className="text-stone-400" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="md:hidden divide-y divide-stone-100">
            {filtered.map(d => (
              <div
                key={d.folio}
                className={`px-4 py-3.5 ${selected.includes(d.folio) ? 'bg-yellow-50' : 'hover:bg-stone-50'} active:bg-stone-100 transition`}
              >
                <div className="flex items-start gap-3">
                  <input
                    type="checkbox"
                    checked={selected.includes(d.folio)}
                    onChange={() => toggleSelect(d.folio)}
                    className="w-4 h-4 rounded border-stone-300 text-stone-900 focus:ring-yellow-400 cursor-pointer mt-1"
                  />
                  <button onClick={() => onView(d)} className="flex-1 text-left">
                    <div className="flex items-start justify-between gap-3 mb-1">
                      <div className="font-mono text-xs font-bold text-stone-900">{d.folio}</div>
                      <div className="text-xs text-stone-500">{fmtDate(d.fecha_entrega)}</div>
                    </div>
                    <div className="font-semibold text-stone-900">{d.empleado_nombre}</div>
                    <div className="text-xs text-stone-500 mt-0.5">{d.area} · {d.empleado_numero}</div>
                    <div className="flex items-center gap-1.5 mt-2 flex-wrap">
                      <span className="bg-stone-100 text-stone-700 px-2 py-0.5 rounded text-[10px] font-semibold">{d.epp.length} EPP</span>
                      <span className="text-[10px] text-stone-500">· {d.motivo}</span>
                    </div>
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {selected.length > 0 && (
        <div className="fixed bottom-4 left-1/2 -translate-x-1/2 z-40 animate-slide-up">
          <div className="bg-stone-950 text-white rounded-2xl shadow-2xl border border-stone-800 p-4 flex items-center gap-4">
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 rounded-full bg-yellow-400 text-stone-950 flex items-center justify-center font-bold text-sm">
                {selected.length}
              </div>
              <span className="text-sm font-semibold">
                {selected.length} {selected.length === 1 ? 'registro seleccionado' : 'registros seleccionados'}
              </span>
            </div>
            
            {confirmDeleteMultiple ? (
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setConfirmDeleteMultiple(false)}
                  className="px-4 py-2 rounded-lg bg-stone-800 hover:bg-stone-700 text-white font-semibold text-sm transition"
                >
                  Cancelar
                </button>
                <button
                  onClick={handleDeleteMultiple}
                  className="px-4 py-2 rounded-lg bg-red-600 hover:bg-red-700 text-white font-semibold text-sm transition flex items-center gap-1.5"
                >
                  <Trash2 size={14} /> Confirmar eliminación
                </button>
              </div>
            ) : (
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setSelected([])}
                  className="px-3 py-2 rounded-lg bg-stone-800 hover:bg-stone-700 text-white font-semibold text-sm transition"
                >
                  Limpiar
                </button>
                <button
                  onClick={() => setConfirmDeleteMultiple(true)}
                  className="px-4 py-2 rounded-lg bg-red-600 hover:bg-red-700 text-white font-semibold text-sm transition flex items-center gap-1.5"
                >
                  <Trash2 size={14} /> Eliminar seleccionados
                </button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ============================================================================
// DASHBOARD
// ============================================================================

const PIE_COLORS = ['#0c0a09', '#facc15', '#78716c', '#d6d3d1', '#f59e0b'];

function Dashboard({ deliveries }) {
  const [filters, setFilters] = useState({ area: '', motivo: '', fechaIni: '', fechaFin: '' });

  const data = useMemo(() => {
    return deliveries.filter(d => {
      if (filters.area && !d.area.toLowerCase().includes(filters.area.toLowerCase())) return false;
      if (filters.motivo && d.motivo !== filters.motivo) return false;
      if (filters.fechaIni && d.fecha_entrega < filters.fechaIni) return false;
      if (filters.fechaFin && d.fecha_entrega > filters.fechaFin) return false;
      return true;
    });
  }, [deliveries, filters]);

  const stats = useMemo(() => {
    const workers = new Set(data.map(d => d.empleado_numero));
    const eppCount = {};
    const reasonCount = {};
    const areaCount = {};
    const monthCount = {};
    data.forEach(d => {
      d.epp.forEach(e => { eppCount[e] = (eppCount[e] || 0) + 1; });
      reasonCount[d.motivo] = (reasonCount[d.motivo] || 0) + 1;
      areaCount[d.area] = (areaCount[d.area] || 0) + 1;
      const mk = monthKey(d.fecha_entrega);
      monthCount[mk] = (monthCount[mk] || 0) + 1;
    });
    const topEpp = Object.entries(eppCount).sort((a, b) => b[1] - a[1])[0];
    const topReason = Object.entries(reasonCount).sort((a, b) => b[1] - a[1])[0];

    return {
      total: data.length,
      workers: workers.size,
      topEpp: topEpp ? topEpp[0] : '—',
      topReason: topReason ? topReason[0] : '—',
      eppData: Object.entries(eppCount).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value })),
      reasonData: Object.entries(reasonCount).map(([name, value]) => ({ name, value })),
      areaData: Object.entries(areaCount).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value })),
      monthData: Object.keys(monthCount).sort().map(k => ({ name: monthLabel(k), value: monthCount[k] }))
    };
  }, [data]);

  const pivot = useMemo(() => {
    const map = {};
    const eppSet = new Set();
    data.forEach(d => {
      const key = `${d.empleado_numero}|${d.empleado_nombre}`;
      if (!map[key]) map[key] = { numero: d.empleado_numero, nombre: d.empleado_nombre, area: d.area, items: {}, total: 0 };
      d.epp.forEach(e => {
        eppSet.add(e);
        map[key].items[e] = (map[key].items[e] || 0) + 1;
        map[key].total += 1;
      });
    });
    return {
      rows: Object.values(map).sort((a, b) => b.total - a.total),
      epps: Array.from(eppSet).sort()
    };
  }, [data]);

  const alerts = useMemo(() => {
    const reposReasons = ['Reposición por daño', 'Reposición por extravío'];
    const counter = {};
    data.forEach(d => {
      if (!reposReasons.includes(d.motivo)) return;
      const key = `${d.empleado_numero}|${d.empleado_nombre}`;
      counter[key] = counter[key] || { numero: d.empleado_numero, nombre: d.empleado_nombre, area: d.area, count: 0, motivos: {} };
      counter[key].count += 1;
      counter[key].motivos[d.motivo] = (counter[key].motivos[d.motivo] || 0) + 1;
    });
    return Object.values(counter)
      .filter(x => x.count >= 2)
      .sort((a, b) => b.count - a.count);
  }, [data]);

  const exportData = () => exportToExcel(data, 'dashboard');

  if (deliveries.length === 0) {
    return <EmptyState />;
  }

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <h1 className="font-display text-2xl sm:text-3xl font-bold text-stone-900">Dashboard</h1>
          <p className="text-sm text-stone-500 mt-0.5">Seguimiento de entrega de EPP</p>
        </div>
        <button
          onClick={exportData}
          className="px-3 py-2 rounded-lg text-sm font-semibold flex items-center gap-1.5 bg-white border border-stone-200 text-stone-700 hover:border-stone-400 transition"
        >
          <FileSpreadsheet size={14} /> Exportar
        </button>
      </div>

      <div className="bg-white border border-stone-200 rounded-xl p-3 grid grid-cols-2 sm:grid-cols-4 gap-2">
        <Input type="date" value={filters.fechaIni} onChange={e => setFilters({ ...filters, fechaIni: e.target.value })} placeholder="Desde" />
        <Input type="date" value={filters.fechaFin} onChange={e => setFilters({ ...filters, fechaFin: e.target.value })} placeholder="Hasta" />
        <Input value={filters.area} onChange={e => setFilters({ ...filters, area: e.target.value })} placeholder="Filtrar área" />
        <Select value={filters.motivo} onChange={e => setFilters({ ...filters, motivo: e.target.value })}>
          <option value="">Todos los motivos</option>
          {REASON_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
        </Select>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <StatCard label="Total entregas" value={stats.total} icon={ClipboardList} accent="bg-yellow-100 text-yellow-700" />
        <StatCard label="Trabajadores" value={stats.workers} icon={User} accent="bg-blue-100 text-blue-700" />
        <StatCard label="EPP más entregado" value={stats.topEpp} icon={Shield} accent="bg-green-100 text-green-700" />
        <StatCard label="Motivo principal" value={stats.topReason} icon={Tag} accent="bg-purple-100 text-purple-700" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <ChartCard title="Entregas por tipo de EPP">
          <ResponsiveContainer width="100%" height={320}>
            <BarChart data={stats.eppData} layout="vertical" margin={{ left: 10, right: 16 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e7e5e4" horizontal={false} />
              <XAxis type="number" stroke="#78716c" fontSize={11} />
              <YAxis type="category" dataKey="name" stroke="#78716c" fontSize={11} width={120} />
              <Tooltip contentStyle={{ backgroundColor: '#0c0a09', border: 'none', borderRadius: 8, color: '#fff' }} cursor={{ fill: '#fafaf9' }} />
              <Bar dataKey="value" fill="#facc15" radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Motivos de entrega">
          <ResponsiveContainer width="100%" height={320}>
            <PieChart>
              <Pie data={stats.reasonData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={100} innerRadius={55} paddingAngle={2}>
                {stats.reasonData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
              </Pie>
              <Tooltip contentStyle={{ backgroundColor: '#0c0a09', border: 'none', borderRadius: 8, color: '#fff' }} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
            </PieChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <ChartCard title="Entregas por mes">
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={stats.monthData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e7e5e4" />
              <XAxis dataKey="name" stroke="#78716c" fontSize={11} />
              <YAxis stroke="#78716c" fontSize={11} allowDecimals={false} />
              <Tooltip contentStyle={{ backgroundColor: '#0c0a09', border: 'none', borderRadius: 8, color: '#fff' }} />
              <Line type="monotone" dataKey="value" stroke="#0c0a09" strokeWidth={2.5} dot={{ fill: '#facc15', r: 5, strokeWidth: 2, stroke: '#0c0a09' }} />
            </LineChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Entregas por área o proceso">
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={stats.areaData.slice(0, 8)}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e7e5e4" />
              <XAxis dataKey="name" stroke="#78716c" fontSize={10} angle={-15} textAnchor="end" height={60} />
              <YAxis stroke="#78716c" fontSize={11} allowDecimals={false} />
              <Tooltip contentStyle={{ backgroundColor: '#0c0a09', border: 'none', borderRadius: 8, color: '#fff' }} cursor={{ fill: '#fafaf9' }} />
              <Bar dataKey="value" fill="#0c0a09" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <ChartCard title="Trabajador × EPP recibido">
        <div className="overflow-x-auto -mx-4 sm:mx-0">
          <table className="w-full text-xs">
            <thead className="bg-stone-50">
              <tr>
                <th className="text-left px-3 py-2 text-[10px] uppercase tracking-widest font-bold text-stone-600 sticky left-0 bg-stone-50">Trabajador</th>
                {pivot.epps.map(e => (
                  <th key={e} className="px-2 py-2 text-[10px] uppercase tracking-wider font-bold text-stone-600 text-center whitespace-nowrap">{e}</th>
                ))}
                <th className="px-3 py-2 text-[10px] uppercase tracking-widest font-bold text-yellow-700 text-center bg-yellow-50">Total</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-stone-100">
              {pivot.rows.slice(0, 30).map(r => (
                <tr key={r.numero} className="hover:bg-stone-50">
                  <td className="px-3 py-2 sticky left-0 bg-white">
                    <div className="font-semibold text-stone-900 whitespace-nowrap">{r.nombre}</div>
                    <div className="text-stone-500 font-mono text-[10px]">{r.numero}</div>
                  </td>
                  {pivot.epps.map(e => (
                    <td key={e} className="px-2 py-2 text-center">
                      {r.items[e] ? (
                        <span className="inline-block min-w-[24px] bg-stone-900 text-yellow-400 rounded font-mono font-bold text-[10px] px-1.5 py-0.5">{r.items[e]}</span>
                      ) : <span className="text-stone-300">·</span>}
                    </td>
                  ))}
                  <td className="px-3 py-2 text-center bg-yellow-50 font-bold text-stone-900">{r.total}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {pivot.rows.length > 30 && (
            <div className="px-3 py-2 text-xs text-stone-500 bg-stone-50 border-t">+ {pivot.rows.length - 30} trabajadores más</div>
          )}
        </div>
      </ChartCard>

      <ChartCard title="Alertas: trabajadores con reposiciones frecuentes" accent="red">
        {alerts.length === 0 ? (
          <div className="py-8 text-center text-stone-500 text-sm flex items-center justify-center gap-2">
            <Check size={16} className="text-green-600" /> Sin alertas — ningún trabajador con reposiciones repetidas
          </div>
        ) : (
          <div className="divide-y divide-stone-100">
            {alerts.map(a => (
              <div key={a.numero} className="py-3 flex items-center justify-between gap-3 flex-wrap">
                <div className="flex items-center gap-3">
                  <div className="w-9 h-9 rounded-full bg-red-100 text-red-700 flex items-center justify-center">
                    <AlertTriangle size={16} />
                  </div>
                  <div>
                    <div className="font-semibold text-stone-900 text-sm">{a.nombre}</div>
                    <div className="text-xs text-stone-500">{a.numero} · {a.area}</div>
                  </div>
                </div>
                <div className="flex items-center gap-2 text-xs">
                  {Object.entries(a.motivos).map(([m, c]) => (
                    <span key={m} className="bg-red-50 text-red-700 px-2 py-1 rounded font-medium">{m}: {c}</span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
      </ChartCard>
    </div>
  );
}

const ChartCard = ({ title, children, accent }) => (
  <div className="bg-white border border-stone-200 rounded-2xl p-4 sm:p-5">
    <h3 className="font-display text-base font-bold text-stone-900 mb-4 flex items-center gap-2">
      <span className={`w-1 h-5 rounded-full ${accent === 'red' ? 'bg-red-500' : 'bg-yellow-400'}`} />
      {title}
    </h3>
    {children}
  </div>
);

// ============================================================================
// EXPORT
// ============================================================================

function exportToExcel(deliveries, filename = 'entregas_epp') {
  if (deliveries.length === 0) {
    alert('No hay datos para exportar');
    return;
  }
  const rows = deliveries.map(d => ({
    Folio: d.folio,
    'Fecha registro': fmtDateTime(d.fecha_registro),
    'Fecha entrega': d.fecha_entrega,
    'Nombre empleado': d.empleado_nombre,
    'N° empleado': d.empleado_numero,
    Puesto: d.puesto,
    'Área/Proceso': d.area,
    'EPP entregado': [...d.epp.filter(e => e !== 'Otro'), d.epp.includes('Otro') ? `Otro: ${d.epp_otro}` : null].filter(Boolean).join(', '),
    Motivo: d.motivo,
    'Texto conformidad': d.conformidad || CONFORMITY_TEXT,
    'Entregado por (nombre)': d.entrega_nombre,
    'Entregado por (cargo)': d.entrega_cargo,
    'Registrado por': d.registrado_por || ''
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0]).map(k => ({ wch: Math.max(k.length + 2, 18) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Entregas EPP');
  XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function ExportView({ deliveries }) {
  return (
    <div className="space-y-5">
      <div>
        <h1 className="font-display text-2xl sm:text-3xl font-bold text-stone-900">Exportar reportes</h1>
        <p className="text-sm text-stone-500 mt-0.5">Descarga la base completa o entrega individual</p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="bg-white border border-stone-200 rounded-2xl p-6">
          <div className="w-10 h-10 rounded-lg bg-green-100 text-green-700 flex items-center justify-center mb-3">
            <FileSpreadsheet size={20} />
          </div>
          <h3 className="font-display font-bold text-lg text-stone-900 mb-1">Excel completo</h3>
          <p className="text-sm text-stone-500 mb-4">Exporta los {deliveries.length} registros con todos los campos.</p>
          <button
            onClick={() => exportToExcel(deliveries, 'entregas_epp_completo')}
            disabled={deliveries.length === 0}
            className="w-full px-4 py-2.5 bg-stone-950 text-white font-semibold rounded-lg hover:bg-stone-800 transition disabled:opacity-40 text-sm flex items-center justify-center gap-2"
          >
            <Download size={14} /> Descargar .xlsx
          </button>
        </div>

        <div className="bg-white border border-stone-200 rounded-2xl p-6">
          <div className="w-10 h-10 rounded-lg bg-blue-100 text-blue-700 flex items-center justify-center mb-3">
            <Printer size={20} />
          </div>
          <h3 className="font-display font-bold text-lg text-stone-900 mb-1">PDF de un registro</h3>
          <p className="text-sm text-stone-500 mb-4">Abre cualquier registro desde el historial y usa "Imprimir / PDF".</p>
          <div className="w-full px-4 py-2.5 bg-stone-100 text-stone-500 font-semibold rounded-lg text-sm text-center">
            Disponible en el detalle
          </div>
        </div>
      </div>

      <div className="bg-yellow-50 border border-yellow-200 rounded-2xl p-5">
        <h3 className="font-display font-bold text-stone-900 mb-1 flex items-center gap-2">
          <AlertTriangle size={16} className="text-yellow-700" /> Migración a base de datos externa
        </h3>
        <p className="text-sm text-stone-700 leading-relaxed">
          Los datos se almacenan en el navegador. Para producción, reemplaza la capa <code className="font-mono text-xs bg-white px-1.5 py-0.5 rounded">db</code> en el código por llamadas a tu backend de Google Sheets, Airtable o SQL.
          La estructura de columnas ya coincide con la propuesta en la arquitectura.
        </p>
      </div>
    </div>
  );
}

// ============================================================================
// HOME
// ============================================================================

function EmptyState() {
  return (
    <div className="bg-white border border-dashed border-stone-300 rounded-2xl p-8 sm:p-12 text-center">
      <HardHat size={48} className="mx-auto text-stone-300 mb-3" />
      <h3 className="font-display text-xl font-bold text-stone-900 mb-1">Aún no hay entregas registradas</h3>
      <p className="text-stone-500 text-sm">Crea la primera entrega desde "Nueva entrega"</p>
    </div>
  );
}

function HomeView({ deliveries, onNavigate }) {
  const recent = deliveries.slice().sort((a, b) => b.fecha_registro.localeCompare(a.fecha_registro)).slice(0, 5);
  const workers = new Set(deliveries.map(d => d.empleado_numero)).size;
  const thisMonth = deliveries.filter(d => monthKey(d.fecha_entrega) === monthKey(new Date().toISOString())).length;

  return (
    <div className="space-y-6">
      <div className="bg-stone-950 text-white rounded-2xl p-6 sm:p-8 relative overflow-hidden">
        <div className="absolute top-0 right-0 w-64 h-64 bg-yellow-400 opacity-10 rounded-full -mr-20 -mt-20" />
        <div className="absolute bottom-0 left-0 w-48 h-48 bg-yellow-400 opacity-5 rounded-full -ml-16 -mb-16" />
        <div className="relative">
          <div className="flex items-center gap-2 mb-3">
            <HardHat size={20} className="text-yellow-400" />
            <span className="text-[10px] uppercase tracking-widest font-bold text-yellow-400">Seguridad y Salud en el Trabajo</span>
          </div>
          <h1 className="font-display text-3xl sm:text-5xl font-bold leading-tight mb-2 text-white">Control de entrega<br />de EPP</h1>
          <p className="text-stone-300 max-w-md text-sm sm:text-base">Registra, almacena y da seguimiento a la entrega de equipo de protección personal con firma digital y evidencia auditable.</p>
          <button
            onClick={() => onNavigate('new')}
            className="mt-5 inline-flex items-center gap-2 bg-yellow-400 text-stone-950 px-5 py-3 rounded-lg font-bold text-sm hover:bg-yellow-300 transition"
          >
            <Plus size={16} /> Nueva entrega de EPP
          </button>
        </div>
      </div>

      <div className="grid grid-cols-3 gap-3">
        <StatCard label="Total" value={deliveries.length} icon={ClipboardList} accent="bg-yellow-100 text-yellow-700" />
        <StatCard label="Trabajadores" value={workers} icon={User} accent="bg-blue-100 text-blue-700" />
        <StatCard label="Este mes" value={thisMonth} icon={Calendar} accent="bg-green-100 text-green-700" />
      </div>

      <div className="bg-white border border-stone-200 rounded-2xl p-5 sm:p-6">
        <div className="flex items-center justify-between mb-4">
          <h2 className="font-display text-lg font-bold text-stone-900 flex items-center gap-2">
            <span className="w-1 h-5 bg-yellow-400 rounded-full" />
            Entregas recientes
          </h2>
          {deliveries.length > 5 && (
            <button onClick={() => onNavigate('history')} className="text-xs font-semibold text-stone-700 hover:text-stone-950 uppercase tracking-wider">Ver todas →</button>
          )}
        </div>
        {recent.length === 0 ? (
          <EmptyState />
        ) : (
          <div className="divide-y divide-stone-100 -mx-2">
            {recent.map(d => (
              <button
                key={d.folio}
                onClick={() => onNavigate('history')}
                className="w-full px-2 py-3 hover:bg-stone-50 text-left transition flex items-center gap-3"
              >
                <div className="w-10 h-10 rounded-lg bg-stone-100 flex items-center justify-center font-mono text-[10px] font-bold text-stone-700 flex-shrink-0">
                  {d.folio.split('-')[2]}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="font-semibold text-stone-900 text-sm truncate">{d.empleado_nombre}</div>
                  <div className="text-xs text-stone-500 truncate">{d.area} · {d.epp.length} EPP · {fmtDate(d.fecha_entrega)}</div>
                </div>
                <Eye size={14} className="text-stone-400 flex-shrink-0" />
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ============================================================================
// APP PRINCIPAL
// ============================================================================

export default function App() {
  const [tab, setTab] = useState('home');
  const [deliveries, setDeliveries] = useState([]);
  const [loading, setLoading] = useState(true);
  const [detail, setDetail] = useState(null);

  const refresh = async () => {
    const data = await db.getIndex();
    setDeliveries(data);
  };

  useEffect(() => {
    (async () => {
      await refresh();
      setLoading(false);
    })();
  }, []);

  const handleSaveSuccess = async (folio) => {
    await refresh();
    setTab('history');
  };

  const handleDelete = async (folio) => {
    await db.deleteDelivery(folio);
    setDetail(null);
    await refresh();
  };

  const handleDeleteMultiple = async (folios) => {
    if (folios.length === 0) return;
    await db.deleteDelivery(folios);
    await refresh();
  };

  const tabs = [
    { id: 'home', label: 'Inicio', icon: Home },
    { id: 'new', label: 'Nueva entrega', icon: Plus },
    { id: 'history', label: 'Historial', icon: ClipboardList },
    { id: 'dashboard', label: 'Dashboard', icon: BarChart3 },
    { id: 'export', label: 'Exportar', icon: Download }
  ];

  return (
    <>

      <div className="min-h-screen bg-stone-100 font-body text-stone-900">
        <header className="bg-stone-950 text-white sticky top-0 z-30 no-print">
          <div className="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between gap-4">
            <div className="flex items-center gap-2.5">
              <div className="w-8 h-8 rounded-lg bg-yellow-400 text-stone-950 flex items-center justify-center">
                <HardHat size={18} strokeWidth={2.5} />
              </div>
              <div>
                <div className="font-display font-bold text-base leading-tight">EPP Control</div>
                <div className="text-[9px] uppercase tracking-widest text-stone-400 leading-tight">Seguridad &amp; Salud en el Trabajo</div>
              </div>
            </div>
            <div className="hidden sm:flex items-center gap-1">
              {tabs.map(t => (
                <button
                  key={t.id}
                  onClick={() => setTab(t.id)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 ${
                    tab === t.id ? 'bg-yellow-400 text-stone-950' : 'text-stone-300 hover:text-white hover:bg-stone-800'
                  }`}
                >
                  <t.icon size={13} /> {t.label}
                </button>
              ))}
            </div>
          </div>
          <div className="sm:hidden overflow-x-auto border-t border-stone-800">
            <div className="flex px-2 py-2 gap-1 min-w-max">
              {tabs.map(t => (
                <button
                  key={t.id}
                  onClick={() => setTab(t.id)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition flex items-center gap-1.5 whitespace-nowrap ${
                    tab === t.id ? 'bg-yellow-400 text-stone-950' : 'text-stone-300 hover:text-white hover:bg-stone-800'
                  }`}
                >
                  <t.icon size={13} /> {t.label}
                </button>
              ))}
            </div>
          </div>
        </header>

        <main className="max-w-6xl mx-auto px-3 sm:px-4 py-5 sm:py-8 pb-24">
          {loading ? (
            <div className="flex items-center justify-center py-20 text-stone-500">
              <Loader2 className="animate-spin mr-2" size={20} /> Cargando…
            </div>
          ) : (
            <>
              {tab === 'home' && <HomeView deliveries={deliveries} onNavigate={setTab} />}
              {tab === 'new' && <DeliveryForm onSuccess={handleSaveSuccess} onCancel={() => setTab('home')} />}
              {tab === 'history' && <HistoryView deliveries={deliveries} onView={setDetail} onDeleteMultiple={handleDeleteMultiple} />}
              {tab === 'dashboard' && <Dashboard deliveries={deliveries} />}
              {tab === 'export' && <ExportView deliveries={deliveries} />}
            </>
          )}
        </main>

        <footer className="border-t border-stone-200 bg-white no-print">
          <div className="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between text-[10px] uppercase tracking-widest text-stone-500">
            <span className="font-mono">EPP-CONTROL · v1.0</span>
            <span>{deliveries.length} {deliveries.length === 1 ? 'registro' : 'registros'}</span>
          </div>
        </footer>

        {detail && <DeliveryDetail delivery={detail} onClose={() => setDetail(null)} onDelete={handleDelete} />}
      </div>
    </>
  );
}
