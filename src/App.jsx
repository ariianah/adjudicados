import { storage } from "./lib/storage";
import React, { useState, useEffect, useCallback, useRef } from "react";
import { Phone, Check, X, Plus, ArrowRight, FileText, Clock, Tag, Trash2, RefreshCw, Upload } from "lucide-react";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------
   Panel de Adjudicaciones
   Paleta: petróleo profundo + ámbar (guiño al guiño del auto)
   Tipografía: Space Grotesk (display) / Inter (cuerpo) / IBM Plex Mono (datos)
--------------------------------------------------------- */

const COLORS = {
  bg: "#0F1720",
  panel: "#18222D",
  panelAlt: "#1E2A37",
  border: "#2B3948",
  text: "#EDEFF2",
  textMuted: "#8A97A6",
  amber: "#E8A23D",
  amberSoft: "#3A2E1A",
  green: "#4F9D69",
  greenSoft: "#16261C",
  red: "#C1493F",
  redSoft: "#2C1917",
};

const DOC_ITEMS = ["DNI", "Recibo de sueldo", "Constancia CUIL", "Comprobante de domicilio"];
const DEFAULT_PLANES = [
  "Cronos 80/20",
  "Cronos 90/10",
  "Cronos 70/30",
  "Mobi 80/20",
  "Fastback 70/30",
  "Fiorino 70/30",
  "Pulse 60/40",
  "Titano 70/30",
  "Titano 60/40",
];

const baseModelo = (planLabel) => (planLabel || "").split(" ")[0];

function parseMonto(valor) {
  if (!valor) return null;
  const n = Number(String(valor).replace(/[^\d.,-]/g, "").replace(/\./g, "").replace(",", "."));
  return Number.isFinite(n) && n > 0 ? n : null;
}

function calcularPorcentaje(montoAdjudica, valorUnidad) {
  const monto = parseMonto(montoAdjudica);
  const valor = parseMonto(valorUnidad);
  if (!monto || !valor) return null;
  return (monto / valor) * 100;
}

const uid = () => Math.random().toString(36).slice(2, 10);

function docInfo(c, doc, campo = "docs") {
  const v = c[campo]?.[doc];
  if (v && typeof v === "object") return { ok: !!v.ok, archivo: v.archivo || null, nombreArchivo: v.nombreArchivo || "" };
  return { ok: !!v, archivo: null, nombreArchivo: "" };
}

function comprimirImagen(file, callback) {
  if (!file.type.startsWith("image/")) {
    // No es imagen (ej PDF): guardamos solo el nombre, sin previsualización
    callback({ archivo: null, nombreArchivo: file.name });
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    const img = new Image();
    img.onload = () => {
      const maxW = 700;
      const scale = Math.min(1, maxW / img.width);
      const canvas = document.createElement("canvas");
      canvas.width = img.width * scale;
      canvas.height = img.height * scale;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      callback({ archivo: canvas.toDataURL("image/jpeg", 0.6), nombreArchivo: file.name });
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(dateStr + "T00:00:00");
  return Math.round((target - today) / 86400000);
}

function paymentStatus(cliente) {
  if (cliente.pagado) return "pagado";
  const d = daysUntil(cliente.vencimiento);
  if (d === null) return "sin-fecha";
  if (d < 0) return "vencido";
  if (d <= 5) return "proximo";
  return "al-dia";
}

const STATUS_STYLE = {
  pagado: { label: "Pagado", color: COLORS.green, bg: COLORS.greenSoft },
  "al-dia": { label: "Al día", color: COLORS.green, bg: COLORS.greenSoft },
  proximo: { label: "Vence pronto", color: COLORS.amber, bg: COLORS.amberSoft },
  vencido: { label: "Vencido", color: COLORS.red, bg: COLORS.redSoft },
  "sin-fecha": { label: "Sin fecha", color: COLORS.textMuted, bg: COLORS.panelAlt },
};

function quitarAcentos(txt) {
  return (txt || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function detectarColumna(headers, candidatos) {
  const norm = headers.map((h) => quitarAcentos(String(h)).toLowerCase().trim());
  for (const cand of candidatos) {
    const idx = norm.findIndex((h) => h.includes(cand));
    if (idx !== -1) return headers[idx];
  }
  return null;
}

function mapearFilaImportada(fila, headers) {
  const colNombre = detectarColumna(headers, ["nombre", "cliente", "apellido"]);
  const colTelefono = detectarColumna(headers, ["telefono", "whatsapp", "celular", "numero"]);
  const colPlan = detectarColumna(headers, ["plan", "modelo"]);
  const colGrupo = detectarColumna(headers, ["grupo"]);
  const colOrden = detectarColumna(headers, ["orden"]);
  const colMonto = detectarColumna(headers, ["monto", "adjudica", "importe"]);
  return {
    nombre: colNombre ? String(fila[colNombre] ?? "").trim() : "",
    telefono: colTelefono ? String(fila[colTelefono] ?? "").trim() : "",
    plan: colPlan ? String(fila[colPlan] ?? "").trim() : "",
    grupo: colGrupo ? String(fila[colGrupo] ?? "").trim() : "",
    orden: colOrden ? String(fila[colOrden] ?? "").trim() : "",
    montoAdjudica: colMonto ? String(fila[colMonto] ?? "").trim() : "",
  };
}

function normalizarTelefono(numero) {
  let clean = (numero || "").replace(/[^\d]/g, "");
  if (!clean) return "";
  if (clean.startsWith("54")) return clean;
  if (clean.startsWith("9")) return "54" + clean;
  if (clean.startsWith("0")) clean = clean.slice(1);
  return "549" + clean;
}

function waLink(telefono, mensaje) {
  const clean = normalizarTelefono(telefono);
  return `https://wa.me/${clean}?text=${encodeURIComponent(mensaje)}`;
}

function reminderMessage(cliente) {
  const st = paymentStatus(cliente);
  if (st === "vencido") {
    return `Hola ${cliente.nombre}! Te escribo de Fiat Fadua: tu cuota de adjudicación venció el ${cliente.vencimiento}. ¿Podemos coordinar el pago a la brevedad?`;
  }
  return `Hola ${cliente.nombre}! Te recuerdo que tu cuota de adjudicación vence el ${cliente.vencimiento}. Cualquier consulta, quedo atenta.`;
}

/* -------------------- Bloques chicos de UI -------------------- */

function Pill({ children, color, bg }) {
  return (
    <span
      style={{
        color,
        background: bg,
        border: `1px solid ${color}55`,
        borderRadius: 999,
        padding: "3px 10px",
        fontSize: 12,
        fontWeight: 600,
        letterSpacing: 0.2,
        whiteSpace: "nowrap",
      }}
    >
      {children}
    </span>
  );
}

function IconBtn({ onClick, title, children, danger }) {
  return (
    <button
      onClick={onClick}
      title={title}
      style={{
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        width: 30,
        height: 30,
        borderRadius: 8,
        border: `1px solid ${COLORS.border}`,
        background: "transparent",
        color: danger ? COLORS.red : COLORS.textMuted,
        cursor: "pointer",
      }}
      onMouseEnter={(e) => (e.currentTarget.style.borderColor = danger ? COLORS.red : COLORS.amber)}
      onMouseLeave={(e) => (e.currentTarget.style.borderColor = COLORS.border)}
    >
      {children}
    </button>
  );
}

function ConfigPanel({ planes, onAgregarPlan, onEliminarPlan, valoresUnidades, onActualizarValor }) {
  const [nuevoPlan, setNuevoPlan] = useState("");
  const modelos = [...new Set(planes.map(baseModelo))];

  return (
    <div
      style={{
        background: COLORS.panel,
        border: `1px solid ${COLORS.border}`,
        borderRadius: 10,
        padding: 16,
        marginBottom: 18,
      }}
    >
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 20 }}>
        <div>
          <div style={{ fontSize: 12, color: COLORS.textMuted, marginBottom: 8, fontWeight: 600 }}>Planes disponibles</div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
            {planes.map((p) => (
              <span
                key={p}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 6,
                  background: COLORS.panelAlt,
                  border: `1px solid ${COLORS.border}`,
                  borderRadius: 999,
                  padding: "4px 4px 4px 10px",
                  fontSize: 12,
                  fontFamily: "'IBM Plex Mono', monospace",
                }}
              >
                {p}
                <button
                  onClick={() => onEliminarPlan(p)}
                  style={{ background: "none", border: "none", color: COLORS.textMuted, cursor: "pointer", display: "flex", padding: 2 }}
                  title="Quitar plan"
                >
                  <X size={12} />
                </button>
              </span>
            ))}
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <TextField value={nuevoPlan} onChange={setNuevoPlan} placeholder="Ej: Argo 70/30" width={160} />
            <IconBtn
              title="Agregar plan"
              onClick={() => {
                if (nuevoPlan.trim()) {
                  onAgregarPlan(nuevoPlan.trim());
                  setNuevoPlan("");
                }
              }}
            >
              <Plus size={14} />
            </IconBtn>
          </div>
        </div>

        <div>
          <div style={{ fontSize: 12, color: COLORS.textMuted, marginBottom: 8, fontWeight: 600 }}>
            Valor real de la unidad ($) — actualizalo cada mes
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {modelos.map((m) => (
              <div key={m} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 12, width: 90, color: COLORS.textMuted }}>{m}</span>
                <TextField
                  value={valoresUnidades[m] || ""}
                  onChange={(v) => onActualizarValor(m, v)}
                  placeholder="$ valor actual"
                  mono
                  width={160}
                />
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

function SelectField({ value, onChange, options, placeholder, width }) {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        background: COLORS.bg,
        border: `1px solid ${COLORS.border}`,
        borderRadius: 6,
        color: value ? COLORS.text : COLORS.textMuted,
        fontFamily: "'IBM Plex Mono', monospace",
        fontSize: 13,
        padding: "6px 8px",
        width: width || "100%",
        outline: "none",
      }}
    >
      <option value="">{placeholder}</option>
      {options.map((o) => (
        <option key={o} value={o} style={{ color: COLORS.text, background: COLORS.panel }}>
          {o}
        </option>
      ))}
    </select>
  );
}

function TextField({ value, onChange, placeholder, mono, width }) {
  return (
    <input
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      style={{
        background: COLORS.bg,
        border: `1px solid ${COLORS.border}`,
        borderRadius: 6,
        color: COLORS.text,
        fontFamily: mono ? "'IBM Plex Mono', monospace" : "inherit",
        fontSize: 13,
        padding: "6px 8px",
        width: width || "100%",
        outline: "none",
      }}
    />
  );
}

/* -------------------- App -------------------- */

export default function App() {
  const [tab, setTab] = useState("invitados");
  const [invitados, setInvitados] = useState([]);
  const [adjudicados, setAdjudicados] = useState([]);
  const [planes, setPlanes] = useState(DEFAULT_PLANES);
  const [valoresUnidades, setValoresUnidades] = useState({});
  const [mostrarConfig, setMostrarConfig] = useState(false);
  const [loading, setLoading] = useState(true);
  const [saveError, setSaveError] = useState(false);

  // Formularios de alta rápida
  const [nuevoInvitado, setNuevoInvitado] = useState({ nombre: "", telefono: "", grupo: "", orden: "", plan: "", montoAdjudica: "", prorrateo: false, subite: false, subiteCobertura: "" });
  const [nuevoAdj, setNuevoAdj] = useState({ nombre: "", telefono: "", plan: "", montoAdjudica: "" });

  const persist = useCallback(async (key, value) => {
    try {
      const res = await storage.set(key, JSON.stringify(value));
      setSaveError(!res);
    } catch {
      setSaveError(true);
    }
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const [inv, adj, pl, vu] = await Promise.all([
          storage.get("invitados").catch(() => null),
          storage.get("adjudicados").catch(() => null),
          storage.get("planes").catch(() => null),
          storage.get("valoresUnidades").catch(() => null),
        ]);
        setInvitados(inv ? JSON.parse(inv.value) : []);
        setAdjudicados(adj ? JSON.parse(adj.value) : []);
        setPlanes(pl ? JSON.parse(pl.value) : DEFAULT_PLANES);
        setValoresUnidades(vu ? JSON.parse(vu.value) : {});
      } catch {
        setSaveError(true);
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  
 /* const persist = useCallback((key, value) => {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      setSaveError(false);
    } catch {
      setSaveError(true);
    }
  }, []);

  useEffect(() => {
    try {
      const inv = localStorage.getItem("invitados");
      const adj = localStorage.getItem("adjudicados");
      const pl = localStorage.getItem("planes");
      const vu = localStorage.getItem("valoresUnidades");
      setInvitados(inv ? JSON.parse(inv) : []);
      setAdjudicados(adj ? JSON.parse(adj) : []);
      setPlanes(pl ? JSON.parse(pl) : DEFAULT_PLANES);
      setValoresUnidades(vu ? JSON.parse(vu) : {});
    } catch {
      setSaveError(true);
    } finally {
      setLoading(false);
    }
  }, []);
  */

  useEffect(() => {
    if (!loading) persist("invitados", invitados);
  }, [invitados, loading, persist]);

  useEffect(() => {
    if (!loading) persist("adjudicados", adjudicados);
  }, [adjudicados, loading, persist]);

  useEffect(() => {
    if (!loading) persist("planes", planes);
  }, [planes, loading, persist]);

  useEffect(() => {
    if (!loading) persist("valoresUnidades", valoresUnidades);
  }, [valoresUnidades, loading, persist]);

  const agregarPlan = (plan) => setPlanes((prev) => (prev.includes(plan) ? prev : [...prev, plan]));
  const eliminarPlan = (plan) => setPlanes((prev) => prev.filter((p) => p !== plan));
  const actualizarValorUnidad = (modelo, valor) => setValoresUnidades((prev) => ({ ...prev, [modelo]: valor }));

  /* ---- acciones invitados ---- */
  const agregarInvitado = () => {
    if (!nuevoInvitado.nombre.trim()) return;
    setInvitados((prev) => [
      ...prev,
      {
        id: uid(),
        nombre: nuevoInvitado.nombre.trim(),
        telefono: nuevoInvitado.telefono.trim(),
        grupo: nuevoInvitado.grupo.trim(),
        orden: nuevoInvitado.orden.trim(),
        plan: nuevoInvitado.plan,
        montoAdjudica: nuevoInvitado.montoAdjudica,
        estado: "invitado",
        prorrateo: nuevoInvitado.prorrateo,
        subite: nuevoInvitado.subite,
        subiteCobertura: nuevoInvitado.subiteCobertura,
      },
    ]);
    setNuevoInvitado({ nombre: "", telefono: "", grupo: "", orden: "", plan: "", montoAdjudica: "", prorrateo: false, subite: false, subiteCobertura: "" });
  };

  const marcarConfirmado = (id) => {
    setInvitados((prev) => prev.map((p) => (p.id === id ? { ...p, estado: "confirmado" } : p)));
  };

  const pasarAAdjudicados = (id) => {
    const persona = invitados.find((p) => p.id === id);
    if (!persona) return;
    setAdjudicados((prev) => [
      ...prev,
      {
        id: uid(),
        nombre: persona.nombre,
        telefono: persona.telefono,
        grupo: persona.grupo || "",
        orden: persona.orden || "",
        plan: persona.plan || "",
        montoAdjudica: persona.montoAdjudica || "",
        prorrateo: persona.prorrateo || false,
        subite: persona.subite || false,
        subiteCobertura: persona.subiteCobertura || "",
        modelo: "",
        color: "",
        monto: "",
        vencimiento: "",
        pagado: false,
        pedidoRealizado: false,
        observaciones: "",
        veraz: false,
        necesitaGarante: false,
        garanteNombre: "",
        docs: Object.fromEntries(DOC_ITEMS.map((d) => [d, { ok: false, archivo: null, nombreArchivo: "" }])),
        docsGarante: Object.fromEntries(DOC_ITEMS.map((d) => [d, { ok: false, archivo: null, nombreArchivo: "" }])),
        carpetaAprobada: false,
      },
    ]);
    setInvitados((prev) => prev.filter((p) => p.id !== id));
    setTab("adjudicados");
  };

  const eliminarInvitado = (id) => setInvitados((prev) => prev.filter((p) => p.id !== id));
  const actualizarInvitado = (id, campo, valor) =>
    setInvitados((prev) => prev.map((p) => (p.id === id ? { ...p, [campo]: valor } : p)));

  const [importInfo, setImportInfo] = useState(null);
  const importarInvitados = (filas) => {
    const nuevos = filas
      .filter((f) => f.nombre)
      .map((f) => ({ id: uid(), estado: "invitado", ...f }));
    setInvitados((prev) => [...prev, ...nuevos]);
    setImportInfo({ agregados: nuevos.length, descartados: filas.length - nuevos.length });
  };

  /* ---- acciones adjudicados ---- */
  const agregarAdjudicadoDirecto = () => {
    if (!nuevoAdj.nombre.trim()) return;
    setAdjudicados((prev) => [
      ...prev,
      {
        id: uid(),
        nombre: nuevoAdj.nombre.trim(),
        telefono: nuevoAdj.telefono.trim(),
        plan: nuevoAdj.plan,
        montoAdjudica: nuevoAdj.montoAdjudica,
        modelo: "",
        color: "",
        monto: "",
        vencimiento: "",
        pagado: false,
        pedidoRealizado: false,
        observaciones: "",
        veraz: false,
        necesitaGarante: false,
        garanteNombre: "",
        docs: Object.fromEntries(DOC_ITEMS.map((d) => [d, { ok: false, archivo: null, nombreArchivo: "" }])),
        docsGarante: Object.fromEntries(DOC_ITEMS.map((d) => [d, { ok: false, archivo: null, nombreArchivo: "" }])),
        carpetaAprobada: false,
      },
    ]);
    setNuevoAdj({ nombre: "", telefono: "", plan: "", montoAdjudica: "" });
  };

  const actualizarAdj = (id, campo, valor) => {
    setAdjudicados((prev) => prev.map((c) => (c.id === id ? { ...c, [campo]: valor } : c)));
  };

  const toggleDoc = (id, doc, campo = "docs") => {
    setAdjudicados((prev) =>
      prev.map((c) => {
        if (c.id !== id) return c;
        const actual = docInfo(c, doc, campo);
        return { ...c, [campo]: { ...c[campo], [doc]: { ...actual, ok: !actual.ok } } };
      })
    );
  };

  const adjuntarDoc = (id, doc, { archivo, nombreArchivo }, campo = "docs") => {
    setAdjudicados((prev) =>
      prev.map((c) => (c.id === id ? { ...c, [campo]: { ...c[campo], [doc]: { ok: true, archivo, nombreArchivo } } } : c))
    );
  };

  const quitarDocArchivo = (id, doc, campo = "docs") => {
    setAdjudicados((prev) =>
      prev.map((c) => (c.id === id ? { ...c, [campo]: { ...c[campo], [doc]: { ok: false, archivo: null, nombreArchivo: "" } } } : c))
    );
  };

  const eliminarAdj = (id) => setAdjudicados((prev) => prev.filter((c) => c.id !== id));

  const exportarAprobados = () => {
    const aprobados = adjudicados.filter((c) => c.carpetaAprobada);
    if (aprobados.length === 0) return;
    const filas = aprobados.map((c) => {
      const pct = calcularPorcentaje(c.montoAdjudica, valoresUnidades[baseModelo(c.plan)]);
      return {
        Nombre: c.nombre,
        WhatsApp: c.telefono,
        Grupo: c.grupo || "",
        Orden: c.orden || "",
        Plan: c.plan || "",
        "Adjudica con $": c.montoAdjudica || "",
        "% real": pct !== null ? pct.toFixed(1) : "",
        Modelo: c.modelo || "",
        Color: c.color || "",
        "Monto de cuota": c.monto || "",
        Vencimiento: c.vencimiento || "",
        Pagado: c.pagado ? "Sí" : "No",
        "Pedido en SGA": c.pedidoRealizado ? "Sí" : "No",
        "Afectado en Veraz": c.veraz ? "Sí" : "No",
        "Necesita garante": c.necesitaGarante ? "Sí" : "No",
        Garante: c.garanteNombre || "",
        Observaciones: c.observaciones || "",
      };
    });
    const ws = XLSX.utils.json_to_sheet(filas);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Aprobados");
    XLSX.writeFile(wb, `adjudicados-aprobados-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const contadores = {
    invitados: invitados.filter((i) => i.estado === "invitado").length,
    confirmados: invitados.filter((i) => i.estado === "confirmado").length,
    vencidos: adjudicados.filter((c) => paymentStatus(c) === "vencido").length,
    proximos: adjudicados.filter((c) => paymentStatus(c) === "proximo").length,
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        background: COLORS.bg,
        color: COLORS.text,
        fontFamily: "'Inter', system-ui, sans-serif",
        padding: "28px 18px 60px",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500&display=swap');
        * { box-sizing: border-box; }
        ::placeholder { color: ${COLORS.textMuted}; }
        table { border-collapse: collapse; width: 100%; }
        @media (prefers-reduced-motion: reduce) { * { transition: none !important; } }
      `}</style>

      <div style={{ maxWidth: 1000, margin: "0 auto" }}>
        {/* Encabezado */}
        <div style={{ display: "flex", alignItems: "baseline", gap: 10, marginBottom: 4 }}>
          <span style={{ width: 10, height: 10, borderRadius: "50%", background: COLORS.amber, boxShadow: `0 0 12px ${COLORS.amber}` }} />
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 26, margin: 0 }}>
            Panel de Adjudicaciones
          </h1>
        </div>
        <p style={{ color: COLORS.textMuted, margin: "0 0 22px", fontSize: 14 }}>
          Del guiño de invitación a la entrega de la unidad — dos etapas, un solo lugar.
        </p>

        {/* Contadores tipo tablero */}
        <div style={{ display: "flex", gap: 10, marginBottom: 22, flexWrap: "wrap" }}>
          {[
            ["Invitados sin confirmar", contadores.invitados, COLORS.textMuted],
            ["Confirmados por pasar", contadores.confirmados, COLORS.amber],
            ["Cuotas por vencer", contadores.proximos, COLORS.amber],
            ["Cuotas vencidas", contadores.vencidos, COLORS.red],
          ].map(([label, val, color]) => (
            <div
              key={label}
              style={{
                flex: "1 1 200px",
                background: COLORS.panel,
                border: `1px solid ${COLORS.border}`,
                borderLeft: `3px solid ${color}`,
                borderRadius: 8,
                padding: "10px 14px",
              }}
            >
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 22, color }}>{val}</div>
              <div style={{ fontSize: 12, color: COLORS.textMuted }}>{label}</div>
            </div>
          ))}
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", gap: 6, marginBottom: 18, borderBottom: `1px solid ${COLORS.border}`, alignItems: "center" }}>
          {[
            ["invitados", "1 · Invitados"],
            ["adjudicados", "2 · Adjudicados"],
          ].map(([key, label]) => (
            <button
              key={key}
              onClick={() => setTab(key)}
              style={{
                background: "none",
                border: "none",
                borderBottom: tab === key ? `2px solid ${COLORS.amber}` : "2px solid transparent",
                color: tab === key ? COLORS.text : COLORS.textMuted,
                fontFamily: "'Space Grotesk', sans-serif",
                fontWeight: 600,
                fontSize: 14,
                padding: "8px 4px",
                cursor: "pointer",
              }}
            >
              {label}
            </button>
          ))}
          <button
            onClick={() => setMostrarConfig((v) => !v)}
            style={{
              marginLeft: "auto",
              background: "none",
              border: "none",
              color: mostrarConfig ? COLORS.amber : COLORS.textMuted,
              fontSize: 12,
              fontWeight: 600,
              cursor: "pointer",
              padding: "8px 4px",
            }}
          >
            ⚙ Planes y valores
          </button>
        </div>

        {mostrarConfig && (
          <ConfigPanel
            planes={planes}
            onAgregarPlan={agregarPlan}
            onEliminarPlan={eliminarPlan}
            valoresUnidades={valoresUnidades}
            onActualizarValor={actualizarValorUnidad}
          />
        )}

        {loading ? (
          <div style={{ color: COLORS.textMuted, fontSize: 13 }}>Cargando…</div>
        ) : tab === "invitados" ? (
          <InvitadosTab
            invitados={invitados}
            planes={planes}
            valoresUnidades={valoresUnidades}
            nuevo={nuevoInvitado}
            setNuevo={setNuevoInvitado}
            onAgregar={agregarInvitado}
            onConfirmar={marcarConfirmado}
            onPasar={pasarAAdjudicados}
            onEliminar={eliminarInvitado}
            onActualizar={actualizarInvitado}
            onImportar={importarInvitados}
            importInfo={importInfo}
          />
        ) : (
          <AdjudicadosTab
            adjudicados={adjudicados}
            planes={planes}
            valoresUnidades={valoresUnidades}
            nuevo={nuevoAdj}
            setNuevo={setNuevoAdj}
            onAgregar={agregarAdjudicadoDirecto}
            onActualizar={actualizarAdj}
            onToggleDoc={toggleDoc}
            onAdjuntarDoc={adjuntarDoc}
            onQuitarDocArchivo={quitarDocArchivo}
            onEliminar={eliminarAdj}
            onExportar={exportarAprobados}
          />
        )}

        {saveError && (
          <div style={{ marginTop: 18, fontSize: 12, color: COLORS.red, display: "flex", alignItems: "center", gap: 6 }}>
            <RefreshCw size={13} /> No se pudo guardar el último cambio. Los datos quedan solo en esta sesión.
          </div>
        )}
      </div>
    </div>
  );
}

/* -------------------- Tab: Invitados -------------------- */

function InvitadosTab({ invitados, planes, valoresUnidades, nuevo, setNuevo, onAgregar, onConfirmar, onPasar, onEliminar, onActualizar, onImportar, importInfo }) {
  const [busqueda, setBusqueda] = useState("");
  const [expandido, setExpandido] = useState(null);
  const fileRef = useRef(null);

  const manejarArchivo = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const hoja = wb.Sheets[wb.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(hoja, { defval: "" });
        if (!filas.length) return;
        const headers = Object.keys(filas[0]);
        const mapeadas = filas.map((f) => mapearFilaImportada(f, headers));
        onImportar(mapeadas);
      } catch {
        onImportar([]);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const filtrados = invitados.filter((p) => {
    if (!busqueda.trim()) return true;
    const q = busqueda.trim().toLowerCase();
    return (
      (p.grupo || "").toLowerCase().includes(q) ||
      (p.orden || "").toLowerCase().includes(q) ||
      (p.nombre || "").toLowerCase().includes(q)
    );
  });

  return (
    <div>
      <div
        style={{
          background: COLORS.panel,
          border: `1px solid ${COLORS.border}`,
          borderRadius: 10,
          padding: 14,
          marginBottom: 16,
          display: "flex",
          gap: 8,
          flexWrap: "wrap",
          alignItems: "center",
        }}
      >
        <TextField value={nuevo.nombre} onChange={(v) => setNuevo((n) => ({ ...n, nombre: v }))} placeholder="Nombre del invitado" width={200} />
        <TextField value={nuevo.telefono} onChange={(v) => setNuevo((n) => ({ ...n, telefono: v }))} placeholder="WhatsApp (5493...)" mono width={170} />
        <TextField value={nuevo.grupo} onChange={(v) => setNuevo((n) => ({ ...n, grupo: v }))} placeholder="Grupo" mono width={90} />
        <TextField value={nuevo.orden} onChange={(v) => setNuevo((n) => ({ ...n, orden: v }))} placeholder="Orden" mono width={90} />
        <SelectField value={nuevo.plan} onChange={(v) => setNuevo((n) => ({ ...n, plan: v }))} options={planes} placeholder="Plan" width={140} />
        <TextField value={nuevo.montoAdjudica} onChange={(v) => setNuevo((n) => ({ ...n, montoAdjudica: v }))} placeholder="Adjudica con $" mono width={130} />
        <button
          onClick={onAgregar}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            background: COLORS.amber,
            color: "#1A1200",
            border: "none",
            borderRadius: 6,
            padding: "7px 12px",
            fontWeight: 700,
            fontSize: 13,
            cursor: "pointer",
          }}
        >
          <Plus size={14} /> Invitar
        </button>
      </div>

      <div style={{ marginBottom: 12, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
        <TextField value={busqueda} onChange={setBusqueda} placeholder="Buscar por grupo, orden o nombre..." width={280} />
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={manejarArchivo} style={{ display: "none" }} />
        <button
          onClick={() => fileRef.current?.click()}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            background: "transparent",
            border: `1px solid ${COLORS.border}`,
            color: COLORS.textMuted,
            borderRadius: 6,
            padding: "7px 12px",
            fontSize: 13,
            cursor: "pointer",
          }}
        >
          <Upload size={14} /> Importar planilla
        </button>
        {importInfo && (
          <span style={{ fontSize: 12, color: COLORS.amber }}>
            Se agregaron {importInfo.agregados} invitado(s){importInfo.descartados > 0 && `, se descartaron ${importInfo.descartados} sin nombre`}.
          </span>
        )}
      </div>

      {invitados.length === 0 ? (
        <EmptyState text="Todavía no invitaste a nadie. Sumá un nombre y su WhatsApp para arrancar la ronda de invitación." />
      ) : filtrados.length === 0 ? (
        <EmptyState text="No hay invitados que coincidan con esa búsqueda." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {filtrados.map((p) => {
            const pct = calcularPorcentaje(p.montoAdjudica, valoresUnidades[baseModelo(p.plan)]);
            const abierto = expandido === p.id;
            return (
            <div
              key={p.id}
              style={{
                background: COLORS.panel,
                border: `1px solid ${COLORS.border}`,
                borderLeft: `3px solid ${p.estado === "confirmado" ? COLORS.amber : COLORS.textMuted}`,
                borderRadius: 8,
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  padding: "10px 14px",
                  display: "flex",
                  alignItems: "center",
                  gap: 12,
                  flexWrap: "wrap",
                  cursor: "pointer",
                }}
                onClick={() => setExpandido(abierto ? null : p.id)}
              >
              <div style={{ flex: "1 1 160px", fontWeight: 600 }}>{p.nombre}</div>
              {(p.grupo || p.orden) && (
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.amber, background: COLORS.amberSoft, borderRadius: 6, padding: "2px 7px" }}>
                  G{p.grupo || "—"} / O{p.orden || "—"}
                </span>
              )}
              <div style={{ color: COLORS.textMuted, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
                {p.telefono || "sin número"}
              </div>
              <Pill color={p.estado === "confirmado" ? COLORS.amber : COLORS.textMuted} bg={p.estado === "confirmado" ? COLORS.amberSoft : COLORS.panelAlt}>
                {p.estado === "confirmado" ? "Confirmado" : "Esperando respuesta"}
              </Pill>
              {p.prorrateo && <Pill color={COLORS.amber} bg={COLORS.amberSoft}>Prorrateo</Pill>}
              {p.subite && (
                <Pill color={COLORS.green} bg={COLORS.greenSoft}>
                  Subite{p.subiteCobertura ? ` · ${p.subiteCobertura}` : ""}
                </Pill>
              )}
              {(p.plan || p.montoAdjudica) && (
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.textMuted }}>
                  {p.plan} {p.montoAdjudica && `· $${p.montoAdjudica}`} {pct !== null && `· ${pct.toFixed(1)}%`}
                </span>
              )}

              <div style={{ display: "flex", gap: 6, marginLeft: "auto" }} onClick={(e) => e.stopPropagation()}>
                {p.telefono && (
                  <a
                    href={waLink(
                      p.telefono,
                      p.plan
                        ? `Hola ${p.nombre}! Te ofrecemos la posibilidad de adjudicar el plan ${p.plan}. ¿Te interesa? Contame con cuánto te gustaría entrar.`
                        : `Hola ${p.nombre}! Te invitamos a participar de la próxima adjudicación. ¿Con cuánto te gustaría adjudicar?`
                    )}
                    target="_blank"
                    rel="noreferrer"
                  >
                    <IconBtn title="Escribir por WhatsApp"><Phone size={14} /></IconBtn>
                  </a>
                )}
                {p.estado !== "confirmado" && (
                  <IconBtn title="Marcar como confirmado" onClick={() => onConfirmar(p.id)}>
                    <Check size={14} />
                  </IconBtn>
                )}
                {p.estado === "confirmado" && (
                  <button
                    onClick={() => onPasar(p.id)}
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 5,
                      background: "transparent",
                      border: `1px solid ${COLORS.amber}`,
                      color: COLORS.amber,
                      borderRadius: 6,
                      padding: "5px 10px",
                      fontSize: 12,
                      fontWeight: 600,
                      cursor: "pointer",
                    }}
                  >
                    Pasar a adjudicados <ArrowRight size={13} />
                  </button>
                )}
                <IconBtn title="Eliminar" danger onClick={() => onEliminar(p.id)}>
                  <Trash2 size={14} />
                </IconBtn>
              </div>
              </div>

              {abierto && (
                <div style={{ padding: "0 14px 16px", borderTop: `1px solid ${COLORS.border}` }}>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginTop: 12 }}>
                    <Field label="Nombre">
                      <TextField value={p.nombre} onChange={(v) => onActualizar(p.id, "nombre", v)} placeholder="Nombre" />
                    </Field>
                    <Field label="WhatsApp">
                      <TextField value={p.telefono} onChange={(v) => onActualizar(p.id, "telefono", v)} placeholder="5493..." mono />
                    </Field>
                    <Field label="Grupo">
                      <TextField value={p.grupo} onChange={(v) => onActualizar(p.id, "grupo", v)} placeholder="Grupo" mono />
                    </Field>
                    <Field label="Orden">
                      <TextField value={p.orden} onChange={(v) => onActualizar(p.id, "orden", v)} placeholder="Orden" mono />
                    </Field>
                    <Field label="Plan">
                      <SelectField value={p.plan} onChange={(v) => onActualizar(p.id, "plan", v)} options={planes} placeholder="Elegir plan" />
                    </Field>
                    <Field label="Adjudica con $">
                      <TextField value={p.montoAdjudica} onChange={(v) => onActualizar(p.id, "montoAdjudica", v)} placeholder="$" mono />
                    </Field>
                  </div>
                  <div style={{ marginTop: 10, fontSize: 12, color: p.telefono ? COLORS.green : COLORS.textMuted }}>
                    {p.telefono
                      ? `Número reconocido para WhatsApp: +${normalizarTelefono(p.telefono)}`
                      : "Cargá un número para poder enviarle WhatsApp"}
                  </div>

                  <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 14, fontSize: 13, cursor: "pointer" }}>
                    <input type="checkbox" checked={!!p.prorrateo} onChange={(e) => onActualizar(p.id, "prorrateo", e.target.checked)} />
                    Realiza prorrateo (adjudica con menos % y la diferencia se prorratea en las cuotas)
                  </label>

                  <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, fontSize: 13, cursor: "pointer" }}>
                    <input type="checkbox" checked={!!p.subite} onChange={(e) => onActualizar(p.id, "subite", e.target.checked)} />
                    Tiene plan Subite
                  </label>
                  {p.subite && (
                    <div style={{ marginTop: 6, marginLeft: 24 }}>
                      <SelectField
                        value={p.subiteCobertura}
                        onChange={(v) => onActualizar(p.id, "subiteCobertura", v)}
                        options={["Cubre el total", "Cubre una parte (el cliente abona el resto)"]}
                        placeholder="¿Qué cubre el Subite?"
                        width={280}
                      />
                    </div>
                  )}
                </div>
              )}
            </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* -------------------- Tab: Adjudicados -------------------- */

function AdjudicadosTab({ adjudicados, planes, valoresUnidades, nuevo, setNuevo, onAgregar, onActualizar, onToggleDoc, onAdjuntarDoc, onQuitarDocArchivo, onEliminar, onExportar }) {
  const [expandido, setExpandido] = useState(null);
  const aprobadosCount = adjudicados.filter((c) => c.carpetaAprobada).length;

  return (
    <div>
      <div
        style={{
          background: COLORS.panel,
          border: `1px solid ${COLORS.border}`,
          borderRadius: 10,
          padding: 14,
          marginBottom: 16,
          display: "flex",
          gap: 8,
          flexWrap: "wrap",
          alignItems: "center",
        }}
      >
        <TextField value={nuevo.nombre} onChange={(v) => setNuevo((n) => ({ ...n, nombre: v }))} placeholder="Nombre del cliente adjudicado" width={220} />
        <TextField value={nuevo.telefono} onChange={(v) => setNuevo((n) => ({ ...n, telefono: v }))} placeholder="WhatsApp (5493...)" mono width={180} />
        <SelectField value={nuevo.plan} onChange={(v) => setNuevo((n) => ({ ...n, plan: v }))} options={planes} placeholder="Plan" width={140} />
        <TextField value={nuevo.montoAdjudica} onChange={(v) => setNuevo((n) => ({ ...n, montoAdjudica: v }))} placeholder="Adjudica con $" mono width={130} />
        <button
          onClick={onAgregar}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            background: COLORS.amber,
            color: "#1A1200",
            border: "none",
            borderRadius: 6,
            padding: "7px 12px",
            fontWeight: 700,
            fontSize: 13,
            cursor: "pointer",
          }}
        >
          <Plus size={14} /> Cargar directo
        </button>
      </div>

      <div style={{ marginBottom: 16, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <button
          onClick={onExportar}
          disabled={aprobadosCount === 0}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            background: "transparent",
            border: `1px solid ${aprobadosCount > 0 ? COLORS.green : COLORS.border}`,
            color: aprobadosCount > 0 ? COLORS.green : COLORS.textMuted,
            borderRadius: 6,
            padding: "7px 12px",
            fontSize: 13,
            cursor: aprobadosCount > 0 ? "pointer" : "not-allowed",
          }}
        >
          <FileText size={14} /> Exportar aprobados a Excel ({aprobadosCount})
        </button>
      </div>

      {adjudicados.length === 0 ? (
        <EmptyState text="Todavía no hay clientes adjudicados. Se suman automáticamente cuando confirmás a alguien en la pestaña de Invitados." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {adjudicados.map((c) => {
            const status = paymentStatus(c);
            const st = STATUS_STYLE[status];
            const docsOk = DOC_ITEMS.filter((d) => docInfo(c, d).ok).length;
            const abierto = expandido === c.id;
            const pct = calcularPorcentaje(c.montoAdjudica, valoresUnidades[baseModelo(c.plan)]);
            return (
              <div
                key={c.id}
                style={{
                  background: COLORS.panel,
                  border: `1px solid ${COLORS.border}`,
                  borderLeft: `3px solid ${st.color}`,
                  borderRadius: 8,
                  overflow: "hidden",
                }}
              >
                <div
                  style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 14px", flexWrap: "wrap", cursor: "pointer" }}
                  onClick={() => setExpandido(abierto ? null : c.id)}
                >
                  <div style={{ flex: "1 1 160px", fontWeight: 600 }}>{c.nombre || "Sin nombre"}</div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.textMuted }}>
                    {c.modelo || "—"} {c.color && `· ${c.color}`}
                  </div>
                  {(c.plan || c.montoAdjudica) && (
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.textMuted }}>
                      {c.plan} {c.montoAdjudica && `· $${c.montoAdjudica}`} {pct !== null && `· ${pct.toFixed(1)}%`}
                    </span>
                  )}
                  <Pill color={st.color} bg={st.bg}>{st.label}</Pill>
                  <Pill color={c.pedidoRealizado ? COLORS.green : COLORS.textMuted} bg={c.pedidoRealizado ? COLORS.greenSoft : COLORS.panelAlt}>
                    {c.pedidoRealizado ? "Pedido hecho" : "Pedido pendiente"}
                  </Pill>
                  <span style={{ fontSize: 12, color: docsOk === DOC_ITEMS.length ? COLORS.green : COLORS.textMuted, display: "flex", alignItems: "center", gap: 4 }}>
                    <FileText size={13} /> {docsOk}/{DOC_ITEMS.length}
                  </span>
                  <div style={{ display: "flex", gap: 6, marginLeft: "auto" }} onClick={(e) => e.stopPropagation()}>
                    {c.telefono && (
                      <a href={waLink(c.telefono, reminderMessage(c))} target="_blank" rel="noreferrer">
                        <IconBtn title="Enviar recordatorio de pago"><Phone size={14} /></IconBtn>
                      </a>
                    )}
                    <IconBtn title="Eliminar" danger onClick={() => onEliminar(c.id)}>
                      <Trash2 size={14} />
                    </IconBtn>
                  </div>
                </div>

                {abierto && (
                  <div style={{ padding: "0 14px 16px", borderTop: `1px solid ${COLORS.border}` }}>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginTop: 12 }}>
                      <Field label="Modelo">
                        <TextField value={c.modelo} onChange={(v) => onActualizar(c.id, "modelo", v)} placeholder="Ej: Cronos" />
                      </Field>
                      <Field label="Color">
                        <TextField value={c.color} onChange={(v) => onActualizar(c.id, "color", v)} placeholder="Ej: Gris Digital" />
                      </Field>
                      <Field label="Plan">
                        <SelectField value={c.plan} onChange={(v) => onActualizar(c.id, "plan", v)} options={planes} placeholder="Elegir plan" />
                      </Field>
                      <Field label="Adjudica con $">
                        <TextField value={c.montoAdjudica} onChange={(v) => onActualizar(c.id, "montoAdjudica", v)} placeholder="$" mono />
                      </Field>
                      <Field label="Monto de cuota">
                        <TextField value={c.monto} onChange={(v) => onActualizar(c.id, "monto", v)} placeholder="$" mono />
                      </Field>
                      <Field label="Vencimiento">
                        <input
                          type="date"
                          value={c.vencimiento}
                          onChange={(e) => onActualizar(c.id, "vencimiento", e.target.value)}
                          style={{
                            background: COLORS.bg,
                            border: `1px solid ${COLORS.border}`,
                            borderRadius: 6,
                            color: COLORS.text,
                            fontSize: 13,
                            padding: "6px 8px",
                            width: "100%",
                          }}
                        />
                      </Field>
                    </div>

                    <div style={{ marginTop: 10, fontSize: 12, color: pct !== null ? COLORS.amber : COLORS.textMuted }}>
                      {pct !== null
                        ? `Porcentaje real de adjudicación: ${pct.toFixed(1)}% (según valor de unidad cargado)`
                        : "Cargá el valor real de la unidad en \"⚙ Planes y valores\" para calcular el porcentaje automáticamente"}
                    </div>

                    <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 12, fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={!!c.pagado} onChange={(e) => onActualizar(c.id, "pagado", e.target.checked)} />
                      Cuota actual pagada
                    </label>
                    <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={!!c.pedidoRealizado} onChange={(e) => onActualizar(c.id, "pedidoRealizado", e.target.checked)} />
                      Pedido cargado en el SGA
                    </label>
                    <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={!!c.veraz} onChange={(e) => onActualizar(c.id, "veraz", e.target.checked)} />
                      Titular afectado en Veraz
                    </label>
                    <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, fontSize: 13, cursor: "pointer" }}>
                      <input type="checkbox" checked={!!c.necesitaGarante} onChange={(e) => onActualizar(c.id, "necesitaGarante", e.target.checked)} />
                      Necesita garante
                    </label>

                    {c.necesitaGarante && (
                      <div style={{ marginTop: 6, marginLeft: 24 }}>
                        <TextField
                          value={c.garanteNombre}
                          onChange={(v) => onActualizar(c.id, "garanteNombre", v)}
                          placeholder="Nombre del garante"
                          width={220}
                        />
                      </div>
                    )}

                    <div style={{ marginTop: 12 }}>
                      <Field label="Observaciones">
                        <textarea
                          value={c.observaciones}
                          onChange={(e) => onActualizar(c.id, "observaciones", e.target.value)}
                          placeholder="Otras notas: situación crediticia, acuerdos particulares, etc."
                          rows={2}
                          style={{
                            background: COLORS.bg,
                            border: `1px solid ${COLORS.border}`,
                            borderRadius: 6,
                            color: COLORS.text,
                            fontSize: 13,
                            padding: "6px 8px",
                            width: "100%",
                            fontFamily: "inherit",
                            resize: "vertical",
                          }}
                        />
                      </Field>
                    </div>

                    <DocsChecklist
                      cliente={c}
                      campo="docs"
                      titulo="Documentación a solicitar (titular)"
                      onToggleDoc={onToggleDoc}
                      onAdjuntarDoc={onAdjuntarDoc}
                      onQuitarDocArchivo={onQuitarDocArchivo}
                    />

                    {c.necesitaGarante && (
                      <DocsChecklist
                        cliente={c}
                        campo="docsGarante"
                        titulo={`Documentación a solicitar (garante${c.garanteNombre ? ": " + c.garanteNombre : ""})`}
                        onToggleDoc={onToggleDoc}
                        onAdjuntarDoc={onAdjuntarDoc}
                        onQuitarDocArchivo={onQuitarDocArchivo}
                      />
                    )}

                    <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 14, fontSize: 13, cursor: "pointer", fontWeight: 600 }}>
                      <input type="checkbox" checked={!!c.carpetaAprobada} onChange={(e) => onActualizar(c.id, "carpetaAprobada", e.target.checked)} />
                      <span style={{ color: c.carpetaAprobada ? COLORS.green : COLORS.text }}>Carpeta crediticia aprobada (último paso)</span>
                    </label>

                    {c.vencimiento && status !== "pagado" && (
                      <div style={{ marginTop: 12, fontSize: 12, color: st.color, display: "flex", alignItems: "center", gap: 5 }}>
                        <Clock size={13} />
                        {status === "vencido"
                          ? `Vencida hace ${Math.abs(daysUntil(c.vencimiento))} día(s)`
                          : `Vence en ${daysUntil(c.vencimiento)} día(s)`}
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function DocsChecklist({ cliente, campo, titulo, onToggleDoc, onAdjuntarDoc, onQuitarDocArchivo }) {
  return (
    <div style={{ marginTop: 14 }}>
      <div style={{ fontSize: 12, color: COLORS.textMuted, marginBottom: 6, display: "flex", alignItems: "center", gap: 5 }}>
        <Tag size={12} /> {titulo}
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {DOC_ITEMS.map((doc) => {
          const info = docInfo(cliente, doc, campo);
          const inputId = `doc-${campo}-${cliente.id}-${doc.replace(/\s+/g, "")}`;
          return (
            <div
              key={doc}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                background: info.ok ? COLORS.greenSoft : COLORS.panelAlt,
                border: `1px solid ${info.ok ? COLORS.green : COLORS.border}`,
                borderRadius: 8,
                padding: "7px 10px",
                flexWrap: "wrap",
              }}
            >
              <button
                onClick={() => onToggleDoc(cliente.id, doc, campo)}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 5,
                  background: "transparent",
                  border: "none",
                  color: info.ok ? COLORS.green : COLORS.textMuted,
                  fontSize: 12,
                  cursor: "pointer",
                  padding: 0,
                }}
              >
                {info.ok ? <Check size={13} /> : <X size={13} />} {doc}
              </button>

              {info.archivo && (
                <a href={info.archivo} target="_blank" rel="noreferrer">
                  <img
                    src={info.archivo}
                    alt={doc}
                    style={{ height: 34, width: 34, objectFit: "cover", borderRadius: 4, border: `1px solid ${COLORS.border}` }}
                  />
                </a>
              )}
              {info.nombreArchivo && (
                <span style={{ fontSize: 11, color: COLORS.textMuted, maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {info.nombreArchivo}
                </span>
              )}

              <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
                <input
                  id={inputId}
                  type="file"
                  accept="image/*,.pdf"
                  style={{ display: "none" }}
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    comprimirImagen(file, (res) => onAdjuntarDoc(cliente.id, doc, res, campo));
                    e.target.value = "";
                  }}
                />
                <label
                  htmlFor={inputId}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 4,
                    fontSize: 11,
                    color: COLORS.amber,
                    border: `1px solid ${COLORS.amber}55`,
                    borderRadius: 6,
                    padding: "3px 8px",
                    cursor: "pointer",
                  }}
                >
                  <Upload size={11} /> {info.archivo || info.nombreArchivo ? "Reemplazar" : "Adjuntar"}
                </label>
                {(info.archivo || info.nombreArchivo) && (
                  <IconBtn title="Quitar archivo" danger onClick={() => onQuitarDocArchivo(cliente.id, doc, campo)}>
                    <Trash2 size={12} />
                  </IconBtn>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div>
      <div style={{ fontSize: 11, color: COLORS.textMuted, marginBottom: 4 }}>{label}</div>
      {children}
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div
      style={{
        border: `1px dashed ${COLORS.border}`,
        borderRadius: 10,
        padding: "28px 20px",
        textAlign: "center",
        color: COLORS.textMuted,
        fontSize: 13,
      }}
    >
      {text}
    </div>
  );
}
