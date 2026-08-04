import React, { useState, useEffect, useCallback, useRef } from "react";import {Phone,Check,X,Plus,ArrowRight,FileText,Clock,Trash2,RefreshCw,Upload,Search,Download,Eye,AlertTriangle,ShieldCheck,Filter,FileSpreadsheet,MessageSquare,ChevronDown,ChevronUp,Copy,Settings,CheckCircle2,AlertCircle,FileCheck,UserCheck,FolderDownload,Building2,Percent,Coins,HardDrive,ExternalLink,CreditCard,Building,Briefcase} from "lucide-react";import * as XLSX from "xlsx";/* ---------------------------------------------------------Panel de Adjudicaciones WOCH - Fiat FaduaPaleta: petróleo profundo + ámbar + acentos verdes/rojos--------------------------------------------------------- */const COLORS = {bg: "#0F1720",panel: "#18222D",panelAlt: "#1E2A37",border: "#2B3948",text: "#EDEFF2",textMuted: "#8A97A6",amber: "#E8A23D",amberSoft: "#3A2E1A",green: "#4F9D69",greenSoft: "#16261C",red: "#C1493F",redSoft: "#2C1917",blue: "#3B82F6",blueSoft: "#1E293B",purple: "#8B5CF6",purpleSoft: "#2E1065",};const DOC_ITEMS = ["DNI", "Recibo de sueldo", "Constancia CUIL", "Comprobante de domicilio"];const DEFAULT_PLANES = ["Cronos 80/20","Cronos 90/10","Cronos 70/30","Mobi 80/20","Fastback 70/30","Fiorino 70/30","Pulse 60/40","Titano 70/30","Titano 60/40","Argo 70/30",];const MEDIOS_PAGO = ["Banco Galicia","Agencia - Transferencia","Agencia - Efectivo","Agencia - Tarjeta de Crédito","Agencia - Tarjeta de Débito","Homebanking",];const baseModelo = (planLabel) => (planLabel || "").split(" ")[0];function parseMonto(valor) {if (!valor) return null;const n = Number(String(valor).replace(/[^\d.,-]/g, "").replace(/./g, "").replace(",", "."));return Number.isFinite(n) && n > 0 ? n : null;}function calcularPorcentaje(montoAdjudica, valorUnidad) {const monto = parseMonto(montoAdjudica);const valor = parseMonto(valorUnidad);if (!monto || !valor) return null;return (monto / valor) * 100;}const uid = () => Math.random().toString(36).slice(2, 10);function docInfo(c, doc, campo = "docs") {const v = c[campo]?.[doc];if (v && typeof v === "object") return { ok: !!v.ok, archivo: v.archivo || null, nombreArchivo: v.nombreArchivo || "" };return { ok: !!v, archivo: null, nombreArchivo: "" };}/* Limpieza y extracción inteligente de múltiples números telefónicos */function extraerTelefonos(cadena) {if (!cadena) return [];const raw = String(cadena);// Dividir por delimitadores comunes como separadores de notas (c), (p), /, ,, ;, retornos de líneaconst partes = raw.split(/[/\n,;(]|(?<=\d)\s*(?=[a-zA-Z]|()/);const resultados = [];partes.forEach((p) => {// Extraer etiqueta opcional como (c) celular, (p) particularlet etiqueta = "Teléfono";if (/\b[cC]\b|celular|cel/i.test(p)) etiqueta = "Celular";else if (/\b[pP]\b|part|particular/i.test(p)) etiqueta = "Particular";else if (/trabajo|laboral/i.test(p)) etiqueta = "Trabajo";let soloNum = p.replace(/[^\d]/g, "");
if (!soloNum || soloNum.length < 6) return;

// Normalización de números de Argentina
if (soloNum.startsWith("0")) soloNum = soloNum.slice(1);
if (!soloNum.startsWith("54")) {
  if (soloNum.startsWith("15") && soloNum.length === 10) {
    soloNum = "549" + soloNum.slice(2);
  } else if (soloNum.startsWith("9")) {
    soloNum = "54" + soloNum;
  } else {
    soloNum = "549" + soloNum;
  }
} else if (soloNum.startsWith("54") && !soloNum.startsWith("549") && soloNum.length === 12) {
  soloNum = "549" + soloNum.slice(2);
}

if (!resultados.some((r) => r.clean === soloNum)) {
  resultados.push({
    original: p.trim(),
    clean: soloNum,
    etiqueta,
  });
}
});return resultados;}function waLink(telefonoClean, mensaje) {return https://wa.me/${telefonoClean}?text=${encodeURIComponent(mensaje)};}function comprimirImagen(file, callback) {if (!file.type.startsWith("image/")) {callback({ archivo: null, nombreArchivo: file.name });return;}const reader = new FileReader();reader.onload = (e) => {const img = new Image();img.onload = () => {const maxW = 900;const scale = Math.min(1, maxW / img.width);const canvas = document.createElement("canvas");canvas.width = img.width * scale;canvas.height = img.height * scale;const ctx = canvas.getContext("2d");ctx.drawImage(img, 0, 0, canvas.width, canvas.height);callback({ archivo: canvas.toDataURL("image/jpeg", 0.7), nombreArchivo: file.name });};img.src = e.target.result;};reader.readAsDataURL(file);}function daysUntil(dateStr) {if (!dateStr) return null;const today = new Date();today.setHours(0, 0, 0, 0);const target = new Date(dateStr + "T00:00:00");return Math.round((target - today) / 86400000);}function paymentStatus(cliente) {if (cliente.pagado) return "pagado";const d = daysUntil(cliente.vencimiento);if (d === null) return "sin-fecha";if (d < 0) return "vencido";if (d <= 5) return "proximo";return "al-dia";}const STATUS_STYLE = {pagado: { label: "Pagado", color: COLORS.green, bg: COLORS.greenSoft },"al-dia": { label: "Al día", color: COLORS.green, bg: COLORS.greenSoft },proximo: { label: "Vence pronto", color: COLORS.amber, bg: COLORS.amberSoft },vencido: { label: "Vencido", color: COLORS.red, bg: COLORS.redSoft },"sin-fecha": { label: "Sin fecha", color: COLORS.textMuted, bg: COLORS.panelAlt },};function quitarAcentos(txt) {return (txt || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");}function detectarColumna(headers, candidatos) {const norm = headers.map((h) => quitarAcentos(String(h)).toLowerCase().trim());for (const cand of candidatos) {const idx = norm.findIndex((h) => h.includes(cand));if (idx !== -1) return headers[idx];}return null;}function mapearFilaImportada(fila, headers) {const colNombre = detectarColumna(headers, ["nombre", "cliente", "apellido"]);const colTelefono = detectarColumna(headers, ["telefono", "whatsapp", "celular", "numero", "telefonos"]);const colPlan = detectarColumna(headers, ["plan", "modelo"]);const colGrupo = detectarColumna(headers, ["grupo"]);const colOrden = detectarColumna(headers, ["orden"]);const colMonto = detectarColumna(headers, ["monto", "adjudica", "importe"]);return {nombre: colNombre ? String(fila[colNombre] ?? "").trim() : "",telefono: colTelefono ? String(fila[colTelefono] ?? "").trim() : "",plan: colPlan ? String(fila[colPlan] ?? "").trim() : "",grupo: colGrupo ? String(fila[colGrupo] ?? "").trim() : "",orden: colOrden ? String(fila[colOrden] ?? "").trim() : "",montoAdjudica: colMonto ? String(fila[colMonto] ?? "").trim() : "",subite: false,prorratea: false,porcentajeProrrateo: "",montoProrrateo: "",abonaEnAgencia: false,};}function generarMensajeWA(tipo, cliente, params = {}) {const nombre = cliente.nombre || "Cliente";const fecha = params.fecha || cliente.vencimiento || "esta semana";const promoPct = params.promoPct || "20%";switch (tipo) {case "invitacion":return Hola ${nombre}! Te escribo de Fiat Fadua. Te informamos que este mes contas con una bonificación especial del ${promoPct} para adjudicar tu plan ${cliente.plan || ""}. ¿Coordinamos una llamada antes del ${fecha}?;case "documentos":return Hola ${nombre}! Para avanzar con la carpeta de tu adjudicación en Fiat Fadua, necesitamos adjuntar la siguiente documentación pendiente: DNI, Recibo de sueldo y Comprobante de domicilio. ¿Me los podrías enviar por aquí?;case "pago_proximo":return Hola ${nombre}! Te recordamos que la fecha límite para el pago de tu cuota de adjudicación es el ${fecha}. Ante cualquier duda estoy a disposición.;case "pago_vencido":return Hola ${nombre}! Te contactamos de Fiat Fadua. Tu cuota de adjudicación venció el ${fecha}. Por favor comunícate con nosotros para regularizar el pago y no perder la adjudicación.;case "aprobado":return ¡Buenas noticias ${nombre}! Tu carpeta ha sido APROBADA con éxito en Fiat Fadua. Estamos listos para continuar con la carga del pedido del vehículo en el sistema SGA.;default:return Hola ${nombre}! Te escribo de Fiat Fadua en relación a tu adjudicación.;}}/* Cargar dinámicamente JSZip si no está presente para empaquetar carpetas */async function descargarCarpetasZip(nombreCarpetaMadre, listaAdjudicados) {if (!window.JSZip) {await new Promise((resolve) => {const script = document.createElement("script");script.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";script.onload = resolve;document.head.appendChild(script);});}const zip = new window.JSZip();const folderRaiz = zip.folder(nombreCarpetaMadre || "adj_agosto");let archivosAgregados = 0;listaAdjudicados.forEach((cliente) => {const folderClienteName = (cliente.nombre || "Sin_Nombre").replace(/[^a-zA-Z0-9_- ]/g, "_");const folderCliente = folderRaiz.folder(folderClienteName);// Titular docs
DOC_ITEMS.forEach((d) => {
  const info = docInfo(cliente, d, "docs");
  if (info.ok && info.archivo) {
    const base64Data = info.archivo.split(",")[1];
    if (base64Data) {
      folderCliente.file(`TITULAR_${d.replace(/\s+/g, "_")}.jpg`, base64Data, { base64: true });
      archivosAgregados++;
    }
  }
});

// Garante docs
if (cliente.necesitaGarante) {
  const folderGarante = folderCliente.folder(`GARANTE_${(cliente.garanteNombre || "Garante").replace(/[^a-zA-Z0-9_\- ]/g, "_")}`);
  DOC_ITEMS.forEach((d) => {
    const info = docInfo(cliente, d, "docsGarante");
    if (info.ok && info.archivo) {
      const base64Data = info.archivo.split(",")[1];
      if (base64Data) {
        folderGarante.file(`GARANTE_${d.replace(/\s+/g, "_")}.jpg`, base64Data, { base64: true });
        archivosAgregados++;
      }
    }
  });
}
});if (archivosAgregados === 0) {alert("No se encontraron archivos de imágenes adjuntos en las carpetas de los clientes seleccionados.");return;}const blob = await zip.generateAsync({ type: "blob" });const a = document.createElement("a");a.href = URL.createObjectURL(blob);a.download = ${nombreCarpetaMadre || "adj_agosto"}_Documentacion.zip;a.click();}/* -------------------- UI Elements -------------------- */function Pill({ children, color, bg }) {return (<spanstyle={{color,background: bg,border: 1px solid ${color}55,borderRadius: 999,padding: "3px 10px",fontSize: 12,fontWeight: 600,whiteSpace: "nowrap",display: "inline-flex",alignItems: "center",gap: 4}}>{children});}function IconBtn({ onClick, title, children, danger, highlight }) {return (<buttononClick={onClick}title={title}style={{display: "inline-flex",alignItems: "center",justifyContent: "center",width: 32,height: 32,borderRadius: 8,border: 1px solid ${highlight ? COLORS.amber : COLORS.border},background: highlight ? COLORS.amberSoft : "transparent",color: danger ? COLORS.red : highlight ? COLORS.amber : COLORS.textMuted,cursor: "pointer",transition: "all 0.15s ease",}}>{children});}function TextField({ value, onChange, placeholder, mono, width, type = "text" }) {return (<inputtype={type}value={value}onChange={(e) => onChange(e.target.value)}placeholder={placeholder}style={{background: COLORS.bg,border: 1px solid ${COLORS.border},borderRadius: 6,color: COLORS.text,fontFamily: mono ? "'IBM Plex Mono', monospace" : "inherit",fontSize: 13,padding: "7px 10px",width: width || "100%",outline: "none",}}/>);}function SelectField({ value, onChange, options, placeholder, width }) {return (<selectvalue={value}onChange={(e) => onChange(e.target.value)}style={{background: COLORS.bg,border: 1px solid ${COLORS.border},borderRadius: 6,color: value ? COLORS.text : COLORS.textMuted,fontFamily: "'IBM Plex Mono', monospace",fontSize: 13,padding: "7px 10px",width: width || "100%",outline: "none",cursor: "pointer"}}>{placeholder}{options.map((o) => (<option key={o} value={o} style={{ color: COLORS.text, background: COLORS.panel }}>{o}))});}function EmptyState({ text }) {return (<divstyle={{padding: "40px 20px",textAlign: "center",color: COLORS.textMuted,background: COLORS.panel,borderRadius: 10,border: 1px dashed ${COLORS.border},fontSize: 14,}}>{text});}/* -------------------- Modals -------------------- */function DocModal({ docData, onClose }) {if (!docData) return null;return (<divstyle={{position: "fixed",top: 0, left: 0, right: 0, bottom: 0,background: "rgba(0,0,0,0.85)",display: "flex",alignItems: "center",justifyContent: "center",zIndex: 1000,padding: 20}}onClick={onClose}><divstyle={{background: COLORS.panel,border: 1px solid ${COLORS.border},borderRadius: 12,padding: 20,maxWidth: "90vw",maxHeight: "90vh",display: "flex",flexDirection: "column",gap: 12}}onClick={(e) => e.stopPropagation()}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}><h3 style={{ margin: 0, fontSize: 16, color: COLORS.text }}>{docData.nombreArchivo || "Vista Previa de Documento"}<button onClick={onClose} style={{ background: "none", border: "none", color: COLORS.textMuted, cursor: "pointer" }}><div style={{ overflow: "auto", textAlign: "center", maxHeight: "75vh" }}>{docData.archivo ? (<img src={docData.archivo} alt="Documento" style={{ maxWidth: "100%", height: "auto", borderRadius: 8 }} />) : (<div style={{ padding: 40, color: COLORS.textMuted }}>Vista previa no disponible.)});}function WAModal({ cliente, onClose }) {const [tipo, setTipo] = useState("invitacion");const [fechaCustom, setFechaCustom] = useState("");const [promoPct, setPromoPct] = useState("20%");const [mensaje, setMensaje] = useState("");const [copiado, setCopiado] = useState(false);const [telefonoSeleccionado, setTelefonoSeleccionado] = useState("");const telefonosEncontrados = extraerTelefonos(cliente?.telefono);useEffect(() => {if (telefonosEncontrados.length > 0 && !telefonoSeleccionado) {setTelefonoSeleccionado(telefonosEncontrados[0].clean);}}, [cliente, telefonosEncontrados, telefonoSeleccionado]);useEffect(() => {if (cliente) {setMensaje(generarMensajeWA(tipo, cliente, { fecha: fechaCustom, promoPct }));}}, [tipo, cliente, fechaCustom, promoPct]);if (!cliente) return null;const handleCopiar = () => {navigator.clipboard.writeText(mensaje);setCopiado(true);setTimeout(() => setCopiado(false), 2000);};return (<divstyle={{position: "fixed",top: 0, left: 0, right: 0, bottom: 0,background: "rgba(0,0,0,0.85)",display: "flex",alignItems: "center",justifyContent: "center",zIndex: 1000,padding: 20}}onClick={onClose}><divstyle={{background: COLORS.panel,border: 1px solid ${COLORS.border},borderRadius: 12,padding: 20,width: 520,maxWidth: "100%",display: "flex",flexDirection: "column",gap: 14}}onClick={(e) => e.stopPropagation()}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}><div style={{ display: "flex", alignItems: "center", gap: 8 }}><h3 style={{ margin: 0, fontSize: 16, color: COLORS.text }}>Enviar WhatsApp a {cliente.nombre}<button onClick={onClose} style={{ background: "none", border: "none", color: COLORS.textMuted, cursor: "pointer" }}>    {/* Selección de teléfono destino si hay varios */}
    {telefonosEncontrados.length > 0 && (
      <div>
        <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 4 }}>
          Teléfonos detectados para este cliente:
        </label>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {telefonosEncontrados.map((t) => (
            <button
              key={t.clean}
              onClick={() => setTelefonoSeleccionado(t.clean)}
              style={{
                background: telefonoSeleccionado === t.clean ? COLORS.amberSoft : COLORS.bg,
                border: `1px solid ${telefonoSeleccionado === t.clean ? COLORS.amber : COLORS.border}`,
                color: telefonoSeleccionado === t.clean ? COLORS.amber : COLORS.text,
                borderRadius: 6,
                padding: "4px 8px",
                fontSize: 12,
                fontFamily: "'IBM Plex Mono', monospace",
                cursor: "pointer"
              }}
            >
              {t.etiqueta}: {t.clean}
            </button>
          ))}
        </div>
      </div>
    )}

    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
      <div>
        <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 4 }}>Plantilla</label>
        <select
          value={tipo}
          onChange={(e) => setTipo(e.target.value)}
          style={{
            width: "100%",
            background: COLORS.bg,
            border: `1px solid ${COLORS.border}`,
            color: COLORS.text,
            borderRadius: 6,
            padding: "7px 10px",
            fontSize: 12
          }}
        >
          <option value="invitacion">Invitación a adjudicar (Promo %)</option>
          <option value="documentos">Solicitud de documentación</option>
          <option value="pago_proximo">Recordatorio de vencimiento</option>
          <option value="pago_vencido">Alerta de cuota vencida</option>
          <option value="aprobado">Notificación Carpeta Aprobada</option>
        </select>
      </div>

      {tipo === "invitacion" ? (
        <div>
          <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 4 }}>% Bonificación Promo</label>
          <TextField value={promoPct} onChange={setPromoPct} placeholder="Ej: 20%" mono />
        </div>
      ) : (
        <div>
          <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 4 }}>Fecha tope del mes</label>
          <TextField value={fechaCustom} onChange={setFechaCustom} placeholder="Ej: 15 de Agosto" />
        </div>
      )}
    </div>

    <div>
      <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 4 }}>Mensaje a enviar</label>
      <textarea
        value={mensaje}
        onChange={(e) => setMensaje(e.target.value)}
        rows={5}
        style={{
          width: "100%",
          background: COLORS.bg,
          border: `1px solid ${COLORS.border}`,
          color: COLORS.text,
          borderRadius: 6,
          padding: 10,
          fontSize: 13,
          resize: "none"
        }}
      />
    </div>

    <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 4 }}>
      <button
        onClick={handleCopiar}
        style={{
          display: "inline-flex",
          alignItems: "center",
          gap: 6,
          background: "transparent",
          border: `1px solid ${COLORS.border}`,
          color: COLORS.text,
          borderRadius: 6,
          padding: "8px 14px",
          fontSize: 13,
          cursor: "pointer"
        }}
      >
        {copiado ? <Check size={14} color={COLORS.green} /> : <Copy size={14} />} {copiado ? "¡Copiado!" : "Copiar"}
      </button>
      <a
        href={waLink(telefonoSeleccionado || cliente.telefono, mensaje)}
        target="_blank"
        rel="noreferrer"
        style={{
          display: "inline-flex",
          alignItems: "center",
          gap: 6,
          background: COLORS.green,
          color: "#FFF",
          borderRadius: 6,
          padding: "8px 16px",
          fontSize: 13,
          fontWeight: 600,
          textDecoration: "none"
        }}
      >
        <Phone size={14} /> Abrir WhatsApp
      </a>
    </div>
  </div>
</div>
);}/* Modal de Integración SGA & Google Drive */function IntegrationHelpModal({ onClose }) {return (<divstyle={{position: "fixed",top: 0, left: 0, right: 0, bottom: 0,background: "rgba(0,0,0,0.85)",display: "flex",alignItems: "center",justifyContent: "center",zIndex: 1000,padding: 20}}onClick={onClose}><divstyle={{background: COLORS.panel,border: 1px solid ${COLORS.border},borderRadius: 12,padding: 22,maxWidth: 620,width: "100%",maxHeight: "85vh",overflowY: "auto"}}onClick={(e) => e.stopPropagation()}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}><div style={{ display: "flex", alignItems: "center", gap: 8 }}><h3 style={{ margin: 0, fontSize: 17, color: COLORS.text }}>Alojamiento en Drive y Conexión SGA<button onClick={onClose} style={{ background: "none", border: "none", color: COLORS.textMuted, cursor: "pointer" }}>    <div style={{ display: "flex", flexDirection: "column", gap: 16, fontSize: 13, color: COLORS.textMuted }}>
      <div style={{ background: COLORS.panelAlt, padding: 14, borderRadius: 8, border: `1px solid ${COLORS.border}` }}>
        <div style={{ fontWeight: 700, color: COLORS.amber, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
          <HardDrive size={15} /> Backup directo a Google Drive
        </div>
        <p style={{ margin: 0, lineHeight: 1.5 }}>
          Puedes descargar la copia de seguridad `.json` o la carpeta comprimida `.zip` con todos los documentos de los clientes organizados por carpetas y guardarla en tu Google Drive sincronizado en tu PC (Google Drive para Escritorio), logrando almacenamiento en la nube sin costo adicional.
        </p>
      </div>

      <div style={{ background: COLORS.panelAlt, padding: 14, borderRadius: 8, border: `1px solid ${COLORS.border}` }}>
        <div style={{ fontWeight: 700, color: COLORS.green, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
          <ExternalLink size={15} /> Estrategias de Conexión con SGA (Sistema de Gestión de Ahorro)
        </div>
        <p style={{ margin: "0 0 8px", lineHeight: 1.5 }}>
          Como el portal SGA corporativo de Fiat/Stellantis no cuenta con una API pública abierta para concesionarios, la conexión se realiza de 3 formas:
        </p>
        <ol style={{ margin: 0, paddingLeft: 20, lineHeight: 1.6 }}>
          <li><strong>Exportación/Importación de Planillas:</strong> Exportar la lista de pedidos aprobados desde WOCH a Excel y cargarla en lote en SGA.</li>
          <li><strong>RPA / Robot Automatizado (Selenium/Puppeteer):</strong> Un script que lee la carpeta aprobada e ingresa automáticamente los datos en el portal SGA.</li>
          <li><strong>Extensión de Navegador Chrome:</strong> Una extensión que detecta cuando estás logueado en SGA y sincroniza el estado del pedido con WOCH en un clic.</li>
        </ol>
      </div>
    </div>
  </div>
</div>
);}function ConfigPanel({ planes, onAgregarPlan, onEliminarPlan, valoresUnidades, onActualizarValor, onBackup, onRestore }) {const [nuevoPlan, setNuevoPlan] = useState("");const modelos = [...new Set(planes.map(baseModelo))];const backupRef = useRef(null);return (<divstyle={{background: COLORS.panel,border: 1px solid ${COLORS.border},borderRadius: 10,padding: 16,marginBottom: 18,}}><div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 20 }}><div style={{ fontSize: 12, color: COLORS.textMuted, marginBottom: 8, fontWeight: 600 }}>Planes disponibles<div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>{planes.map((p) => (<spankey={p}style={{display: "inline-flex",alignItems: "center",gap: 6,background: COLORS.panelAlt,border: 1px solid ${COLORS.border},borderRadius: 999,padding: "4px 4px 4px 10px",fontSize: 12,fontFamily: "'IBM Plex Mono', monospace",}}>{p}<buttononClick={() => onEliminarPlan(p)}style={{ background: "none", border: "none", color: COLORS.textMuted, cursor: "pointer", display: "flex", padding: 2 }}>))}<div style={{ display: "flex", gap: 6 }}><IconBtn title="Agregar plan" onClick={() => { if (nuevoPlan.trim()) { onAgregarPlan(nuevoPlan.trim()); setNuevoPlan(""); } }}>    <div>
      <div style={{ fontSize: 12, color: COLORS.textMuted, marginBottom: 8, fontWeight: 600 }}>
        Valor real de la unidad ($) — actualización mensual
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

    <div style={{ borderTop: `1px solid ${COLORS.border}`, paddingTop: 12, gridColumn: "1 / -1", display: "flex", gap: 12, alignItems: "center", justifyContent: "space-between", flexWrap: "wrap" }}>
      <span style={{ fontSize: 12, color: COLORS.textMuted }}>Copia de seguridad local y exportación</span>
      <div style={{ display: "flex", gap: 8 }}>
        <button
          onClick={onBackup}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            background: COLORS.panelAlt,
            border: `1px solid ${COLORS.border}`,
            color: COLORS.text,
            borderRadius: 6,
            padding: "6px 10px",
            fontSize: 12,
            cursor: "pointer"
          }}
        >
          <Download size={13} /> Backup JSON
        </button>
        <input type="file" accept=".json" ref={backupRef} onChange={onRestore} style={{ display: "none" }} />
        <button
          onClick={() => backupRef.current?.click()}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            background: COLORS.panelAlt,
            border: `1px solid ${COLORS.border}`,
            color: COLORS.text,
            borderRadius: 6,
            padding: "6px 10px",
            fontSize: 12,
            cursor: "pointer"
          }}
        >
          <Upload size={13} /> Restaurar Backup
        </button>
      </div>
    </div>
  </div>
</div>
);}/* -------------------- Main App -------------------- */export default function App() {const [tab, setTab] = useState("invitados");const [invitados, setInvitados] = useState([]);const [adjudicados, setAdjudicados] = useState([]);const [planes, setPlanes] = useState(DEFAULT_PLANES);const [valoresUnidades, setValoresUnidades] = useState({});const [mostrarConfig, setMostrarConfig] = useState(false);const [mostrarHelpModal, setMostrarHelpModal] = useState(false);const [nombreCarpetaZip, setNombreCarpetaZip] = useState("adj agosto");const [loading, setLoading] = useState(true);const [docModalData, setDocModalData] = useState(null);const [waClienteModal, setWaClienteModal] = useState(null);// Form de alta Invitadosconst [nuevoInvitado, setNuevoInvitado] = useState({nombre: "",telefono: "",grupo: "",orden: "",plan: "",montoAdjudica: "",subite: false,prorratea: false,porcentajeProrrateo: "",montoProrrateo: "",abonaEnAgencia: false,});// Form de alta Adjudicadosconst [nuevoAdj, setNuevoAdj] = useState({nombre: "",telefono: "",plan: "",montoAdjudica: "",medioPagoAdjudicacion: "",adjudicacionAdministrativa: false,});const persist = useCallback((key, value) => {try {localStorage.setItem(key, JSON.stringify(value));} catch (err) {console.error("Storage error", err);}}, []);useEffect(() => {try {const inv = localStorage.getItem("invitados");const adj = localStorage.getItem("adjudicados");const pl = localStorage.getItem("planes");const vu = localStorage.getItem("valoresUnidades");setInvitados(inv ? JSON.parse(inv) : []);setAdjudicados(adj ? JSON.parse(adj) : []);setPlanes(pl ? JSON.parse(pl) : DEFAULT_PLANES);setValoresUnidades(vu ? JSON.parse(vu) : {});} catch {// ignore} finally {setLoading(false);}}, []);useEffect(() => { if (!loading) persist("invitados", invitados); }, [invitados, loading, persist]);useEffect(() => { if (!loading) persist("adjudicados", adjudicados); }, [adjudicados, loading, persist]);useEffect(() => { if (!loading) persist("planes", planes); }, [planes, loading, persist]);useEffect(() => { if (!loading) persist("valoresUnidades", valoresUnidades); }, [valoresUnidades, loading, persist]);const agregarPlan = (plan) => setPlanes((prev) => (prev.includes(plan) ? prev : [...prev, plan]));const eliminarPlan = (plan) => setPlanes((prev) => prev.filter((p) => p !== plan));const actualizarValorUnidad = (modelo, valor) => setValoresUnidades((prev) => ({ ...prev, [modelo]: valor }));const exportarBackupData = () => {const data = { invitados, adjudicados, planes, valoresUnidades, exportDate: new Date().toISOString() };const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });const url = URL.createObjectURL(blob);const a = document.createElement("a");a.href = url;a.download = backup-woch-${new Date().toISOString().slice(0, 10)}.json;a.click();};const restaurarBackupData = (e) => {const file = e.target.files?.[0];if (!file) return;const reader = new FileReader();reader.onload = (ev) => {try {const data = JSON.parse(ev.target.result);if (data.invitados) setInvitados(data.invitados);if (data.adjudicados) setAdjudicados(data.adjudicados);if (data.planes) setPlanes(data.planes);if (data.valoresUnidades) setValoresUnidades(data.valoresUnidades);alert("¡Copia de seguridad restaurada con éxito!");} catch {alert("Error al leer el archivo JSON.");}};reader.readAsText(file);e.target.value = "";};/* ---- Invitados actions ---- */const agregarInvitado = () => {if (!nuevoInvitado.nombre.trim()) return;setInvitados((prev) => [...prev,{id: uid(),...nuevoInvitado,nombre: nuevoInvitado.nombre.trim(),telefono: nuevoInvitado.telefono.trim(),estado: "invitado",},]);setNuevoInvitado({nombre: "",telefono: "",grupo: "",orden: "",plan: "",montoAdjudica: "",subite: false,prorratea: false,porcentajeProrrateo: "",montoProrrateo: "",abonaEnAgencia: false,});};const marcarConfirmado = (id) => {setInvitados((prev) => prev.map((p) => (p.id === id ? { ...p, estado: "confirmado" } : p)));};const pasarAAdjudicados = (id) => {const persona = invitados.find((p) => p.id === id);if (!persona) return;setAdjudicados((prev) => [...prev,{id: uid(),nombre: persona.nombre,telefono: persona.telefono,grupo: persona.grupo || "",orden: persona.orden || "",plan: persona.plan || "",montoAdjudica: persona.montoAdjudica || "",abonaEnAgencia: persona.abonaEnAgencia || false,subite: persona.subite || false,prorratea: persona.prorratea || false,porcentajeProrrateo: persona.porcentajeProrrateo || "",montoProrrateo: persona.montoProrrateo || "",medioPagoAdjudicacion: "",adjudicacionAdministrativa: false,modelo: "",color: "",monto: "",vencimiento: "",pagado: false,pedidoRealizado: false,observaciones: "",veraz: false,necesitaGarante: false,garanteNombre: "",docs: Object.fromEntries(DOC_ITEMS.map((d) => [d, { ok: false, archivo: null, nombreArchivo: "" }])),docsGarante: Object.fromEntries(DOC_ITEMS.map((d) => [d, { ok: false, archivo: null, nombreArchivo: "" }])),carpetaAprobada: false,},]);setInvitados((prev) => prev.filter((p) => p.id !== id));setTab("adjudicados");};const eliminarInvitado = (id) => setInvitados((prev) => prev.filter((p) => p.id !== id));const actualizarInvitado = (id, campo, valor) =>setInvitados((prev) => prev.map((p) => (p.id === id ? { ...p, [campo]: valor } : p)));const importarInvitados = (filas) => {const nuevos = filas.filter((f) => f.nombre).map((f) => ({ id: uid(), estado: "invitado", ...f }));setInvitados((prev) => [...prev, ...nuevos]);};/* ---- Adjudicados actions ---- */const agregarAdjudicadoDirecto = () => {if (!nuevoAdj.nombre.trim()) return;setAdjudicados((prev) => [...prev,{id: uid(),nombre: nuevoAdj.nombre.trim(),telefono: nuevoAdj.telefono.trim(),plan: nuevoAdj.plan,montoAdjudica: nuevoAdj.montoAdjudica,medioPagoAdjudicacion: nuevoAdj.medioPagoAdjudicacion,adjudicacionAdministrativa: nuevoAdj.adjudicacionAdministrativa,modelo: "",color: "",monto: "",vencimiento: "",pagado: false,pedidoRealizado: false,observaciones: "",veraz: false,necesitaGarante: false,garanteNombre: "",docs: Object.fromEntries(DOC_ITEMS.map((d) => [d, { ok: false, archivo: null, nombreArchivo: "" }])),docsGarante: Object.fromEntries(DOC_ITEMS.map((d) => [d, { ok: false, archivo: null, nombreArchivo: "" }])),carpetaAprobada: false,},]);setNuevoAdj({ nombre: "", telefono: "", plan: "", montoAdjudica: "", medioPagoAdjudicacion: "", adjudicacionAdministrativa: false });};const actualizarAdj = (id, campo, valor) => {setAdjudicados((prev) => prev.map((c) => (c.id === id ? { ...c, [campo]: valor } : c)));};const toggleDoc = (id, doc, campo = "docs") => {setAdjudicados((prev) =>prev.map((c) => {if (c.id !== id) return c;const actual = docInfo(c, doc, campo);return { ...c, [campo]: { ...c[campo], [doc]: { ...actual, ok: !actual.ok } } };}));};const adjuntarDoc = (id, doc, { archivo, nombreArchivo }, campo = "docs") => {setAdjudicados((prev) =>prev.map((c) => (c.id === id ? { ...c, [campo]: { ...c[campo], [doc]: { ok: true, archivo, nombreArchivo } } } : c)));};const quitarDocArchivo = (id, doc, campo = "docs") => {setAdjudicados((prev) =>prev.map((c) => (c.id === id ? { ...c, [campo]: { ...c[campo], [doc]: { ok: false, archivo: null, nombreArchivo: "" } } } : c)));};const eliminarAdj = (id) => setAdjudicados((prev) => prev.filter((c) => c.id !== id));const exportarAprobados = () => {const aprobados = adjudicados.filter((c) => c.carpetaAprobada);if (aprobados.length === 0) return;const filas = aprobados.map((c) => {const pct = calcularPorcentaje(c.montoAdjudica, valoresUnidades[baseModelo(c.plan)]);return {Nombre: c.nombre,WhatsApp: c.telefono,Grupo: c.grupo || "",Orden: c.orden || "",Plan: c.plan || "","Monto Adjudicación ($)": c.montoAdjudica || "","% real": pct !== null ? pct.toFixed(1) : "","Medio Pago Adjudicación": c.medioPagoAdjudicacion || "","Adj. Administrativa": c.adjudicacionAdministrativa ? "Sí" : "No",Modelo: c.modelo || "",Color: c.color || "","Monto Cuota": c.monto || "",Vencimiento: c.vencimiento || "",Pagado: c.pagado ? "Sí" : "No","Pedido SGA": c.pedidoRealizado ? "Sí" : "No",Observaciones: c.observaciones || "",};});const ws = XLSX.utils.json_to_sheet(filas);const wb = XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb, ws, "Aprobados");XLSX.writeFile(wb, adjudicados-aprobados-${new Date().toISOString().slice(0, 10)}.xlsx);};const contadores = {invitados: invitados.filter((i) => i.estado === "invitado").length,confirmados: invitados.filter((i) => i.estado === "confirmado").length,aprobados: adjudicados.filter((c) => c.carpetaAprobada).length,vencidos: adjudicados.filter((c) => paymentStatus(c) === "vencido").length,agencia: invitados.filter((i) => i.abonaEnAgencia).length + adjudicados.filter((a) => a.abonaEnAgencia).length,};return (<divstyle={{minHeight: "100vh",background: COLORS.bg,color: COLORS.text,fontFamily: "'Inter', system-ui, sans-serif",padding: "28px 18px 60px",}}>{@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500&display=swap'); * { box-sizing: border-box; } ::placeholder { color: ${COLORS.textMuted}; } table { border-collapse: collapse; width: 100%; }}  <div style={{ maxWidth: 1100, margin: "0 auto" }}>
    {/* Encabezado */}
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12, marginBottom: 16 }}>
      <div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ width: 10, height: 10, borderRadius: "50%", background: COLORS.amber, boxShadow: `0 0 12px ${COLORS.amber}` }} />
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 26, margin: 0 }}>
            Panel de Adjudicaciones WOCH
          </h1>
        </div>
        <p style={{ color: COLORS.textMuted, margin: "4px 0 0", fontSize: 13 }}>
          Gestión de convocados, cobro en agencia, aprobación de carpetas y empaquetado de documentos — Fiat Fadua
        </p>
      </div>

      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <button
          onClick={() => setMostrarHelpModal(true)}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            background: COLORS.panelAlt,
            border: `1px solid ${COLORS.border}`,
            color: COLORS.amber,
            borderRadius: 6,
            padding: "8px 12px",
            fontSize: 12,
            fontWeight: 600,
            cursor: "pointer"
          }}
        >
          <HardDrive size={14} /> Drive / SGA Info
        </button>
      </div>
    </div>

    {/* Tablero de métricas */}
    <div style={{ display: "flex", gap: 10, marginBottom: 22, flexWrap: "wrap" }}>
      {[
        ["Invitados pendientes", contadores.invitados, COLORS.textMuted],
        ["Confirmados", contadores.confirmados, COLORS.amber],
        ["Carpetas aprobadas", contadores.aprobados, COLORS.green],
        ["Cuotas vencidas", contadores.vencidos, COLORS.red],
        ["Abona en Agencia", contadores.agencia, COLORS.blue],
      ].map(([label, val, color]) => (
        <div
          key={label}
          style={{
            flex: "1 1 170px",
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

    {/* Pestañas de navegación */}
    <div style={{ display: "flex", gap: 6, marginBottom: 18, borderBottom: `1px solid ${COLORS.border}`, alignItems: "center" }}>
      {[
        ["invitados", `1 · Invitados (${invitados.length})`],
        ["adjudicados", `2 · Adjudicados (${adjudicados.length})`],
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
            padding: "8px 10px",
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
          display: "flex",
          alignItems: "center",
          gap: 4
        }}
      >
        <Settings size={14} /> Configuración de Valores
      </button>
    </div>

    {mostrarConfig && (
      <ConfigPanel
        planes={planes}
        onAgregarPlan={agregarPlan}
        onEliminarPlan={eliminarPlan}
        valoresUnidades={valoresUnidades}
        onActualizarValor={actualizarValorUnidad}
        onBackup={exportarBackupData}
        onRestore={restaurarBackupData}
      />
    )}

    {loading ? (
      <div style={{ color: COLORS.textMuted, fontSize: 13 }}>Cargando datos del panel...</div>
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
        onOpenWA={(c) => setWaClienteModal(c)}
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
        onVerDoc={(docObj) => setDocModalData(docObj)}
        onOpenWA={(c) => setWaClienteModal(c)}
        nombreCarpetaZip={nombreCarpetaZip}
        setNombreCarpetaZip={setNombreCarpetaZip}
      />
    )}
  </div>

  <DocModal docData={docModalData} onClose={() => setDocModalData(null)} />
  <WAModal cliente={waClienteModal} onClose={() => setWaClienteModal(null)} />
  {mostrarHelpModal && <IntegrationHelpModal onClose={() => setMostrarHelpModal(false)} />}
</div>
);}/* -------------------- Tab 1: Invitados -------------------- */function InvitadosTab({invitados,planes,valoresUnidades,nuevo,setNuevo,onAgregar,onConfirmar,onPasar,onEliminar,onActualizar,onImportar,onOpenWA,}) {const [busqueda, setBusqueda] = useState("");const [expandido, setExpandido] = useState(null);const fileRef = useRef(null);const manejarArchivo = (e) => {const file = e.target.files?.[0];if (!file) return;const reader = new FileReader();reader.onload = (ev) => {try {const wb = XLSX.read(ev.target.result, { type: "binary" });const hoja = wb.Sheets[wb.SheetNames[0]];const filas = XLSX.utils.sheet_to_json(hoja, { defval: "" });if (!filas.length) return;const headers = Object.keys(filas[0]);const mapeadas = filas.map((f) => mapearFilaImportada(f, headers));onImportar(mapeadas);} catch {onImportar([]);}};reader.readAsBinaryString(file);e.target.value = "";};const filtrados = invitados.filter((p) => {if (!busqueda.trim()) return true;const q = busqueda.trim().toLowerCase();return ((p.grupo || "").toLowerCase().includes(q) ||(p.orden || "").toLowerCase().includes(q) ||(p.nombre || "").toLowerCase().includes(q) ||(p.telefono || "").toLowerCase().includes(q));});return ({/* Alta de Invitado con nuevos campos Subite, Prorrateo, Abona en Agencia */}<divstyle={{background: COLORS.panel,border: 1px solid ${COLORS.border},borderRadius: 10,padding: 14,marginBottom: 16,display: "flex",flexDirection: "column",gap: 12}}><div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}><TextField value={nuevo.nombre} onChange={(v) => setNuevo((n) => ({ ...n, nombre: v }))} placeholder="Nombre del cliente" width={180} /><TextField value={nuevo.telefono} onChange={(v) => setNuevo((n) => ({ ...n, telefono: v }))} placeholder="(38)8517-0699 (c) (11)5665241..." mono width={220} /><TextField value={nuevo.grupo} onChange={(v) => setNuevo((n) => ({ ...n, grupo: v }))} placeholder="Grupo" mono width={75} /><TextField value={nuevo.orden} onChange={(v) => setNuevo((n) => ({ ...n, orden: v }))} placeholder="Orden" mono width={75} /><SelectField value={nuevo.plan} onChange={(v) => setNuevo((n) => ({ ...n, plan: v }))} options={planes} placeholder="Plan" width={130} /><TextField value={nuevo.montoAdjudica} onChange={(v) => setNuevo((n) => ({ ...n, montoAdjudica: v }))} placeholder="Monto lícito ($)" mono width={110} />    {/* Fila de Modales/Checks especiales: Subite, Prorratea, Agencia */}
    <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "center", background: COLORS.panelAlt, padding: "8px 12px", borderRadius: 6 }}>
      <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer", color: nuevo.subite ? COLORS.amber : "inherit" }}>
        <input type="checkbox" checked={nuevo.subite} onChange={(e) => setNuevo((n) => ({ ...n, subite: e.target.checked }))} />
        <strong>Plan Subite</strong> (Capital para invertir)
      </label>

      <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer", color: nuevo.prorratea ? COLORS.amber : "inherit" }}>
        <input type="checkbox" checked={nuevo.prorratea} onChange={(e) => setNuevo((n) => ({ ...n, prorratea: e.target.checked }))} />
        <strong>Prorratea</strong>
      </label>

      {nuevo.prorratea && (
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <TextField value={nuevo.porcentajeProrrateo} onChange={(v) => setNuevo((n) => ({ ...n, porcentajeProrrateo: v }))} placeholder="% Prorrateo" mono width={90} />
          <TextField value={nuevo.montoProrrateo} onChange={(v) => setNuevo((n) => ({ ...n, montoProrrateo: v }))} placeholder="$ Prorrateo" mono width={100} />
        </div>
      )}

      <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer", color: nuevo.abonaEnAgencia ? COLORS.green : "inherit" }}>
        <input type="checkbox" checked={nuevo.abonaEnAgencia} onChange={(e) => setNuevo((n) => ({ ...n, abonaEnAgencia: e.target.checked }))} />
        <strong>Abona en Agencia</strong>
      </label>

      <button
        onClick={onAgregar}
        style={{
          marginLeft: "auto",
          display: "inline-flex",
          alignItems: "center",
          gap: 6,
          background: COLORS.amber,
          color: "#1A1200",
          border: "none",
          borderRadius: 6,
          padding: "7px 14px",
          fontWeight: 700,
          fontSize: 13,
          cursor: "pointer",
        }}
      >
        <Plus size={14} /> Registrar Invitado
      </button>
    </div>
  </div>

  {/* Búsqueda e Importación */}
  <div style={{ marginBottom: 12, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
    <div style={{ position: "relative", flex: "1 1 240px" }}>
      <TextField value={busqueda} onChange={setBusqueda} placeholder="Buscar por grupo, orden, cliente o teléfono..." />
    </div>
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
      <Upload size={14} /> Importar Excel
    </button>
  </div>

  {/* Lista de Invitados */}
  {invitados.length === 0 ? (
    <EmptyState text="No hay invitados registrados. Suma un cliente arriba o importa un Excel." />
  ) : filtrados.length === 0 ? (
    <EmptyState text="No existen resultados para la búsqueda realizada." />
  ) : (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      {filtrados.map((p) => {
        const pct = calcularPorcentaje(p.montoAdjudica, valoresUnidades[baseModelo(p.plan)]);
        const abierto = expandido === p.id;
        const tels = extraerTelefonos(p.telefono);

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
                gap: 10,
                flexWrap: "wrap",
                cursor: "pointer",
              }}
              onClick={() => setExpandido(abierto ? null : p.id)}
            >
              <div style={{ flex: "1 1 180px", fontWeight: 600 }}>{p.nombre}</div>

              {(p.grupo || p.orden) && (
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.amber, background: COLORS.amberSoft, borderRadius: 6, padding: "2px 7px" }}>
                  G:{p.grupo || "—"} / O:{p.orden || "—"}
                </span>
              )}

              {/* Listado rápido de teléfonos detectados */}
              <div style={{ display: "flex", gap: 4, flexWrap: "wrap", alignItems: "center" }}>
                {tels.length > 0 ? (
                  tels.map((t, idx) => (
                    <a
                      key={idx}
                      href={waLink(t.clean, `Hola ${p.nombre}!`)}
                      target="_blank"
                      rel="noreferrer"
                      onClick={(e) => e.stopPropagation()}
                      style={{
                        fontSize: 11,
                        fontFamily: "'IBM Plex Mono', monospace",
                        background: COLORS.greenSoft,
                        color: COLORS.green,
                        border: `1px solid ${COLORS.green}44`,
                        padding: "2px 6px",
                        borderRadius: 4,
                        textDecoration: "none",
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 3
                      }}
                      title={`Enviar WhatsApp a ${t.clean}`}
                    >
                      <Phone size={10} /> {t.etiqueta}: {t.clean.slice(-8)}
                    </a>
                  ))
                ) : (
                  <span style={{ fontSize: 11, color: COLORS.textMuted }}>sin número</span>
                )}
              </div>

              <Pill color={p.estado === "confirmado" ? COLORS.amber : COLORS.textMuted} bg={p.estado === "confirmado" ? COLORS.amberSoft : COLORS.panelAlt}>
                {p.estado === "confirmado" ? "Confirmado" : "Pendiente"}
              </Pill>

              {p.subite && <Pill color={COLORS.purple} bg={COLORS.purpleSoft}>Subite</Pill>}
              {p.abonaEnAgencia && <Pill color={COLORS.blue} bg={COLORS.blueSoft}>Abona en Agencia</Pill>}

              {(p.plan || p.montoAdjudica) && (
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.textMuted }}>
                  {p.plan} {p.montoAdjudica && `· $${p.montoAdjudica}`} {pct !== null && `(${pct.toFixed(1)}%)`}
                </span>
              )}

              <div style={{ display: "flex", gap: 6, marginLeft: "auto", alignItems: "center" }} onClick={(e) => e.stopPropagation()}>
                <IconBtn title="Plantillas de WhatsApp" onClick={() => onOpenWA(p)} highlight>
                  <MessageSquare size={14} />
                </IconBtn>
                {p.estado !== "confirmado" ? (
                  <IconBtn title="Marcar como Confirmado" onClick={() => onConfirmar(p.id)}>
                    <Check size={14} color={COLORS.amber} />
                  </IconBtn>
                ) : (
                  <button
                    onClick={() => onPasar(p.id)}
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 6,
                      background: COLORS.amber,
                      color: "#1A1200",
                      border: "none",
                      borderRadius: 6,
                      padding: "5px 10px",
                      fontSize: 12,
                      fontWeight: 700,
                      cursor: "pointer",
                    }}
                  >
                    Pasar a Adjudicados <ArrowRight size={13} />
                  </button>
                )}
                <IconBtn title="Eliminar" danger onClick={() => onEliminar(p.id)}>
                  <Trash2 size={14} />
                </IconBtn>
                <button style={{ background: "none", border: "none", color: COLORS.textMuted, cursor: "pointer", padding: 2 }}>
                  {abierto ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                </button>
              </div>
            </div>

            {abierto && (
              <div
                style={{
                  padding: "12px 14px",
                  borderTop: `1px solid ${COLORS.border}`,
                  background: COLORS.panelAlt,
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
                  gap: 10,
                }}
              >
                <div>
                  <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Nombre completo</label>
                  <TextField value={p.nombre} onChange={(v) => onActualizar(p.id, "nombre", v)} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Teléfonos (varios)</label>
                  <TextField value={p.telefono} onChange={(v) => onActualizar(p.id, "telefono", v)} mono />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Grupo</label>
                  <TextField value={p.grupo} onChange={(v) => onActualizar(p.id, "grupo", v)} mono />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Orden</label>
                  <TextField value={p.orden} onChange={(v) => onActualizar(p.id, "orden", v)} mono />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Plan de ahorro</label>
                  <SelectField value={p.plan} onChange={(v) => onActualizar(p.id, "plan", v)} options={planes} placeholder="Seleccionar" />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Monto Adjudica ($)</label>
                  <TextField value={p.montoAdjudica} onChange={(v) => onActualizar(p.id, "montoAdjudica", v)} mono />
                </div>

                {/* Opciones especiales */}
                <div style={{ gridColumn: "1 / -1", display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap", paddingTop: 6, borderTop: `1px dashed ${COLORS.border}` }}>
                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}>
                    <input type="checkbox" checked={p.subite || false} onChange={(e) => onActualizar(p.id, "subite", e.target.checked)} />
                    <span>Plan Subite</span>
                  </label>

                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}>
                    <input type="checkbox" checked={p.prorratea || false} onChange={(e) => onActualizar(p.id, "prorratea", e.target.checked)} />
                    <span>Prorratea</span>
                  </label>

                  {p.prorratea && (
                    <>
                      <TextField value={p.porcentajeProrrateo || ""} onChange={(v) => onActualizar(p.id, "porcentajeProrrateo", v)} placeholder="% Prorrateo" mono width={100} />
                      <TextField value={p.montoProrrateo || ""} onChange={(v) => onActualizar(p.id, "montoProrrateo", v)} placeholder="$ Prorrateo" mono width={120} />
                    </>
                  )}

                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}>
                    <input type="checkbox" checked={p.abonaEnAgencia || false} onChange={(e) => onActualizar(p.id, "abonaEnAgencia", e.target.checked)} />
                    <span>Abona en Agencia</span>
                  </label>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  )}
</div>
);}/* -------------------- Tab 2: Adjudicados -------------------- */function AdjudicadosTab({adjudicados,planes,valoresUnidades,nuevo,setNuevo,onAgregar,onActualizar,onToggleDoc,onAdjuntarDoc,onQuitarDocArchivo,onEliminar,onExportar,onVerDoc,onOpenWA,nombreCarpetaZip,setNombreCarpetaZip}) {const [busqueda, setBusqueda] = useState("");const [filtroEstado, setFiltroEstado] = useState("todos");const [expandido, setExpandido] = useState(null);const filtrados = adjudicados.filter((c) => {const q = busqueda.trim().toLowerCase();const coincideTexto =!q ||(c.nombre || "").toLowerCase().includes(q) ||(c.grupo || "").toLowerCase().includes(q) ||(c.orden || "").toLowerCase().includes(q) ||(c.modelo || "").toLowerCase().includes(q);if (!coincideTexto) return false;

if (filtroEstado === "aprobados") return c.carpetaAprobada;
if (filtroEstado === "pendientes") return !c.carpetaAprobada;
if (filtroEstado === "vencidos") return paymentStatus(c) === "vencido";
if (filtroEstado === "sga") return c.pedidoRealizado;
if (filtroEstado === "agencia") return c.abonaEnAgencia || c.medioPagoAdjudicacion?.includes("Agencia");
if (filtroEstado === "admin") return c.adjudicacionAdministrativa;

return true;
});return ({/* Alta rápida directos */}<divstyle={{background: COLORS.panel,border: 1px solid ${COLORS.border},borderRadius: 10,padding: 14,marginBottom: 16,display: "flex",gap: 8,flexWrap: "wrap",alignItems: "center",}}><TextField value={nuevo.nombre} onChange={(v) => setNuevo((n) => ({ ...n, nombre: v }))} placeholder="Nombre cliente adjudicado" width={200} /><TextField value={nuevo.telefono} onChange={(v) => setNuevo((n) => ({ ...n, telefono: v }))} placeholder="WhatsApp/Teléfonos" mono width={160} /><SelectField value={nuevo.plan} onChange={(v) => setNuevo((n) => ({ ...n, plan: v }))} options={planes} placeholder="Plan" width={130} /><TextField value={nuevo.montoAdjudica} onChange={(v) => setNuevo((n) => ({ ...n, montoAdjudica: v }))} placeholder="Monto ($)" mono width={110} /><SelectField value={nuevo.medioPagoAdjudicacion} onChange={(v) => setNuevo((n) => ({ ...n, medioPagoAdjudicacion: v }))} options={MEDIOS_PAGO} placeholder="Medio de Pago Adjudicacion" width={180} />    <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}>
      <input type="checkbox" checked={nuevo.adjudicacionAdministrativa} onChange={(e) => setNuevo((n) => ({ ...n, adjudicacionAdministrativa: e.target.checked }))} />
      <span>Adj. Administrativa</span>
    </label>

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
        padding: "8px 14px",
        fontWeight: 700,
        fontSize: 13,
        cursor: "pointer",
        marginLeft: "auto"
      }}
    >
      <Plus size={14} /> Cargar Adjudicado
    </button>
  </div>

  {/* Acciones de descarga de documentos en carpetas y exportación */}
  <div style={{ background: COLORS.panelAlt, padding: 12, borderRadius: 8, border: `1px solid ${COLORS.border}`, marginBottom: 16, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center", justifyContent: "space-between" }}>
    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
      <FolderDownload size={18} color={COLORS.amber} />
      <span style={{ fontSize: 12, fontWeight: 600 }}>Descarga organizada de carpetas con documentos:</span>
      <TextField value={nombreCarpetaZip} onChange={setNombreCarpetaZip} placeholder="Nombre carpeta raíz (Ej: adj agosto)" width={160} />
      <button
        onClick={() => descargarCarpetasZip(nombreCarpetaZip, filtrados)}
        style={{
          display: "inline-flex",
          alignItems: "center",
          gap: 6,
          background: COLORS.amber,
          color: "#1A1200",
          border: "none",
          borderRadius: 6,
          padding: "7px 12px",
          fontSize: 12,
          fontWeight: 700,
          cursor: "pointer"
        }}
      >
        <FolderDownload size={14} /> Descargar Zip por Cliente
      </button>
    </div>

    <button
      onClick={onExportar}
      disabled={!adjudicados.some((c) => c.carpetaAprobada)}
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        background: adjudicados.some((c) => c.carpetaAprobada) ? COLORS.green : COLORS.panel,
        color: adjudicados.some((c) => c.carpetaAprobada) ? "#FFF" : COLORS.textMuted,
        border: `1px solid ${COLORS.border}`,
        borderRadius: 6,
        padding: "7px 12px",
        fontSize: 12,
        fontWeight: 600,
        cursor: adjudicados.some((c) => c.carpetaAprobada) ? "pointer" : "not-allowed",
      }}
    >
      <FileSpreadsheet size={14} /> Exportar Aprobados Excel
    </button>
  </div>

  {/* Controles de Búsqueda y Filtros */}
  <div style={{ marginBottom: 16, display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
    <TextField value={busqueda} onChange={setBusqueda} placeholder="Buscar por cliente, modelo, grupo..." width={240} />
    <select
      value={filtroEstado}
      onChange={(e) => setFiltroEstado(e.target.value)}
      style={{
        background: COLORS.bg,
        border: `1px solid ${COLORS.border}`,
        color: COLORS.text,
        borderRadius: 6,
        padding: "7px 10px",
        fontSize: 13
      }}
    >
      <option value="todos">Todos los adjudicados</option>
      <option value="aprobados">Carpetas Aprobadas</option>
      <option value="pendientes">Pendientes de Aprobación</option>
      <option value="agencia">Abona en Agencia / Medios Agencia</option>
      <option value="admin">Adjudicación Administrativa</option>
      <option value="vencidos">Cuota Vencida</option>
      <option value="sga">Pedido en SGA</option>
    </select>
  </div>

  {/* Tarjetas de Adjudicados */}
  {adjudicados.length === 0 ? (
    <EmptyState text="Todavía no hay clientes adjudicados registrados." />
  ) : filtrados.length === 0 ? (
    <EmptyState text="No hay registros que coincidan con los filtros aplicados." />
  ) : (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      {filtrados.map((c) => {
        const st = STATUS_STYLE[paymentStatus(c)];
        const pct = calcularPorcentaje(c.montoAdjudica, valoresUnidades[baseModelo(c.plan)]);
        const abierto = expandido === c.id;
        const tels = extraerTelefonos(c.telefono);

        const countDocs = (campo) => DOC_ITEMS.filter((d) => docInfo(c, d, campo).ok).length;
        const docsClienteOk = countDocs("docs");

        return (
          <div
            key={c.id}
            style={{
              background: COLORS.panel,
              border: `1px solid ${c.carpetaAprobada ? COLORS.green : COLORS.border}`,
              borderLeft: `4px solid ${c.carpetaAprobada ? COLORS.green : c.veraz ? COLORS.red : COLORS.amber}`,
              borderRadius: 10,
              overflow: "hidden",
            }}
          >
            {/* Cabecera de tarjeta */}
            <div
              style={{
                padding: "12px 16px",
                display: "flex",
                alignItems: "center",
                gap: 12,
                flexWrap: "wrap",
                cursor: "pointer",
              }}
              onClick={() => setExpandido(abierto ? null : c.id)}
            >
              <div style={{ flex: "1 1 220px" }}>
                <div style={{ fontWeight: 700, fontSize: 15, display: "flex", alignItems: "center", gap: 8 }}>
                  {c.nombre}
                  {c.carpetaAprobada && <ShieldCheck size={16} color={COLORS.green} title="Carpeta Aprobada" />}
                  {c.adjudicacionAdministrativa && <Briefcase size={15} color={COLORS.purple} title="Adjudicación Administrativa" />}
                  {c.veraz && <AlertTriangle size={16} color={COLORS.red} title="Afectado en Veraz" />}
                </div>
                <div style={{ fontSize: 12, color: COLORS.textMuted, display: "flex", gap: 10, marginTop: 2, flexWrap: "wrap" }}>
                  <span>G:{c.grupo || "—"} O:{c.orden || "—"}</span>
                  <span>·</span>
                  <span>{c.plan || "Sin plan"}</span>
                  {pct !== null && <span>({pct.toFixed(1)}%)</span>}
                </div>
              </div>

              {/* Listado de teléfonos detectados */}
              <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                {tels.map((t, idx) => (
                  <a
                    key={idx}
                    href={waLink(t.clean, `Hola ${c.nombre}!`)}
                    target="_blank"
                    rel="noreferrer"
                    onClick={(e) => e.stopPropagation()}
                    style={{
                      fontSize: 11,
                      fontFamily: "'IBM Plex Mono', monospace",
                      background: COLORS.greenSoft,
                      color: COLORS.green,
                      border: `1px solid ${COLORS.green}44`,
                      padding: "2px 6px",
                      borderRadius: 4,
                      textDecoration: "none",
                    }}
                  >
                    WA: {t.clean.slice(-8)}
                  </a>
                ))}
              </div>

              <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                {c.medioPagoAdjudicacion && (
                  <Pill color={COLORS.amber} bg={COLORS.amberSoft}>
                    <CreditCard size={11} /> {c.medioPagoAdjudicacion}
                  </Pill>
                )}

                <Pill color={st.color} bg={st.bg}>{st.label}</Pill>

                <Pill
                  color={docsClienteOk === DOC_ITEMS.length ? COLORS.green : COLORS.amber}
                  bg={docsClienteOk === DOC_ITEMS.length ? COLORS.greenSoft : COLORS.amberSoft}
                >
                  <FileCheck size={12} /> Docs: {docsClienteOk}/{DOC_ITEMS.length}
                </Pill>
              </div>

              <div style={{ display: "flex", gap: 6, marginLeft: "auto", alignItems: "center" }} onClick={(e) => e.stopPropagation()}>
                <IconBtn title="Plantilla WhatsApp" onClick={() => onOpenWA(c)} highlight>
                  <MessageSquare size={14} />
                </IconBtn>
                <IconBtn title="Eliminar" danger onClick={() => onEliminar(c.id)}>
                  <Trash2 size={14} />
                </IconBtn>
                <button style={{ background: "none", border: "none", color: COLORS.textMuted, cursor: "pointer", padding: 4 }}>
                  {abierto ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                </button>
              </div>
            </div>

            {/* Detalle desplegable */}
            {abierto && (
              <div style={{ borderTop: `1px solid ${COLORS.border}`, padding: 16, background: COLORS.panelAlt }}>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12, marginBottom: 16 }}>
                  <div>
                    <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Medio de Pago Adjudicación</label>
                    <SelectField value={c.medioPagoAdjudicacion || ""} onChange={(v) => onActualizar(c.id, "medioPagoAdjudicacion", v)} options={MEDIOS_PAGO} placeholder="Seleccionar Pago" />
                  </div>

                  <div>
                    <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Modelo asignado</label>
                    <TextField value={c.modelo} onChange={(v) => onActualizar(c.id, "modelo", v)} placeholder="Ej: Cronos Drive 1.3" />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Color elegido</label>
                    <TextField value={c.color} onChange={(v) => onActualizar(c.id, "color", v)} placeholder="Ej: Gris Silverstone" />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Monto Cuota ($)</label>
                    <TextField value={c.monto} onChange={(v) => onActualizar(c.id, "monto", v)} mono placeholder="0.00" />
                  </div>
                  <div>
                    <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Vencimiento Cuota</label>
                    <TextField type="date" value={c.vencimiento} onChange={(v) => onActualizar(c.id, "vencimiento", v)} mono />
                  </div>
                </div>

                {/* Checkboxes de Estado */}
                <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 18, background: COLORS.bg, padding: 12, borderRadius: 8, border: `1px solid ${COLORS.border}` }}>
                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer", color: c.adjudicacionAdministrativa ? COLORS.purple : "inherit" }}>
                    <input type="checkbox" checked={c.adjudicacionAdministrativa || false} onChange={(e) => onActualizar(c.id, "adjudicacionAdministrativa", e.target.checked)} />
                    <strong>Adjudicación Administrativa</strong>
                  </label>

                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}>
                    <input type="checkbox" checked={c.pagado} onChange={(e) => onActualizar(c.id, "pagado", e.target.checked)} />
                    <span>Cuota Pagada</span>
                  </label>

                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}>
                    <input type="checkbox" checked={c.pedidoRealizado} onChange={(e) => onActualizar(c.id, "pedidoRealizado", e.target.checked)} />
                    <span>Pedido Cargado en SGA</span>
                  </label>

                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer", color: c.veraz ? COLORS.red : "inherit" }}>
                    <input type="checkbox" checked={c.veraz} onChange={(e) => onActualizar(c.id, "veraz", e.target.checked)} />
                    <span>Afectado Veraz</span>
                  </label>

                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}>
                    <input type="checkbox" checked={c.necesitaGarante} onChange={(e) => onActualizar(c.id, "necesitaGarante", e.target.checked)} />
                    <span>Requiere Garante</span>
                  </label>

                  <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 13, fontWeight: 700, cursor: "pointer", color: c.carpetaAprobada ? COLORS.green : COLORS.amber, marginLeft: "auto" }}>
                    <input type="checkbox" checked={c.carpetaAprobada} onChange={(e) => onActualizar(c.id, "carpetaAprobada", e.target.checked)} />
                    <span>CARPETA APROBADA</span>
                  </label>
                </div>

                {/* Documentación Titular */}
                <div style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.textMuted, marginBottom: 8 }}>Documentación Titular</div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 8 }}>
                    {DOC_ITEMS.map((doc) => (
                      <DocUploaderBlock
                        key={doc}
                        label={doc}
                        info={docInfo(c, doc, "docs")}
                        onToggle={() => onToggleDoc(c.id, doc, "docs")}
                        onUpload={(data) => onAdjuntarDoc(c.id, doc, data, "docs")}
                        onQuitar={() => onQuitarDocArchivo(c.id, doc, "docs")}
                        onVer={() => onVerDoc(docInfo(c, doc, "docs"))}
                      />
                    ))}
                  </div>
                </div>

                {/* Documentación Garante */}
                {c.necesitaGarante && (
                  <div style={{ marginBottom: 16, background: COLORS.panel, padding: 12, borderRadius: 8, border: `1px dashed ${COLORS.border}` }}>
                    <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 10 }}>
                      <span style={{ fontSize: 12, fontWeight: 600, color: COLORS.amber }}>Documentación del Garante</span>
                      <TextField
                        value={c.garanteNombre || ""}
                        onChange={(v) => onActualizar(c.id, "garanteNombre", v)}
                        placeholder="Nombre del Garante"
                        width={200}
                      />
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 8 }}>
                      {DOC_ITEMS.map((doc) => (
                        <DocUploaderBlock
                          key={doc}
                          label={doc}
                          info={docInfo(c, doc, "docsGarante")}
                          onToggle={() => onToggleDoc(c.id, doc, "docsGarante")}
                          onUpload={(data) => onAdjuntarDoc(c.id, doc, data, "docsGarante")}
                          onQuitar={() => onQuitarDocArchivo(c.id, doc, "docsGarante")}
                          onVer={() => onVerDoc(docInfo(c, doc, "docsGarante"))}
                        />
                      ))}
                    </div>
                  </div>
                )}

                {/* Observaciones */}
                <div>
                  <label style={{ fontSize: 11, color: COLORS.textMuted, display: "block", marginBottom: 2 }}>Observaciones generales</label>
                  <TextField value={c.observaciones || ""} onChange={(v) => onActualizar(c.id, "observaciones", v)} placeholder="Anotaciones de la adjudicación..." />
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  )}
</div>
);}/* Subcomponente para Carga y Vista de Documentos */function DocUploaderBlock({ label, info, onToggle, onUpload, onQuitar, onVer }) {const fileRef = useRef(null);const handleChange = (e) => {const file = e.target.files?.[0];if (!file) return;comprimirImagen(file, (res) => {onUpload(res);});e.target.value = "";};return (<divstyle={{background: COLORS.bg,border: 1px solid ${info.ok ? COLORS.green + "66" : COLORS.border},borderRadius: 6,padding: "6px 8px",display: "flex",alignItems: "center",justifyContent: "space-between",gap: 6}}><div style={{ display: "flex", alignItems: "center", gap: 6, overflow: "hidden" }}><input type="checkbox" checked={info.ok} onChange={onToggle} style={{ cursor: "pointer" }} /><span style={{ fontSize: 12, color: info.ok ? COLORS.text : COLORS.textMuted, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{label}  <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
    {info.archivo && (
      <button onClick={onVer} title="Ver vista previa de imagen" style={{ background: "none", border: "none", color: COLORS.amber, cursor: "pointer", padding: 2 }}>
        <Eye size={13} />
      </button>
    )}

    <input ref={fileRef} type="file" accept="image/*,.pdf" onChange={handleChange} style={{ display: "none" }} />

    <button
      onClick={() => fileRef.current?.click()}
      title={info.nombreArchivo ? info.nombreArchivo : "Adjuntar archivo"}
      style={{ background: "none", border: "none", color: info.nombreArchivo ? COLORS.green : COLORS.textMuted, cursor: "pointer", padding: 2 }}
    >
      <Upload size={13} />
    </button>

    {info.nombreArchivo && (
      <button onClick={onQuitar} title="Quitar archivo adjunto" style={{ background: "none", border: "none", color: COLORS.red, cursor: "pointer", padding: 2 }}>
        <X size={13} />
      </button>
    )}
  </div>
</div>
);}
