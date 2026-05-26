import { useState, useEffect, useCallback } from "react";
import { supabase } from "./supabase";
import * as XLSX from "xlsx";

const MONTHS = ["Ene","Feb","Mar","Apr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];

function fmt(n) { return Number(n).toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function fmtARS(n) { return "$ " + fmt(n); }
function fmtUSD(n) { return "U$S " + fmt(n); }
function getMonthLabel(dateStr) {
  const d = new Date(dateStr + "T12:00:00");
  return MONTHS[d.getMonth()] + " " + d.getFullYear();
}

function exportToExcel(operaciones, cuentas, movimientos) {
  const wb = XLSX.utils.book_new();

  // Operaciones
  const dataOps = operaciones.map(o => ({
    "Persona": o.persona, "Banco": o.banco, "Fecha": o.fecha,
    "USD Comprado": o.usd, "TC Compra": o.tc, "ARS Invertido": o.ars,
    "Estado": o.tc_venta ? "VENDIDO" : "EN STOCK",
    "USD Vendido": o.usd_vendido || "", "TC Venta": o.tc_venta || "",
    "Fecha Venta": o.fecha_venta || "",
    "Ganancia ARS": o.tc_venta ? ((o.usd_vendido || o.usd) * (o.tc_venta - o.tc)).toFixed(2) : "",
  }));
  const ws1 = XLSX.utils.json_to_sheet(dataOps);
  ws1["!cols"] = [14,12,12,14,12,16,12,12,12,12,14].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws1, "Operaciones");

  // Cuentas
  const dataCuentas = cuentas.map(c => ({
    "Cuenta": c.nombre, "Saldo ARS": c.saldo_ars, "USD en cuenta": c.saldo_usd, "Efectivo USD": c.efectivo_usd || 0
  }));
  const ws2 = XLSX.utils.json_to_sheet(dataCuentas);
  XLSX.utils.book_append_sheet(wb, ws2, "Cuentas");

  // Movimientos
  const dataMov = movimientos.map(m => ({
    "Fecha": m.fecha, "Tipo": m.tipo, "Cuenta": m.cuenta,
    "ARS": m.ars || "", "USD": m.usd || "", "Nota": m.nota || ""
  }));
  const ws3 = XLSX.utils.json_to_sheet(dataMov);
  XLSX.utils.book_append_sheet(wb, ws3, "Movimientos");

  XLSX.writeFile(wb, `USD_Tracker_${new Date().toISOString().slice(0,10)}.xlsx`);
}

export default function App() {
  const [personas, setPersonas] = useState([]);
  const [operaciones, setOperaciones] = useState([]);
  const [cuentas, setCuentas] = useState([]);
  const [movimientos, setMovimientos] = useState([]);
  const [loading, setLoading] = useState(true);
  const [syncing, setSyncing] = useState(false);
  const [view, setView] = useState("dashboard");

  // Forms
  const [form, setForm] = useState({ persona: "", banco: "", usd: "", tc: "", fecha: new Date().toISOString().slice(0,10), nota: "" });
  const [ventaSelected, setVentaSelected] = useState([]);
  const [tcVenta, setTcVenta] = useState("");
  const [fechaVenta, setFechaVenta] = useState(new Date().toISOString().slice(0,10));
  const [cuentaVenta, setCuentaVenta] = useState("");
  const [newPersona, setNewPersona] = useState("");
  const [newCuenta, setNewCuenta] = useState("");
  const [filterPersona, setFilterPersona] = useState("todas");
  const [filterMes, setFilterMes] = useState("todos");
  const [toast, setToast] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);

  // Movimiento forms
  const [movForm, setMovForm] = useState({
    tipo: "deposito_ars", cuenta: "", ars: "", usd: "",
    fecha: new Date().toISOString().slice(0,10), nota: ""
  });

  function showToast(msg, type = "ok") {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  }

  const loadData = useCallback(async () => {
    setLoading(true);
    try {
      const [{ data: p }, { data: o }, { data: c }, { data: m }] = await Promise.all([
        supabase.from("personas").select("*").order("nombre"),
        supabase.from("operaciones").select("*").order("created_at", { ascending: false }),
        supabase.from("cuentas").select("*").order("nombre"),
        supabase.from("movimientos").select("*").order("fecha", { ascending: false }),
      ]);
      setPersonas((p || []).map(x => x.nombre));
      setOperaciones(o || []);
      setCuentas(c || []);
      setMovimientos(m || []);
    } catch (e) {
      showToast("Error al cargar datos", "err");
    }
    setLoading(false);
  }, []);

  useEffect(() => { loadData(); }, [loadData]);

  useEffect(() => {
    const ch1 = supabase.channel("op-changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "operaciones" }, () => loadData())
      .subscribe();
    const ch2 = supabase.channel("cu-changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "cuentas" }, () => loadData())
      .subscribe();
    const ch3 = supabase.channel("mo-changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "movimientos" }, () => loadData())
      .subscribe();
    return () => { supabase.removeChannel(ch1); supabase.removeChannel(ch2); supabase.removeChannel(ch3); };
  }, [loadData]);

  // ── Personas ──
  async function addPersona() {
    const nombre = newPersona.trim();
    if (!nombre) return;
    if (personas.includes(nombre)) { showToast("Ya existe", "err"); return; }
    setSyncing(true);
    await supabase.from("personas").insert({ nombre });
    setSyncing(false);
    setNewPersona("");
    showToast("Persona agregada ✓");
  }

  // ── Cuentas ──
  async function addCuenta() {
    const nombre = newCuenta.trim();
    if (!nombre) return;
    if (cuentas.some(c => c.nombre === nombre)) { showToast("Ya existe", "err"); return; }
    setSyncing(true);
    await supabase.from("cuentas").insert({ nombre, saldo_ars: 0, saldo_usd: 0, efectivo_usd: 0 });
    setSyncing(false);
    setNewCuenta("");
    showToast("Cuenta agregada ✓");
  }

  async function deleteCuenta(id) {
    setSyncing(true);
    await supabase.from("cuentas").delete().eq("id", id);
    setSyncing(false);
    showToast("Cuenta eliminada");
  }

  // ── Movimientos (depósito ARS, retiro cajero, ingreso ARS por venta) ──
  async function submitMovimiento() {
    const { tipo, cuenta, ars, usd, fecha, nota } = movForm;
    if (!cuenta || !fecha) { showToast("Completá cuenta y fecha", "err"); return; }

    const cuentaObj = cuentas.find(c => c.nombre === cuenta);
    if (!cuentaObj) return;

    let updateData = {};
    let movData = { tipo, cuenta, fecha, nota: nota || null, ars: null, usd: null };

    if (tipo === "deposito_ars") {
      if (!ars) { showToast("Ingresá el monto ARS", "err"); return; }
      updateData = { saldo_ars: cuentaObj.saldo_ars + parseFloat(ars) };
      movData.ars = parseFloat(ars);
    } else if (tipo === "retiro_cajero") {
      if (!usd) { showToast("Ingresá los USD", "err"); return; }
      updateData = {
        saldo_usd: cuentaObj.saldo_usd - parseFloat(usd),
        efectivo_usd: (cuentaObj.efectivo_usd || 0) + parseFloat(usd)
      };
      movData.usd = parseFloat(usd);
    } else if (tipo === "deposito_usd") {
      if (!usd) { showToast("Ingresá los USD", "err"); return; }
      updateData = { saldo_usd: cuentaObj.saldo_usd + parseFloat(usd) };
      movData.usd = parseFloat(usd);
    } else if (tipo === "retiro_ars") {
      if (!ars) { showToast("Ingresá el monto ARS", "err"); return; }
      updateData = { saldo_ars: cuentaObj.saldo_ars - parseFloat(ars) };
      movData.ars = parseFloat(ars);
    }

    setSyncing(true);
    await supabase.from("cuentas").update(updateData).eq("id", cuentaObj.id);
    await supabase.from("movimientos").insert(movData);
    setSyncing(false);
    setMovForm({ ...movForm, ars: "", usd: "", nota: "" });
    showToast("Movimiento registrado ✓");
  }

  // ── Compra ──
  async function submitOp() {
    const { persona, banco, usd, tc, fecha, nota } = form;
    if (!persona || !banco || !usd || !tc || !fecha) { showToast("Completá todos los campos", "err"); return; }

    const cuentaObj = cuentas.find(c => c.nombre === banco);
    const arsTotal = parseFloat(usd) * parseFloat(tc);

    setSyncing(true);
    await supabase.from("operaciones").insert({
      persona, banco, usd: parseFloat(usd), tc: parseFloat(tc),
      ars: arsTotal, fecha, nota: nota || null,
      tc_venta: null, fecha_venta: null, usd_vendido: null, lote_id: null,
    });

    // Actualizar saldo cuenta: descontar ARS, sumar USD
    if (cuentaObj) {
      await supabase.from("cuentas").update({
        saldo_ars: cuentaObj.saldo_ars - arsTotal,
        saldo_usd: cuentaObj.saldo_usd + parseFloat(usd),
      }).eq("id", cuentaObj.id);
    }

    await supabase.from("movimientos").insert({
      tipo: "compra", cuenta: banco, fecha,
      ars: -arsTotal, usd: parseFloat(usd),
      nota: `Compra para ${persona}${nota ? " - " + nota : ""}`
    });

    setSyncing(false);
    setForm({ ...form, usd: "", nota: "" });
    showToast("Compra registrada ✓");
    setView("historial");
  }

  // ── Venta ──
  const availableForSale = operaciones.filter(o => !o.tc_venta);
  function isSelected(id) { return ventaSelected.some(x => x.id === id); }
  function toggleSelect(op) {
    if (isSelected(op.id)) setVentaSelected(prev => prev.filter(x => x.id !== op.id));
    else setVentaSelected(prev => [...prev, { id: op.id, usdParcial: op.usd }]);
  }
  function setParcial(id, val) {
    setVentaSelected(prev => prev.map(x => x.id === id ? { ...x, usdParcial: val } : x));
  }
  function selectAll() {
    if (ventaSelected.length === availableForSale.length) setVentaSelected([]);
    else setVentaSelected(availableForSale.map(o => ({ id: o.id, usdParcial: o.usd })));
  }

  const selectedItems = ventaSelected.map(x => {
    const op = operaciones.find(o => o.id === x.id);
    const usdV = Math.min(parseFloat(x.usdParcial) || 0, op ? op.usd : 0);
    return { op, usdV };
  }).filter(x => x.op && x.usdV > 0);

  const totalUSDLote = selectedItems.reduce((s, x) => s + x.usdV, 0);
  const totalARSLote = selectedItems.reduce((s, x) => s + x.usdV * x.op.tc, 0);
  const tcPromLote = totalUSDLote > 0 ? totalARSLote / totalUSDLote : 0;
  const gananciaLote = tcVenta && totalUSDLote > 0
    ? totalUSDLote * (parseFloat(tcVenta) - tcPromLote) : null;

  async function submitVenta() {
    if (selectedItems.length === 0) { showToast("Seleccioná al menos una compra", "err"); return; }
    if (!tcVenta || !fechaVenta) { showToast("Completá TC y fecha", "err"); return; }
    if (!cuentaVenta) { showToast("Seleccioná la cuenta donde ingresa el ARS", "err"); return; }

    const tc = parseFloat(tcVenta);
    const loteId = "lote-" + Date.now();
    const cuentaObj = cuentas.find(c => c.nombre === cuentaVenta);
    let totalARSIngresado = 0;
    let totalUSDVendido = 0;

    setSyncing(true);
    for (const { op, usdV } of selectedItems) {
      const esParcial = Math.abs(usdV - op.usd) > 0.001;
      await supabase.from("operaciones").update({
        tc_venta: tc, fecha_venta: fechaVenta,
        usd_vendido: usdV, lote_id: loteId,
      }).eq("id", op.id);

      if (esParcial) {
        await supabase.from("operaciones").insert({
          persona: op.persona, banco: op.banco,
          usd: op.usd - usdV, tc: op.tc,
          ars: (op.usd - usdV) * op.tc,
          fecha: op.fecha, nota: op.nota ? op.nota + " (resto)" : "resto",
          tc_venta: null, fecha_venta: null, usd_vendido: null, lote_id: null,
        });
      }

      // Descontar USD de la cuenta origen
      const cuentaOrigen = cuentas.find(c => c.nombre === op.banco);
      if (cuentaOrigen) {
        await supabase.from("cuentas").update({
          saldo_usd: Math.max(0, cuentaOrigen.saldo_usd - usdV),
        }).eq("id", cuentaOrigen.id);
      }

      totalARSIngresado += usdV * tc;
      totalUSDVendido += usdV;
    }

    // Sumar ARS a la cuenta de venta
    if (cuentaObj) {
      await supabase.from("cuentas").update({
        saldo_ars: cuentaObj.saldo_ars + totalARSIngresado,
      }).eq("id", cuentaObj.id);
    }

    await supabase.from("movimientos").insert({
      tipo: "venta", cuenta: cuentaVenta, fecha: fechaVenta,
      ars: totalARSIngresado, usd: -totalUSDVendido,
      nota: `Venta ${fmtUSD(totalUSDVendido)} @ $${tc}`
    });

    setSyncing(false);
    const ganMsg = gananciaLote !== null ? ` · Ganancia: ${fmtARS(gananciaLote)}` : "";
    setVentaSelected([]);
    setTcVenta("");
    setCuentaVenta("");
    showToast(`Venta registrada ✓${ganMsg}`);
    setView("dashboard");
  }

  async function deleteOp(id) {
    setSyncing(true);
    await supabase.from("operaciones").delete().eq("id", id);
    setSyncing(false);
    setConfirmDel(null);
    showToast("Eliminado");
  }

  // ── Stats ──
  const stockOps = operaciones.filter(o => !o.tc_venta);
  const stockUSD = stockOps.reduce((s, o) => s + o.usd, 0);
  const stockARS = stockOps.reduce((s, o) => s + o.ars, 0);
  const tcPromStock = stockUSD > 0 ? stockARS / stockUSD : 0;
  const totalARSInvertido = operaciones.reduce((s, o) => s + o.ars, 0);
  const gananciaTotal = operaciones.filter(o => o.tc_venta)
    .reduce((s, o) => s + (o.usd_vendido || o.usd) * (o.tc_venta - o.tc), 0);
  const totalARSDisponible = cuentas.reduce((s, c) => s + c.saldo_ars, 0);
  const totalEfectivoUSD = cuentas.reduce((s, c) => s + (c.efectivo_usd || 0), 0);

  const meses = [...new Set(operaciones.map(o => getMonthLabel(o.fecha)))];
  const filteredOps = operaciones.filter(o => {
    const mp = filterPersona === "todas" || o.persona === filterPersona;
    const mm = filterMes === "todos" || getMonthLabel(o.fecha) === filterMes;
    return mp && mm;
  });

  function personaSummary(nombre) {
    const po = operaciones.filter(o => o.persona === nombre);
    const totalU = po.reduce((s, o) => s + o.usd, 0);
    const totalA = po.reduce((s, o) => s + o.ars, 0);
    const stock = po.filter(o => !o.tc_venta).reduce((s, o) => s + o.usd, 0);
    const ganancia = po.filter(o => o.tc_venta).reduce((s, o) => s + (o.usd_vendido || o.usd) * (o.tc_venta - o.tc), 0);
    const tcP = totalU > 0 ? totalA / totalU : 0;
    return { totalU, stock, ganancia, tcP };
  }

  const nav = [
    { id: "dashboard", label: "Panel" },
    { id: "cuentas", label: "Cuentas" },
    { id: "movimientos", label: "Mover" },
    { id: "nueva", label: "+ Compra" },
    { id: "venta", label: "Vender" },
    { id: "historial", label: "Historial" },
    { id: "personas", label: "Personas" },
  ];

  const card = { background: "#12121a", border: "1px solid #1e1e2e", borderRadius: 10, padding: "14px 16px", marginBottom: 8 };
  const lbl = { fontSize: 10, color: "#555", marginBottom: 4, letterSpacing: 1 };
  const MOV_TIPOS = [
    { id: "deposito_ars", label: "💰 Depósito ARS", desc: "Ingresás pesos a una cuenta" },
    { id: "retiro_cajero", label: "🏧 Retiro cajero", desc: "Sacás USD físicos de una cuenta" },
    { id: "deposito_usd", label: "💵 Depósito USD", desc: "Ingresás USD a una cuenta" },
    { id: "retiro_ars", label: "📤 Retiro ARS", desc: "Sacás pesos de una cuenta" },
  ];

  if (loading) return (
    <div style={{ minHeight: "100vh", background: "#0a0a0f", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16 }}>
      <div style={{ fontFamily: "sans-serif", fontSize: 22, fontWeight: 900, color: "#2a9d5c" }}>USD<span style={{ color: "#e8e0d0", fontWeight: 400 }}>TRACKER</span></div>
      <div style={{ fontSize: 12, color: "#444", fontFamily: "monospace" }}>Conectando...</div>
    </div>
  );

  return (
    <div style={{ minHeight: "100vh", background: "#0a0a0f", color: "#e8e0d0", fontFamily: "'IBM Plex Mono', monospace" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=Unbounded:wght@400;700;900&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #0a0a0f; }
        ::-webkit-scrollbar { width: 4px; } ::-webkit-scrollbar-track { background: #111; } ::-webkit-scrollbar-thumb { background: #2a6e4a; }
        input, select { background: #12121a; border: 1px solid #2a2a3a; color: #e8e0d0; padding: 10px 12px; border-radius: 6px; font-family: 'IBM Plex Mono', monospace; font-size: 13px; width: 100%; outline: none; transition: border .2s; }
        input:focus, select:focus { border-color: #2a9d5c; }
        input::placeholder { color: #444; }
        button { cursor: pointer; font-family: 'IBM Plex Mono', monospace; }
        .chk { width:18px; height:18px; accent-color:#2a9d5c; cursor:pointer; flex-shrink:0; }
        .parcial-input { background:#0a1a10 !important; border-color:#2a9d5c !important; color:#2a9d5c !important; width:110px !important; padding:6px 8px !important; font-size:12px !important; }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
      `}</style>

      {/* Header */}
      <div style={{ background: "#0d0d14", borderBottom: "1px solid #1a1a2a", padding: "12px 16px", display: "flex", alignItems: "center", gap: 10, position: "sticky", top: 0, zIndex: 10 }}>
        <div style={{ fontFamily: "'Unbounded'", fontSize: 15, fontWeight: 900, color: "#2a9d5c" }}>USD</div>
        <div style={{ fontFamily: "'Unbounded'", fontSize: 15, fontWeight: 400 }}>TRACKER</div>
        {syncing && <div style={{ fontSize: 10, color: "#2a9d5c", animation: "pulse 1s infinite" }}>● sync</div>}
        <div style={{ marginLeft: "auto" }}>
          <button onClick={() => exportToExcel(operaciones, cuentas, movimientos)} style={{
            background: "#0d2e1a", border: "1px solid #1a4d2e", color: "#2a9d5c",
            borderRadius: 6, padding: "6px 12px", fontSize: 11, fontWeight: 700
          }}>⬇ Excel</button>
        </div>
      </div>

      {/* Nav */}
      <div style={{ display: "flex", gap: 2, padding: "6px 8px", background: "#0d0d14", borderBottom: "1px solid #1a1a2a", overflowX: "auto", position: "sticky", top: "45px", zIndex: 9 }}>
        {nav.map(n => (
          <button key={n.id} onClick={() => setView(n.id)} style={{
            padding: "6px 11px", borderRadius: 6, border: "none", fontSize: 10, fontWeight: 600,
            background: view === n.id ? "#2a9d5c" : "transparent",
            color: view === n.id ? "#fff" : "#555", whiteSpace: "nowrap",
          }}>{n.label}</button>
        ))}
      </div>

      <div style={{ padding: "16px 14px", maxWidth: 680, margin: "0 auto", paddingBottom: 40 }}>

        {/* DASHBOARD */}
        {view === "dashboard" && (
          <div>
            <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 12, letterSpacing: 2 }}>RESUMEN GENERAL</div>

            {/* Liquidez */}
            <div style={{ background: "#0a1a10", border: "1px solid #1a4d2e", borderRadius: 12, padding: "14px 16px", marginBottom: 14 }}>
              <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#2a9d5c", marginBottom: 12, letterSpacing: 2 }}>💰 LIQUIDEZ</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <div>
                  <div style={{ fontSize: 10, color: "#555", marginBottom: 3 }}>ARS DISPONIBLE</div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: "#2a9d5c" }}>{fmtARS(totalARSDisponible)}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10, color: "#555", marginBottom: 3 }}>USD EFECTIVO FÍSICO</div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: "#e07a3a" }}>{fmtUSD(totalEfectivoUSD)}</div>
                </div>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 14 }}>
              {[
                { label: "STOCK USD EN CUENTAS", value: fmtUSD(stockUSD), accent: "#2a9d5c" },
                { label: "TC PROM. STOCK", value: "$ " + fmt(tcPromStock), accent: "#6a8fff" },
                { label: "ARS INVERTIDO TOTAL", value: fmtARS(totalARSInvertido), accent: "#aaa" },
                { label: "GANANCIA REALIZADA", value: fmtARS(gananciaTotal), accent: gananciaTotal >= 0 ? "#2a9d5c" : "#e05a5a" },
              ].map(c => (
                <div key={c.label} style={card}>
                  <div style={{ ...lbl, marginBottom: 6 }}>{c.label}</div>
                  <div style={{ fontSize: 14, fontWeight: 600, color: c.accent }}>{c.value}</div>
                </div>
              ))}
            </div>

            {/* Cuentas resumen */}
            {cuentas.length > 0 && (
              <>
                <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 10, letterSpacing: 2 }}>ESTADO DE CUENTAS</div>
                {cuentas.map(c => (
                  <div key={c.id} style={{ ...card, borderColor: "#1a2e1a" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                      <div style={{ fontWeight: 600, fontSize: 14 }}>{c.nombre}</div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6 }}>
                      <div style={{ fontSize: 10, color: "#555" }}>ARS<br/><span style={{ color: c.saldo_ars >= 0 ? "#2a9d5c" : "#e05a5a", fontSize: 13, fontWeight: 600 }}>{fmtARS(c.saldo_ars)}</span></div>
                      <div style={{ fontSize: 10, color: "#555" }}>USD CUENTA<br/><span style={{ color: "#6a8fff", fontSize: 13, fontWeight: 600 }}>{fmtUSD(c.saldo_usd)}</span></div>
                      <div style={{ fontSize: 10, color: "#555" }}>USD EFECTIVO<br/><span style={{ color: "#e07a3a", fontSize: 13, fontWeight: 600 }}>{fmtUSD(c.efectivo_usd || 0)}</span></div>
                    </div>
                  </div>
                ))}
              </>
            )}

            {/* Por persona */}
            <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 10, marginTop: 14, letterSpacing: 2 }}>POR PERSONA</div>
            {personas.length === 0 && <div style={{ color: "#444", fontSize: 12 }}>Sin personas.</div>}
            {personas.map(p => {
              const s = personaSummary(p);
              return (
                <div key={p} style={card}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                    <div style={{ fontWeight: 600, fontSize: 13 }}>{p}</div>
                    <div style={{ fontSize: 12, color: "#2a9d5c", fontWeight: 600 }}>Stock: {fmtUSD(s.stock)}</div>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6 }}>
                    <div style={{ fontSize: 10, color: "#555" }}>TOTAL USD<br/><span style={{ color: "#e8e0d0", fontSize: 12 }}>{fmtUSD(s.totalU)}</span></div>
                    <div style={{ fontSize: 10, color: "#555" }}>TC PROM.<br/><span style={{ color: "#6a8fff", fontSize: 12 }}>$ {fmt(s.tcP)}</span></div>
                    <div style={{ fontSize: 10, color: "#555" }}>GANANCIA<br/><span style={{ color: s.ganancia >= 0 ? "#2a9d5c" : "#e05a5a", fontSize: 12 }}>{fmtARS(s.ganancia)}</span></div>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* CUENTAS */}
        {view === "cuentas" && (
          <div>
            <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 16, letterSpacing: 2 }}>CUENTAS</div>
            <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
              <input placeholder="Nombre cuenta (ej: Brubank, Galicia...)" value={newCuenta} onChange={e => setNewCuenta(e.target.value)} onKeyDown={e => e.key === "Enter" && addCuenta()} />
              <button onClick={addCuenta} style={{ background: "#2a9d5c", color: "#fff", border: "none", borderRadius: 6, padding: "0 14px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>+ Agregar</button>
            </div>
            {cuentas.length === 0 && <div style={{ color: "#444", fontSize: 12 }}>Sin cuentas. Agregá Brubank, Galicia, Efectivo...</div>}
            {cuentas.map(c => (
              <div key={c.id} style={{ ...card, borderColor: "#1a2e1a" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <div style={{ fontWeight: 600, fontSize: 14 }}>{c.nombre}</div>
                  <button onClick={() => deleteCuenta(c.id)} style={{ background: "none", border: "1px solid #2a2a3a", color: "#555", borderRadius: 6, padding: "4px 8px", fontSize: 10 }}>Eliminar</button>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                  <div>
                    <div style={{ fontSize: 10, color: "#555", marginBottom: 3 }}>ARS DISPONIBLE</div>
                    <div style={{ fontSize: 15, fontWeight: 700, color: c.saldo_ars >= 0 ? "#2a9d5c" : "#e05a5a" }}>{fmtARS(c.saldo_ars)}</div>
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: "#555", marginBottom: 3 }}>USD EN CUENTA</div>
                    <div style={{ fontSize: 15, fontWeight: 700, color: "#6a8fff" }}>{fmtUSD(c.saldo_usd)}</div>
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: "#555", marginBottom: 3 }}>USD EFECTIVO</div>
                    <div style={{ fontSize: 15, fontWeight: 700, color: "#e07a3a" }}>{fmtUSD(c.efectivo_usd || 0)}</div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* MOVIMIENTOS */}
        {view === "movimientos" && (
          <div>
            <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 16, letterSpacing: 2 }}>REGISTRAR MOVIMIENTO</div>

            {/* Tipo */}
            <div style={{ marginBottom: 14 }}>
              <div style={lbl}>TIPO DE MOVIMIENTO</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                {MOV_TIPOS.map(t => (
                  <button key={t.id} onClick={() => setMovForm({...movForm, tipo: t.id})} style={{
                    padding: "10px 12px", borderRadius: 8, border: "1px solid",
                    borderColor: movForm.tipo === t.id ? "#2a9d5c" : "#2a2a3a",
                    background: movForm.tipo === t.id ? "#0d2e1a" : "transparent",
                    color: movForm.tipo === t.id ? "#2a9d5c" : "#555",
                    fontSize: 11, fontWeight: 600, textAlign: "left"
                  }}>
                    <div>{t.label}</div>
                    <div style={{ fontSize: 9, fontWeight: 400, marginTop: 2, color: movForm.tipo === t.id ? "#2a9d5c" : "#444" }}>{t.desc}</div>
                  </button>
                ))}
              </div>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div>
                <div style={lbl}>CUENTA</div>
                <select value={movForm.cuenta} onChange={e => setMovForm({...movForm, cuenta: e.target.value})}>
                  <option value="">Seleccioná...</option>
                  {cuentas.map(c => <option key={c.id}>{c.nombre}</option>)}
                </select>
              </div>

              {(movForm.tipo === "deposito_ars" || movForm.tipo === "retiro_ars") && (
                <div>
                  <div style={lbl}>MONTO ARS</div>
                  <input type="number" placeholder="0" value={movForm.ars} onChange={e => setMovForm({...movForm, ars: e.target.value})} />
                </div>
              )}

              {(movForm.tipo === "retiro_cajero" || movForm.tipo === "deposito_usd") && (
                <div>
                  <div style={lbl}>MONTO USD</div>
                  <input type="number" placeholder="0" value={movForm.usd} onChange={e => setMovForm({...movForm, usd: e.target.value})} />
                  {movForm.tipo === "retiro_cajero" && (
                    <div style={{ fontSize: 10, color: "#e07a3a", marginTop: 4 }}>Los USD se descontarán de la cuenta y se sumarán a tu efectivo físico</div>
                  )}
                </div>
              )}

              <div>
                <div style={lbl}>FECHA</div>
                <input type="date" value={movForm.fecha} onChange={e => setMovForm({...movForm, fecha: e.target.value})} />
              </div>
              <div>
                <div style={lbl}>NOTA (opcional)</div>
                <input placeholder="Descripción..." value={movForm.nota} onChange={e => setMovForm({...movForm, nota: e.target.value})} />
              </div>

              <button onClick={submitMovimiento} disabled={syncing} style={{ background: "#2a9d5c", color: "#fff", border: "none", borderRadius: 8, padding: "14px", fontSize: 13, fontWeight: 700, opacity: syncing ? 0.6 : 1 }}>
                {syncing ? "GUARDANDO..." : "REGISTRAR MOVIMIENTO"}
              </button>
            </div>

            {/* Historial movimientos */}
            {movimientos.length > 0 && (
              <div style={{ marginTop: 24 }}>
                <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 12, letterSpacing: 2 }}>ÚLTIMOS MOVIMIENTOS</div>
                {movimientos.slice(0, 10).map(m => (
                  <div key={m.id} style={{ ...card, padding: "10px 14px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12 }}>
                      <span style={{ color: "#888" }}>{m.tipo.replace("_", " ").toUpperCase()} · {m.cuenta}</span>
                      <span style={{ color: "#555", fontSize: 10 }}>{m.fecha}</span>
                    </div>
                    <div style={{ display: "flex", gap: 12, marginTop: 4, fontSize: 12 }}>
                      {m.ars && <span style={{ color: m.ars > 0 ? "#2a9d5c" : "#e05a5a" }}>{m.ars > 0 ? "+" : ""}{fmtARS(m.ars)}</span>}
                      {m.usd && <span style={{ color: m.usd > 0 ? "#6a8fff" : "#e07a3a" }}>{m.usd > 0 ? "+" : ""}{fmtUSD(m.usd)}</span>}
                    </div>
                    {m.nota && <div style={{ fontSize: 10, color: "#444", marginTop: 2 }}>📝 {m.nota}</div>}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* NUEVA COMPRA */}
        {view === "nueva" && (
          <div>
            <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 16, letterSpacing: 2 }}>REGISTRAR COMPRA</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div>
                <div style={lbl}>PERSONA</div>
                <select value={form.persona} onChange={e => setForm({...form, persona: e.target.value})}>
                  <option value="">Seleccioná...</option>
                  {personas.map(p => <option key={p}>{p}</option>)}
                </select>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <div>
                  <div style={lbl}>USD COMPRADO</div>
                  <input type="number" placeholder="0" value={form.usd} onChange={e => setForm({...form, usd: e.target.value})} />
                </div>
                <div>
                  <div style={lbl}>TC COMPRA</div>
                  <input type="number" placeholder="1385" value={form.tc} onChange={e => setForm({...form, tc: e.target.value})} />
                </div>
              </div>
              {form.usd && form.tc && (
                <div style={{ background: "#0d2e1a", border: "1px solid #1a4d2e", borderRadius: 6, padding: "10px 14px", fontSize: 12, color: "#2a9d5c" }}>
                  ARS a debitar: <strong>{fmtARS(parseFloat(form.usd||0) * parseFloat(form.tc||0))}</strong>
                  {form.banco && cuentas.find(c => c.nombre === form.banco) && (
                    <span style={{ marginLeft: 12, color: "#555" }}>
                      → Saldo restante: {fmtARS(cuentas.find(c => c.nombre === form.banco).saldo_ars - parseFloat(form.usd||0) * parseFloat(form.tc||0))}
                    </span>
                  )}
                </div>
              )}
              <div>
                <div style={lbl}>CUENTA / BANCO</div>
                <select value={form.banco} onChange={e => setForm({...form, banco: e.target.value})}>
                  <option value="">Seleccioná...</option>
                  {cuentas.map(c => <option key={c.id}>{c.nombre} — ARS: {fmtARS(c.saldo_ars)}</option>)}
                </select>
              </div>
              <div>
                <div style={lbl}>FECHA</div>
                <input type="date" value={form.fecha} onChange={e => setForm({...form, fecha: e.target.value})} />
              </div>
              <div>
                <div style={lbl}>NOTA (opcional)</div>
                <input placeholder="Ej: MEP, blue..." value={form.nota} onChange={e => setForm({...form, nota: e.target.value})} />
              </div>
              <button onClick={submitOp} disabled={syncing} style={{ background: "#2a9d5c", color: "#fff", border: "none", borderRadius: 8, padding: "14px", fontSize: 13, fontWeight: 700, opacity: syncing ? 0.6 : 1 }}>
                {syncing ? "GUARDANDO..." : "REGISTRAR COMPRA"}
              </button>
            </div>
          </div>
        )}

        {/* VENTA */}
        {view === "venta" && (
          <div>
            <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 4, letterSpacing: 2 }}>VENTA GRUPAL · PARCIAL</div>
            <div style={{ fontSize: 11, color: "#555", marginBottom: 14 }}>Tildá las compras → ajustá USD si es parcial → confirmá</div>
            {availableForSale.length > 0 && (
              <button onClick={selectAll} style={{ background: "transparent", border: "1px solid #2a2a3a", color: "#777", borderRadius: 6, padding: "7px 14px", fontSize: 11, marginBottom: 12 }}>
                {ventaSelected.length === availableForSale.length ? "Deseleccionar todo" : "Seleccionar todo"}
              </button>
            )}
            {availableForSale.length === 0 && <div style={{ color: "#444", fontSize: 12 }}>No hay compras en stock.</div>}
            {availableForSale.map(op => {
              const sel = isSelected(op.id);
              const item = ventaSelected.find(x => x.id === op.id);
              const usdParcial = item ? item.usdParcial : op.usd;
              const esParcial = sel && parseFloat(usdParcial) < op.usd - 0.001;
              return (
                <div key={op.id} style={{ ...card, border: sel ? "1px solid #2a9d5c" : "1px solid #1e1e2e", background: sel ? "#0a1a10" : "#12121a" }}>
                  <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
                    <input type="checkbox" className="chk" style={{ marginTop: 2 }} checked={sel} onChange={() => toggleSelect(op)} />
                    <div style={{ flex: 1 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                        <span style={{ fontWeight: 600, fontSize: 13 }}>{op.persona}</span>
                        <span style={{ fontSize: 10, color: "#555" }}>{op.banco} · {op.fecha}</span>
                      </div>
                      <div style={{ display: "flex", gap: 14, fontSize: 12, marginBottom: sel ? 10 : 0 }}>
                        <span>{fmtUSD(op.usd)}</span>
                        <span style={{ color: "#6a8fff" }}>TC: $ {fmt(op.tc)}</span>
                      </div>
                      {sel && (
                        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 4 }}>
                          <div style={{ fontSize: 10, color: "#555", whiteSpace: "nowrap" }}>USD A VENDER:</div>
                          <input type="number" className="parcial-input" value={usdParcial} max={op.usd} min={0.01} step={0.01}
                            onClick={e => e.stopPropagation()} onChange={e => setParcial(op.id, e.target.value)} />
                          {esParcial
                            ? <div style={{ fontSize: 10, color: "#e07a3a", whiteSpace: "nowrap" }}>resto: {fmtUSD(op.usd - parseFloat(usdParcial||0))}</div>
                            : <div style={{ fontSize: 10, color: "#2a9d5c", whiteSpace: "nowrap" }}>total</div>
                          }
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}

            {selectedItems.length > 0 && (
              <div style={{ background: "#0a1220", border: "1px solid #1a2540", borderRadius: 12, padding: "16px", marginTop: 6 }}>
                <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#6a8fff", marginBottom: 12, letterSpacing: 2 }}>RESUMEN DEL LOTE</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 14 }}>
                  <div><div style={{ fontSize: 10, color: "#555", marginBottom: 3 }}>USD A VENDER</div><div style={{ fontSize: 16, fontWeight: 700 }}>{fmtUSD(totalUSDLote)}</div></div>
                  <div><div style={{ fontSize: 10, color: "#555", marginBottom: 3 }}>TC PROM. COMPRA</div><div style={{ fontSize: 16, fontWeight: 700, color: "#6a8fff" }}>$ {fmt(tcPromLote)}</div></div>
                  <div><div style={{ fontSize: 10, color: "#555", marginBottom: 3 }}>ARS A RECIBIR</div><div style={{ fontSize: 16, fontWeight: 700, color: "#2a9d5c" }}>{tcVenta ? fmtARS(totalUSDLote * parseFloat(tcVenta)) : "-"}</div></div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
                  <div>
                    <div style={{ ...lbl, marginBottom: 4 }}>TC VENTA</div>
                    <input type="number" placeholder="1420" value={tcVenta} onChange={e => setTcVenta(e.target.value)} />
                  </div>
                  <div>
                    <div style={{ ...lbl, marginBottom: 4 }}>FECHA VENTA</div>
                    <input type="date" value={fechaVenta} onChange={e => setFechaVenta(e.target.value)} />
                  </div>
                </div>
                <div style={{ marginBottom: 12 }}>
                  <div style={{ ...lbl, marginBottom: 4 }}>ARS INGRESA A CUENTA</div>
                  <select value={cuentaVenta} onChange={e => setCuentaVenta(e.target.value)}>
                    <option value="">Seleccioná cuenta...</option>
                    {cuentas.map(c => <option key={c.id}>{c.nombre}</option>)}
                  </select>
                </div>
                {tcVenta && gananciaLote !== null && (
                  <div style={{ marginBottom: 12, padding: "12px", borderRadius: 8, background: gananciaLote >= 0 ? "#0d2e1a" : "#2e0d0d", border: `1px solid ${gananciaLote >= 0 ? "#1a4d2e" : "#4d1a1a"}` }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ fontSize: 12, color: "#888" }}>$ {fmt(tcPromLote)} → $ {fmt(parseFloat(tcVenta))}</div>
                      <div style={{ fontSize: 18, fontWeight: 700, color: gananciaLote >= 0 ? "#2a9d5c" : "#e05a5a" }}>
                        {gananciaLote >= 0 ? "+" : ""}{fmtARS(gananciaLote)}
                      </div>
                    </div>
                  </div>
                )}
                <button onClick={submitVenta} disabled={syncing} style={{ width: "100%", background: "#e07a3a", color: "#fff", border: "none", borderRadius: 8, padding: "13px", fontSize: 13, fontWeight: 700, opacity: syncing ? 0.6 : 1 }}>
                  {syncing ? "GUARDANDO..." : `CONFIRMAR VENTA · ${fmtUSD(totalUSDLote)}`}
                </button>
              </div>
            )}
          </div>
        )}

        {/* HISTORIAL */}
        {view === "historial" && (
          <div>
            <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 12, letterSpacing: 2 }}>HISTORIAL</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 14 }}>
              <select value={filterPersona} onChange={e => setFilterPersona(e.target.value)}>
                <option value="todas">Todas las personas</option>
                {personas.map(p => <option key={p}>{p}</option>)}
              </select>
              <select value={filterMes} onChange={e => setFilterMes(e.target.value)}>
                <option value="todos">Todos los meses</option>
                {meses.map(m => <option key={m}>{m}</option>)}
              </select>
            </div>
            {filteredOps.length === 0 && <div style={{ color: "#444", fontSize: 12 }}>Sin operaciones.</div>}
            {filteredOps.map(op => {
              const usdV = op.usd_vendido || op.usd;
              const gan = op.tc_venta ? usdV * (op.tc_venta - op.tc) : null;
              const esParcial = op.tc_venta && op.usd_vendido && Math.abs(op.usd_vendido - op.usd) > 0.001;
              return (
                <div key={op.id} style={{ ...card, borderColor: op.tc_venta ? "#1a2e1a" : "#1e1e2e" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                    <div style={{ fontWeight: 600, fontSize: 13 }}>{op.persona}</div>
                    <div style={{ display: "flex", gap: 5, alignItems: "center" }}>
                      {esParcial && <span style={{ fontSize: 10, padding: "2px 6px", borderRadius: 4, background: "#2e1a0a", color: "#e07a3a", border: "1px solid #4d2e0a" }}>PARCIAL</span>}
                      <span style={{ fontSize: 10, padding: "2px 7px", borderRadius: 4, background: op.tc_venta ? "#0d2e1a" : "#0d0d2e", color: op.tc_venta ? "#2a9d5c" : "#6a8fff", border: `1px solid ${op.tc_venta ? "#1a4d2e" : "#1a1a4d"}` }}>
                        {op.tc_venta ? "VENDIDO" : "EN STOCK"}
                      </span>
                      <span style={{ fontSize: 10, padding: "2px 6px", borderRadius: 4, background: "#0d0d14", color: "#555", border: "1px solid #1a1a2a" }}>{op.banco}</span>
                      <button onClick={() => setConfirmDel(op.id)} style={{ background: "none", border: "none", color: "#2a2a3a", fontSize: 14 }}>✕</button>
                    </div>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6, fontSize: 11 }}>
                    <div style={{ color: "#555" }}>USD<br/><span style={{ color: "#e8e0d0", fontSize: 13 }}>{fmtUSD(op.usd)}</span></div>
                    <div style={{ color: "#555" }}>TC COMPRA<br/><span style={{ color: "#6a8fff", fontSize: 13 }}>$ {fmt(op.tc)}</span></div>
                    <div style={{ color: "#555" }}>ARS<br/><span style={{ color: "#aaa", fontSize: 13 }}>{fmtARS(op.ars)}</span></div>
                  </div>
                  {op.tc_venta && (
                    <div style={{ marginTop: 8, paddingTop: 8, borderTop: "1px solid #1a2e1a", display: "flex", justifyContent: "space-between", fontSize: 11 }}>
                      <span style={{ color: "#555" }}>{esParcial ? `${fmtUSD(usdV)} @` : "Vendido @"} <span style={{ color: "#e07a3a" }}>$ {fmt(op.tc_venta)}</span> · {op.fecha_venta}</span>
                      <span style={{ color: gan >= 0 ? "#2a9d5c" : "#e05a5a", fontWeight: 700 }}>{gan >= 0 ? "+" : ""}{fmtARS(gan)}</span>
                    </div>
                  )}
                  {op.nota && <div style={{ marginTop: 6, fontSize: 10, color: "#444" }}>📝 {op.nota}</div>}
                  <div style={{ marginTop: 4, fontSize: 10, color: "#222" }}>{op.fecha}</div>
                </div>
              );
            })}
          </div>
        )}

        {/* PERSONAS */}
        {view === "personas" && (
          <div>
            <div style={{ fontFamily: "'Unbounded'", fontSize: 10, color: "#444", marginBottom: 16, letterSpacing: 2 }}>PERSONAS</div>
            <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
              <input placeholder="Nombre..." value={newPersona} onChange={e => setNewPersona(e.target.value)} onKeyDown={e => e.key === "Enter" && addPersona()} />
              <button onClick={addPersona} style={{ background: "#2a9d5c", color: "#fff", border: "none", borderRadius: 6, padding: "0 16px", fontSize: 12, fontWeight: 700, whiteSpace: "nowrap" }}>+ Agregar</button>
            </div>
            {personas.length === 0 && <div style={{ color: "#444", fontSize: 12 }}>Sin personas.</div>}
            {personas.map(p => {
              const s = personaSummary(p);
              return (
                <div key={p} style={{ ...card, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 3 }}>{p}</div>
                    <div style={{ fontSize: 11, color: "#555" }}>{fmtUSD(s.totalU)} comprado · stock <span style={{ color: "#2a9d5c" }}>{fmtUSD(s.stock)}</span></div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {confirmDel && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.8)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}>
          <div style={{ background: "#12121a", border: "1px solid #2a2a3a", borderRadius: 12, padding: 24, maxWidth: 300, width: "90%" }}>
            <div style={{ marginBottom: 16, fontSize: 14 }}>¿Eliminar esta operación?</div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => deleteOp(confirmDel)} style={{ flex: 1, background: "#e05a5a", color: "#fff", border: "none", borderRadius: 6, padding: "10px", fontSize: 12, fontWeight: 700 }}>Eliminar</button>
              <button onClick={() => setConfirmDel(null)} style={{ flex: 1, background: "transparent", color: "#aaa", border: "1px solid #2a2a3a", borderRadius: 6, padding: "10px", fontSize: 12 }}>Cancelar</button>
            </div>
          </div>
        </div>
      )}

      {toast && (
        <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: toast.type === "err" ? "#2e0d0d" : "#0d2e1a", border: `1px solid ${toast.type === "err" ? "#4d1a1a" : "#1a4d2e"}`, color: toast.type === "err" ? "#e05a5a" : "#2a9d5c", padding: "10px 20px", borderRadius: 8, fontSize: 12, fontWeight: 600, zIndex: 100, maxWidth: "92vw", textAlign: "center" }}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}
