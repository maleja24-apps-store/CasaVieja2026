import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const MEDIOS_PAGO = ["Efectivo", "Tarjeta", "Nequi", "Daviplata", "Bancolombia"];
const CATEGORIAS = ["Juguetería", "Didacticos", "Botilo-Termo", "Mug", "Hogar", "Miscelánea", "Escolar", "Ropa Hogar"];
const STOCK_MINIMO = 3;

const exportToExcel = async (ventas, gastos, inventario, onDone, onError) => {
  const wb = XLSX.utils.book_new();

  const ventasData = ventas.map((v) => ({
    Fecha: v.fecha, Referencia: v.ref, Descripción: v.desc, Categoría: v.cat,
    Cantidad: v.cantidad, "Precio Venta": v.precio, Total: v.total, "Medio de Pago": v.medio,
  }));
  const wsVentas = XLSX.utils.json_to_sheet(ventasData.length ? ventasData : [{ Fecha: "", Referencia: "", Descripción: "", Categoría: "", Cantidad: "", "Precio Venta": "", Total: "", "Medio de Pago": "" }]);
  wsVentas["!cols"] = [10, 14, 28, 14, 10, 14, 14, 16].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsVentas, "Ventas");

  const gastosData = gastos.map((g) => ({
    Fecha: g.fecha, Concepto: g.concepto, Valor: g.valor, "Medio de Pago": g.medio,
  }));
  const wsGastos = XLSX.utils.json_to_sheet(gastosData.length ? gastosData : [{ Fecha: "", Concepto: "", Valor: "", "Medio de Pago": "" }]);
  wsGastos["!cols"] = [10, 30, 14, 16].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsGastos, "Gastos");

  const invData = inventario.map((p) => ({
    Referencia: p.ref, Descripción: p.desc, Categoría: p.cat, Proveedor: p.proveedor || "",
    "Stock Actual": p.stock ?? p.cantidad, "Stock Mínimo": p.stockMin,
    "Precio Costo": p.costo || 0, "Precio Venta": p.precioVenta || 0, "Fecha Ingreso": p.fecha,
  }));
  const wsInv = XLSX.utils.json_to_sheet(invData.length ? invData : [{ Referencia: "", Descripción: "", Categoría: "", Proveedor: "", "Stock Actual": "", "Stock Mínimo": "", "Precio Costo": "", "Precio Venta": "", "Fecha Ingreso": "" }]);
  wsInv["!cols"] = [14, 28, 14, 18, 12, 12, 14, 14, 14].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsInv, "Inventario");

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const fecha = new Date().toISOString().split("T")[0];
  const fileName = `Cositas_pa_Sumerce_${fecha}.xlsx`;
  const file = new File([blob], fileName, { type: blob.type });

  // Intentar compartir (funciona en Android e iOS)
  if (navigator.share && navigator.canShare && navigator.canShare({ files: [file] })) {
    try {
      await navigator.share({ files: [file], title: fileName });
      if (onDone) onDone();
    } catch (err) {
      if (err.name !== "AbortError" && onError) onError();
    }
  } else {
    // Fallback: descarga directa
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); document.body.removeChild(a); }, 1000);
    if (onDone) onDone();
  }
};

const formatCOP = (n) =>
  new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", minimumFractionDigits: 0 }).format(n || 0);

const today = () => new Date().toISOString().split("T")[0];

const STORAGE_KEYS = {
  ventas: "cositas_ventas",
  gastos: "cositas_gastos",
  inventario: "cositas_inventario",
};

function useStorage(key, initial) {
  const [data, setData] = useState(() => {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : initial;
    } catch {
      return initial;
    }
  });
  const set = useCallback((val) => {
    setData((prev) => {
      const next = typeof val === "function" ? val(prev) : val;
      localStorage.setItem(key, JSON.stringify(next));
      return next;
    });
  }, [key]);
  return [data, set];
}

// ── ICONS ──────────────────────────────────────────────────────────────────
const Icon = ({ name }) => {
  const icons = {
    venta: "🛍️", gasto: "💸", inventario: "📦", resumen: "📊",
    plus: "＋", check: "✓", alert: "⚠️", trash: "🗑️", close: "✕",
    up: "↑", down: "↓",
  };
  return <span>{icons[name] || "●"}</span>;
};

// ── TOAST ───────────────────────────────────────────────────────────────────
function Toast({ msg, onClose }) {
  useEffect(() => { const t = setTimeout(onClose, 2800); return () => clearTimeout(t); }, [onClose]);
  return (
    <div style={{
      position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)",
      background: "#1a1a2e", color: "#f0e6d3", padding: "12px 24px",
      borderRadius: 32, fontFamily: "inherit", fontSize: 14, zIndex: 999,
      boxShadow: "0 8px 32px rgba(0,0,0,0.4)", border: "1px solid #e8c547",
      display: "flex", alignItems: "center", gap: 8, whiteSpace: "nowrap",
    }}>
      <Icon name="check" /> {msg}
      <button onClick={onClose} style={{ background: "none", border: "none", color: "#f0e6d3", cursor: "pointer", marginLeft: 8 }}>✕</button>
    </div>
  );
}

// ── VENTAS ──────────────────────────────────────────────────────────────────
function Ventas({ ventas, setVentas, inventario, setInventario, toast }) {
  const [form, setForm] = useState({ fecha: today(), ref: "", desc: "", cat: CATEGORIAS[0], cantidad: 1, precio: "", medio: MEDIOS_PAGO[0] });
  const [showForm, setShowForm] = useState(false);
  const [suggs, setSuggs] = useState([]);
  const [showSuggs, setShowSuggs] = useState(false);

  const handleChange = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  const handleRefInput = (v) => {
    handleChange("ref", v);
    if (v.length >= 1) {
      const matches = inventario.filter((p) => p.ref.toLowerCase().includes(v.toLowerCase()) || p.desc.toLowerCase().includes(v.toLowerCase()));
      setSuggs(matches.slice(0, 6));
      setShowSuggs(true);
    } else {
      setSuggs([]);
      setShowSuggs(false);
    }
  };

  const selectProduct = (p) => {
    setForm((f) => ({ ...f, ref: p.ref, desc: p.desc, cat: p.cat || f.cat, precio: p.precioVenta || "" }));
    setSuggs([]);
    setShowSuggs(false);
  };

  const handleSubmit = () => {
    if (!form.ref || !form.desc || !form.precio) return toast("Completa todos los campos obligatorios");
    const total = Number(form.cantidad) * Number(form.precio);
    const venta = { ...form, cantidad: Number(form.cantidad), precio: Number(form.precio), total, id: Date.now() };
    setVentas((v) => [venta, ...v]);
    setInventario((inv) => inv.map((p) =>
      p.ref === form.ref ? { ...p, stock: Math.max(0, (p.stock || 0) - Number(form.cantidad)) } : p
    ));
    setForm({ fecha: today(), ref: "", desc: "", cat: CATEGORIAS[0], cantidad: 1, precio: "", medio: MEDIOS_PAGO[0] });
    setShowForm(false);
    toast("Venta registrada ✓");
  };

  const del = (id) => setVentas((v) => v.filter((x) => x.id !== id));

  return (
    <Section title="Ventas" icon="venta" accent="#e8c547">
      <button className="fab" onClick={() => setShowForm(true)}>＋ Nueva venta</button>

      {showForm && (
        <Modal title="Registrar Venta" onClose={() => { setShowForm(false); setShowSuggs(false); }} onSave={handleSubmit} accent="#e8c547">
          <Field label="Fecha"><input type="date" value={form.fecha} onChange={(e) => handleChange("fecha", e.target.value)} /></Field>
          <Field label="Referencia o nombre del producto *">
            <div style={{ position: "relative" }}>
              <input
                placeholder="Escribe la referencia o nombre..."
                value={form.ref}
                onChange={(e) => handleRefInput(e.target.value)}
                onBlur={() => setTimeout(() => setShowSuggs(false), 150)}
                autoComplete="off"
              />
              {showSuggs && suggs.length > 0 && (
                <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: "#1e1e2e", border: "1px solid #e8c547", borderRadius: 10, zIndex: 200, maxHeight: 200, overflowY: "auto", marginTop: 4 }}>
                  {suggs.map((p) => (
                    <div key={p.id} onMouseDown={() => selectProduct(p)} style={{ padding: "10px 14px", cursor: "pointer", borderBottom: "1px solid rgba(255,255,255,0.06)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div>
                        <span style={{ color: "#e8c547", fontWeight: 700, fontSize: 12 }}>{p.ref}</span>
                        <span style={{ marginLeft: 8, fontSize: 13 }}>{p.desc}</span>
                      </div>
                      <span style={{ fontSize: 12, color: "#56cfe1", fontWeight: 600 }}>{p.precioVenta ? formatCOP(p.precioVenta) : ""}</span>
                    </div>
                  ))}
                </div>
              )}
              {showSuggs && suggs.length === 0 && form.ref.length > 0 && (
                <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: "#1e1e2e", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 10, zIndex: 200, padding: "10px 14px", fontSize: 13, opacity: 0.5, marginTop: 4 }}>
                  Sin coincidencias en inventario
                </div>
              )}
            </div>
          </Field>
          <Field label="Descripción *">
            <input placeholder="Se autocompleta al seleccionar referencia" value={form.desc} onChange={(e) => handleChange("desc", e.target.value)} />
          </Field>
          <Field label="Categoría">
            <select value={form.cat} onChange={(e) => handleChange("cat", e.target.value)}>
              {CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Row>
            <Field label="Cantidad"><input type="number" min="1" value={form.cantidad} onChange={(e) => handleChange("cantidad", e.target.value)} /></Field>
            <Field label="Precio ($) *"><input type="number" placeholder="0" value={form.precio} onChange={(e) => handleChange("precio", e.target.value)} /></Field>
          </Row>
          <Field label="Medio de pago">
            <select value={form.medio} onChange={(e) => handleChange("medio", e.target.value)}>
              {MEDIOS_PAGO.map((m) => <option key={m}>{m}</option>)}
            </select>
          </Field>
          <Total label="Total" value={formatCOP(Number(form.cantidad) * Number(form.precio))} />
        </Modal>
      )}

      <TableWrap>
        <table>
          <thead><tr><th>Fecha</th><th>Ref</th><th>Producto</th><th>Cat</th><th>Cant</th><th>Precio</th><th>Total</th><th>Pago</th><th></th></tr></thead>
          <tbody>
            {ventas.length === 0 && <tr><td colSpan={9} style={{ textAlign: "center", opacity: 0.5, padding: 24 }}>Sin ventas aún</td></tr>}
            {ventas.map((v) => (
              <tr key={v.id}>
                <td>{v.fecha}</td><td><Tag>{v.ref}</Tag></td><td>{v.desc}</td><td><Tag soft>{v.cat}</Tag></td>
                <td style={{ textAlign: "center" }}>{v.cantidad}</td>
                <td>{formatCOP(v.precio)}</td><td style={{ fontWeight: 700, color: "#e8c547" }}>{formatCOP(v.total)}</td>
                <td><Tag>{v.medio}</Tag></td>
                <td><button className="del-btn" onClick={() => del(v.id)}>✕</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </TableWrap>
    </Section>
  );
}

// ── GASTOS ──────────────────────────────────────────────────────────────────
function Gastos({ gastos, setGastos, toast }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ fecha: today(), concepto: "", valor: "", medio: MEDIOS_PAGO[0] });
  const h = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  const handleSubmit = () => {
    if (!form.concepto || !form.valor) return toast("Completa concepto y valor");
    setGastos((g) => [{ ...form, valor: Number(form.valor), id: Date.now() }, ...g]);
    setForm({ fecha: today(), concepto: "", valor: "", medio: MEDIOS_PAGO[0] });
    setShowForm(false);
    toast("Gasto registrado ✓");
  };

  const del = (id) => setGastos((g) => g.filter((x) => x.id !== id));

  return (
    <Section title="Gastos" icon="gasto" accent="#f07167">
      <button className="fab" style={{ background: "#f07167" }} onClick={() => setShowForm(true)}>＋ Nuevo gasto</button>
      {showForm && (
        <Modal title="Registrar Gasto" onClose={() => setShowForm(false)} onSave={handleSubmit} accent="#f07167">
          <Field label="Fecha"><input type="date" value={form.fecha} onChange={(e) => h("fecha", e.target.value)} /></Field>
          <Field label="Concepto *"><input placeholder="Ej: Domicilio, proveedor, servicios..." value={form.concepto} onChange={(e) => h("concepto", e.target.value)} /></Field>
          <Field label="Valor ($) *"><input type="number" placeholder="0" value={form.valor} onChange={(e) => h("valor", e.target.value)} /></Field>
          <Field label="Medio de pago">
            <select value={form.medio} onChange={(e) => h("medio", e.target.value)}>
              {MEDIOS_PAGO.map((m) => <option key={m}>{m}</option>)}
            </select>
          </Field>
          <Total label="Valor" value={formatCOP(Number(form.valor))} color="#f07167" />
        </Modal>
      )}
      <TableWrap>
        <table>
          <thead><tr><th>Fecha</th><th>Concepto</th><th>Valor</th><th>Pago</th><th></th></tr></thead>
          <tbody>
            {gastos.length === 0 && <tr><td colSpan={5} style={{ textAlign: "center", opacity: 0.5, padding: 24 }}>Sin gastos registrados</td></tr>}
            {gastos.map((g) => (
              <tr key={g.id}>
                <td>{g.fecha}</td><td>{g.concepto}</td>
                <td style={{ fontWeight: 700, color: "#f07167" }}>{formatCOP(g.valor)}</td>
                <td><Tag>{g.medio}</Tag></td>
                <td><button className="del-btn" onClick={() => del(g.id)}>✕</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </TableWrap>
    </Section>
  );
}

// ── INVENTARIO ───────────────────────────────────────────────────────────────
function Inventario({ inventario, setInventario, toast }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ fecha: today(), ref: "", desc: "", cat: CATEGORIAS[0], cantidad: 1, costo: "", precioVenta: "", stockMin: STOCK_MINIMO, proveedor: "" });
  const h = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  const handleSubmit = () => {
    if (!form.ref || !form.desc) return toast("Completa referencia y descripción");
    const existing = inventario.findIndex((p) => p.ref === form.ref);
    if (existing >= 0) {
      setInventario((inv) => inv.map((p, i) => i === existing ? { ...p, stock: (p.stock || 0) + Number(form.cantidad), costo: Number(form.costo) || p.costo, precioVenta: Number(form.precioVenta) || p.precioVenta } : p));
      toast("Stock actualizado ✓");
    } else {
      setInventario((inv) => [{ ...form, cantidad: Number(form.cantidad), costo: Number(form.costo), precioVenta: Number(form.precioVenta), stock: Number(form.cantidad), stockMin: Number(form.stockMin), id: Date.now() }, ...inv]);
      toast("Producto ingresado ✓");
    }
    setForm({ fecha: today(), ref: "", desc: "", cat: CATEGORIAS[0], cantidad: 1, costo: "", precioVenta: "", stockMin: STOCK_MINIMO, proveedor: "" });
    setShowForm(false);
  };

  const importFromExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (rows.length === 0) return toast("El archivo está vacío");

        // Mapeo flexible de columnas
        const map = (row, keys) => {
          for (const k of keys) {
            const found = Object.keys(row).find((rk) => rk.toLowerCase().replace(/\s/g, "").includes(k));
            if (found && row[found] !== "") return String(row[found]).trim();
          }
          return "";
        };

        let importados = 0;
        let actualizados = 0;
        const nuevos = [];

        rows.forEach((row, i) => {
          const ref = map(row, ["referencia", "ref", "codigo", "código"]);
          const desc = map(row, ["descripcion", "descripción", "nombre", "producto"]);
          if (!ref && !desc) return;

          const producto = {
            id: Date.now() + i,
            fecha: today(),
            ref: ref || `IMPORT-${i + 1}`,
            desc: desc || ref,
            cat: map(row, ["categoria", "categoría", "category"]) || CATEGORIAS[0],
            proveedor: map(row, ["proveedor", "supplier"]),
            cantidad: Number(map(row, ["cantidad", "qty", "stock"])) || 0,
            stock: Number(map(row, ["stockactual", "stock", "cantidad", "qty"])) || 0,
            stockMin: Number(map(row, ["stockmin", "minimo", "mínimo"])) || STOCK_MINIMO,
            costo: Number(map(row, ["costo", "cost", "preciocosto"])) || 0,
            precioVenta: Number(map(row, ["precioventa", "precio", "price", "pvp"])) || 0,
          };

          const existingIdx = inventario.findIndex((p) => p.ref === producto.ref);
          if (existingIdx >= 0) {
            actualizados++;
          } else {
            nuevos.push(producto);
            importados++;
          }
        });

        setInventario((inv) => [...nuevos, ...inv]);
        toast(`✓ ${importados} productos importados${actualizados > 0 ? `, ${actualizados} ya existían` : ""}`);
      } catch {
        toast("Error leyendo el archivo. Verifica que sea .xlsx");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const del = (id) => setInventario((inv) => inv.filter((x) => x.id !== id));
  const alertas = inventario.filter((p) => (p.stock || 0) <= (p.stockMin || STOCK_MINIMO));

  return (
    <Section title="Inventario" icon="inventario" accent="#56cfe1">
      {alertas.length > 0 && (
        <div style={{ background: "#2d1b00", border: "1px solid #e8c547", borderRadius: 12, padding: "12px 16px", marginBottom: 16, display: "flex", alignItems: "flex-start", gap: 10 }}>
          <span style={{ fontSize: 18 }}>⚠️</span>
          <div>
            <div style={{ fontWeight: 700, color: "#e8c547", marginBottom: 4 }}>Productos por agotarse</div>
            {alertas.map((p) => (
              <div key={p.id} style={{ fontSize: 13, color: "#f0e6d3", opacity: 0.9 }}>
                {p.desc} <span style={{ color: "#f07167", fontWeight: 700 }}>({p.stock} unid.)</span>
              </div>
            ))}
          </div>
        </div>
      )}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 4 }}>
        <button className="fab" style={{ background: "#56cfe1", color: "#0d1117", marginBottom: 0 }} onClick={() => setShowForm(true)}>＋ Ingresar producto</button>
        <label style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "10px 20px", background: "rgba(86,207,225,0.1)", border: "1px solid #56cfe1", color: "#56cfe1", borderRadius: 30, fontWeight: 700, fontFamily: "inherit", fontSize: 14, cursor: "pointer", marginBottom: 16 }}>
          📥 Importar Excel
          <input type="file" accept=".xlsx,.xls" onChange={importFromExcel} style={{ display: "none" }} />
        </label>
      </div>
      {showForm && (
        <Modal title="Ingreso de Producto" onClose={() => setShowForm(false)} onSave={handleSubmit} accent="#56cfe1">
          <Field label="Fecha"><input type="date" value={form.fecha} onChange={(e) => h("fecha", e.target.value)} /></Field>
          <Field label="Referencia *"><input placeholder="REF-001" value={form.ref} onChange={(e) => h("ref", e.target.value)} /></Field>
          <Field label="Descripción *"><input placeholder="Nombre del producto" value={form.desc} onChange={(e) => h("desc", e.target.value)} /></Field>
          <Field label="Categoría">
            <select value={form.cat} onChange={(e) => h("cat", e.target.value)}>
              {CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Row>
            <Field label="Cantidad"><input type="number" min="1" value={form.cantidad} onChange={(e) => h("cantidad", e.target.value)} /></Field>
            <Field label="Stock mín. alerta"><input type="number" min="0" value={form.stockMin} onChange={(e) => h("stockMin", e.target.value)} /></Field>
          </Row>
          <Field label="Proveedor"><input placeholder="Ej: Panamericana, distribuidor..." value={form.proveedor} onChange={(e) => h("proveedor", e.target.value)} /></Field>
          <Row>
            <Field label="Costo ($)"><input type="number" placeholder="0" value={form.costo} onChange={(e) => h("costo", e.target.value)} /></Field>
            <Field label="Precio venta ($)"><input type="number" placeholder="0" value={form.precioVenta} onChange={(e) => h("precioVenta", e.target.value)} /></Field>
          </Row>
        </Modal>
      )}
      <TableWrap>
        <table>
          <thead><tr><th>Ref</th><th>Producto</th><th>Cat</th><th>Proveedor</th><th>Stock</th><th>Costo</th><th>P. Venta</th><th></th></tr></thead>
          <tbody>
            {inventario.length === 0 && <tr><td colSpan={8} style={{ textAlign: "center", opacity: 0.5, padding: 24 }}>Sin productos registrados</td></tr>}
            {inventario.map((p) => {
              const low = (p.stock || 0) <= (p.stockMin || STOCK_MINIMO);
              return (
                <tr key={p.id} style={low ? { background: "rgba(248,113,70,0.08)" } : {}}>
                  <td><Tag>{p.ref}</Tag></td><td>{p.desc}</td><td><Tag soft>{p.cat}</Tag></td>
                  <td style={{ opacity: p.proveedor ? 1 : 0.35 }}>{p.proveedor || "—"}</td>
                  <td style={{ textAlign: "center" }}>
                    <span style={{ fontWeight: 700, color: low ? "#f07167" : "#56cfe1" }}>{p.stock ?? p.cantidad}</span>
                    {low && <span style={{ marginLeft: 4 }}>⚠️</span>}
                  </td>
                  <td>{p.costo ? formatCOP(p.costo) : "—"}</td>
                  <td>{p.precioVenta ? formatCOP(p.precioVenta) : "—"}</td>
                  <td><button className="del-btn" onClick={() => del(p.id)}>✕</button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </TableWrap>
    </Section>
  );
}

// ── RESUMEN ──────────────────────────────────────────────────────────────────
function Resumen({ ventas, gastos }) {
  const [fecha, setFecha] = useState(today());
  const ventasDia = ventas.filter((v) => v.fecha === fecha);
  const gastosDia = gastos.filter((g) => g.fecha === fecha);
  const totalVentas = ventasDia.reduce((s, v) => s + v.total, 0);
  const totalGastos = gastosDia.reduce((s, g) => s + g.valor, 0);
  const utilidad = totalVentas - totalGastos;

  const porMedio = MEDIOS_PAGO.map((m) => ({ medio: m, total: ventasDia.filter((v) => v.medio === m).reduce((s, v) => s + v.total, 0) })).filter((x) => x.total > 0);
  const porCat = CATEGORIAS.map((c) => ({ cat: c, total: ventasDia.filter((v) => v.cat === c).reduce((s, v) => s + v.total, 0) })).filter((x) => x.total > 0);

  return (
    <Section title="Resumen del Día" icon="resumen" accent="#c77dff">
      <Field label="Selecciona fecha">
        <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} style={{ maxWidth: 200 }} />
      </Field>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 24 }}>
        <StatCard label="Total Ventas" value={formatCOP(totalVentas)} color="#e8c547" />
        <StatCard label="Total Gastos" value={formatCOP(totalGastos)} color="#f07167" />
        <StatCard label="Utilidad" value={formatCOP(utilidad)} color={utilidad >= 0 ? "#56cfe1" : "#f07167"} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <MiniTable title="Por medio de pago" rows={porMedio} keyCol="medio" valCol="total" />
        <MiniTable title="Por categoría" rows={porCat} keyCol="cat" valCol="total" />
      </div>

      {ventasDia.length === 0 && gastosDia.length === 0 && (
        <p style={{ textAlign: "center", opacity: 0.45, marginTop: 32 }}>Sin registros para este día</p>
      )}
    </Section>
  );
}

// ── UI HELPERS ───────────────────────────────────────────────────────────────
function StatCard({ label, value, color }) {
  return (
    <div style={{ background: "rgba(255,255,255,0.04)", borderRadius: 14, padding: "16px 12px", textAlign: "center", border: `1px solid ${color}33` }}>
      <div style={{ fontSize: 11, opacity: 0.6, textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>{label}</div>
      <div style={{ fontWeight: 800, fontSize: 18, color }}>{value}</div>
    </div>
  );
}

function MiniTable({ title, rows, keyCol, valCol }) {
  return (
    <div style={{ background: "rgba(255,255,255,0.03)", borderRadius: 12, padding: 16, border: "1px solid rgba(255,255,255,0.08)" }}>
      <div style={{ fontWeight: 700, marginBottom: 12, fontSize: 13, opacity: 0.7 }}>{title}</div>
      {rows.length === 0 && <div style={{ opacity: 0.4, fontSize: 13 }}>Sin datos</div>}
      {rows.map((r) => (
        <div key={r[keyCol]} style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, fontSize: 13 }}>
          <span>{r[keyCol]}</span>
          <span style={{ fontWeight: 700 }}>{formatCOP(r[valCol])}</span>
        </div>
      ))}
    </div>
  );
}

function Section({ title, icon, accent, children }) {
  return (
    <div style={{ padding: "0 0 40px" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
        <span style={{ fontSize: 22 }}><Icon name={icon} /></span>
        <h2 style={{ margin: 0, fontSize: 20, fontWeight: 800, color: accent }}>{title}</h2>
      </div>
      {children}
    </div>
  );
}

function Modal({ title, onClose, onSave, accent, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: "#161622", borderRadius: 20, width: "100%", maxWidth: 480, padding: 24, border: `1px solid ${accent}44`, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 20 }}>
          <h3 style={{ margin: 0, color: accent, fontSize: 17 }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#f0e6d3", cursor: "pointer", fontSize: 18 }}>✕</button>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>{children}</div>
        <div style={{ display: "flex", gap: 10, marginTop: 22 }}>
          <button onClick={onClose} style={{ flex: 1, padding: "12px 0", borderRadius: 10, border: "1px solid rgba(255,255,255,0.15)", background: "transparent", color: "#f0e6d3", cursor: "pointer", fontFamily: "inherit" }}>Cancelar</button>
          <button onClick={onSave} style={{ flex: 2, padding: "12px 0", borderRadius: 10, border: "none", background: accent, color: "#0d1117", fontWeight: 800, cursor: "pointer", fontFamily: "inherit", fontSize: 15 }}>Guardar</button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div>
      <label style={{ display: "block", fontSize: 12, opacity: 0.6, marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.5 }}>{label}</label>
      <div style={{ width: "100%" }}>{children}</div>
    </div>
  );
}

function Row({ children }) {
  return <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>{children}</div>;
}

function Total({ label, value, color = "#e8c547" }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "rgba(255,255,255,0.05)", borderRadius: 10, padding: "12px 16px", marginTop: 4 }}>
      <span style={{ opacity: 0.6, fontSize: 13 }}>{label}</span>
      <span style={{ fontWeight: 800, fontSize: 20, color }}>{value}</span>
    </div>
  );
}

function Tag({ children, soft }) {
  return (
    <span style={{ display: "inline-block", padding: "2px 8px", borderRadius: 20, fontSize: 11, fontWeight: 600, background: soft ? "rgba(255,255,255,0.07)" : "rgba(232,197,71,0.15)", color: soft ? "rgba(240,230,211,0.7)" : "#e8c547", whiteSpace: "nowrap" }}>
      {children}
    </span>
  );
}

function TableWrap({ children }) {
  return (
    <div style={{ overflowX: "auto", borderRadius: 14, border: "1px solid rgba(255,255,255,0.08)" }}>
      {children}
    </div>
  );
}

// ── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [ventas, setVentas] = useStorage(STORAGE_KEYS.ventas, []);
  const [gastos, setGastos] = useStorage(STORAGE_KEYS.gastos, []);
  const [inventario, setInventario] = useStorage(STORAGE_KEYS.inventario, []);
  const [tab, setTab] = useState("ventas");
  const [toastMsg, setToastMsg] = useState(null);
  const toast = (msg) => setToastMsg(msg);

  const tabs = [
    { id: "ventas", label: "Ventas", icon: "🛍️" },
    { id: "gastos", label: "Gastos", icon: "💸" },
    { id: "inventario", label: "Inventario", icon: "📦" },
    { id: "resumen", label: "Resumen", icon: "📊" },
  ];

  const alertCount = inventario.filter((p) => (p.stock || 0) <= (p.stockMin || STOCK_MINIMO)).length;

  return (
    <>
      <style>{`
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #0d1117; color: #f0e6d3; font-family: 'Georgia', serif; }
        input, select { width: 100%; background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.15); border-radius: 10px; padding: 10px 14px; color: #f0e6d3; font-family: inherit; font-size: 14px; outline: none; transition: border 0.2s; }
        input:focus, select:focus { border-color: #e8c547; }
        select option { background: #161622; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th { padding: 10px 14px; text-align: left; font-size: 11px; opacity: 0.5; text-transform: uppercase; letter-spacing: 0.8px; border-bottom: 1px solid rgba(255,255,255,0.08); }
        td { padding: 11px 14px; border-bottom: 1px solid rgba(255,255,255,0.05); vertical-align: middle; }
        tr:last-child td { border-bottom: none; }
        tr:hover td { background: rgba(255,255,255,0.025); }
        .fab { display: inline-flex; align-items: center; gap: 6px; padding: 10px 20px; background: #e8c547; color: #0d1117; border: none; border-radius: 30px; font-weight: 800; font-family: inherit; font-size: 14px; cursor: pointer; margin-bottom: 16px; transition: opacity 0.2s; }
        .fab:hover { opacity: 0.85; }
        .del-btn { background: none; border: none; color: rgba(240,113,103,0.5); cursor: pointer; font-size: 14px; transition: color 0.2s; }
        .del-btn:hover { color: #f07167; }
        ::-webkit-scrollbar { width: 4px; height: 4px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.15); border-radius: 4px; }
      `}</style>

      {/* Header */}
      <div style={{ background: "rgba(13,17,23,0.95)", backdropFilter: "blur(12px)", position: "sticky", top: 0, zIndex: 50, borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
        <div style={{ maxWidth: 800, margin: "0 auto", padding: "14px 20px" }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 10, justifyContent: "space-between" }}>
            <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
              <h1 style={{ fontSize: 18, fontWeight: 900, color: "#e8c547", letterSpacing: -0.5 }}>Cositas pa Sumerce</h1>
              <span style={{ fontSize: 12, opacity: 0.45 }}>Gestión de ventas</span>
            </div>
            <button
            onClick={() => exportToExcel(ventas, gastos, inventario, () => toast("✓ Archivo listo para compartir"), () => toast("No se pudo compartir, intenta de nuevo"))}
              style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", background: "rgba(86,207,225,0.12)", border: "1px solid #56cfe1", borderRadius: 20, color: "#56cfe1", fontFamily: "inherit", fontSize: 12, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}
            >
              ⬇️ Exportar Excel
            </button>
          </div>
        </div>
        {/* Tabs */}
        <div style={{ maxWidth: 800, margin: "0 auto", display: "flex", padding: "0 20px" }}>
          {tabs.map((t) => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              flex: 1, background: "none", border: "none", color: tab === t.id ? "#e8c547" : "rgba(240,230,211,0.45)",
              fontFamily: "inherit", fontSize: 12, fontWeight: tab === t.id ? 800 : 500,
              padding: "10px 4px", cursor: "pointer", borderBottom: tab === t.id ? "2px solid #e8c547" : "2px solid transparent",
              transition: "all 0.2s", display: "flex", alignItems: "center", justifyContent: "center", gap: 4,
              position: "relative",
            }}>
              <span>{t.icon}</span>
              <span style={{ display: "none" }}>{t.label}</span>
              {t.id === "inventario" && alertCount > 0 && (
                <span style={{ position: "absolute", top: 6, right: 4, background: "#f07167", color: "#fff", borderRadius: 10, fontSize: 9, fontWeight: 900, padding: "1px 5px", lineHeight: 1.4 }}>{alertCount}</span>
              )}
            </button>
          ))}
        </div>
      </div>

      {/* Content */}
      <div style={{ maxWidth: 800, margin: "0 auto", padding: "24px 16px" }}>
        {tab === "ventas" && <Ventas ventas={ventas} setVentas={setVentas} inventario={inventario} setInventario={setInventario} toast={toast} />}
        {tab === "gastos" && <Gastos gastos={gastos} setGastos={setGastos} toast={toast} />}
        {tab === "inventario" && <Inventario inventario={inventario} setInventario={setInventario} toast={toast} />}
        {tab === "resumen" && <Resumen ventas={ventas} gastos={gastos} />}
      </div>

      {toastMsg && <Toast msg={toastMsg} onClose={() => setToastMsg(null)} />}
    </>
  );
}
