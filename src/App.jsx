import { useState, useMemo, useCallback, useRef } from "react";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip as RTooltip, Legend, ResponsiveContainer,
} from "recharts";
import * as XLSX from "xlsx";
import {
  LogOut, ChevronLeft, HelpCircle, Users, UserPlus, UserMinus, UserCog,
  AlertTriangle, ClipboardList, LayoutDashboard, Search, Pencil, Bell,
  Download, Copy, ShieldCheck, Trash2, Eye, EyeOff, Check, X, CalendarDays, Clock3,
} from "lucide-react";

const LOGO_SRC = `${import.meta.env.BASE_URL}logo.png`;

/* =========================================================================
   SAMU METROPOLITANO — Sistema de gestión de datos operacionales
   Prototipo de interfaz (React). Todos los datos son ficticios / simulados.
   ========================================================================= */

/* ---------- estilos globales ---------- */
const GlobalStyle = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

    .samu-root {
      --navy: #142347;
      --navy-2: #1c3268;
      --navy-deep: #0a1530;
      --yellow: #FFC933;
      --yellow-deep: #E8A800;
      --paper: #F3F5FA;
      --card: #FFFFFF;
      --line: #DCE1EE;
      --ink: #142347;
      --ink-soft: #4C5876;
      --ok: #1E8E5A;
      --warn: #C6511B;
      --danger: #C0362C;
      font-family: 'Inter', system-ui, sans-serif;
      color: var(--ink);
      background: var(--paper);
      min-height: 640px;
      width: 100%;
      box-sizing: border-box;
      position: relative;
    }
    .samu-root * { box-sizing: border-box; }
    .samu-root h1, .samu-root h2, .samu-root h3, .samu-root .disp {
      font-family: 'Barlow Condensed', sans-serif;
      letter-spacing: 0.01em;
      color: var(--navy);
    }
    .samu-root table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
    .samu-root th, .samu-root td { padding: 6px 9px; border: 1px solid var(--line); text-align: center; white-space: nowrap; }
    .samu-root th {
      background: var(--navy);
      color: #fff;
      font-family: 'Barlow Condensed', sans-serif;
      font-weight: 600;
      font-size: 12px;
      letter-spacing: 0.03em;
      text-transform: uppercase;
      position: sticky; top: 0;
    }
    .samu-root tr:nth-child(even) td { background: #F7F9FD; }
    .samu-root .table-wrap { overflow: auto; max-height: 360px; border: 1px solid var(--line); border-radius: 6px; }

    .hazard {
      height: 6px;
      background: repeating-linear-gradient(-45deg, var(--yellow) 0 14px, var(--navy) 14px 28px);
      flex-shrink: 0;
    }
    .btn {
      font-family: 'Inter', sans-serif;
      font-weight: 600;
      font-size: 13.5px;
      border-radius: 6px;
      border: 1.5px solid transparent;
      padding: 9px 16px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 7px;
      transition: transform .08s ease, filter .12s ease;
    }
    .btn:active { transform: translateY(1px); }
    .btn:disabled { opacity: 0.45; cursor: not-allowed; }
    .btn-primary { background: var(--yellow); color: var(--navy-deep); border-color: var(--yellow-deep); }
    .btn-primary:hover:not(:disabled) { filter: brightness(1.04); }
    .btn-navy { background: var(--navy); color: #fff; }
    .btn-navy:hover:not(:disabled) { background: var(--navy-2); }
    .btn-outline { background: #fff; color: var(--navy); border-color: var(--line); }
    .btn-outline:hover:not(:disabled) { border-color: var(--navy); }
    .btn-danger { background: #fff; color: var(--danger); border-color: var(--danger); }
    .btn-danger:hover:not(:disabled) { background: #FCEEEC; }
    .btn-sm { padding: 6px 10px; font-size: 12.5px; }

    .card {
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 10px;
      padding: 18px 20px;
    }
    .menu-btn {
      display: flex; align-items: center; gap: 12px;
      background: #fff; border: 1.5px solid var(--line); border-radius: 10px;
      padding: 16px 18px; cursor: pointer; text-align: left;
      font-family: 'Barlow Condensed', sans-serif; font-size: 19px; font-weight: 600;
      color: var(--navy); transition: border-color .12s ease, background .12s ease;
    }
    .menu-btn:hover:not(:disabled) { border-color: var(--yellow-deep); background: #FFFBF0; }
    .menu-btn:disabled { opacity: 0.45; cursor: not-allowed; }
    .menu-btn .ic { color: var(--yellow-deep); flex-shrink: 0; }

    .input, .select {
      font-family: 'Inter', sans-serif;
      border: 1.5px solid var(--line);
      border-radius: 6px;
      padding: 8px 10px;
      font-size: 13.5px;
      color: var(--ink);
      background: #fff;
      width: 100%;
    }
    .input:focus, .select:focus { outline: 2px solid var(--yellow-deep); outline-offset: 1px; border-color: var(--yellow-deep); }
    .label { font-size: 12px; font-weight: 600; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 4px; display: block; }

    .badge { font-size: 11px; font-weight: 700; padding: 2px 8px; border-radius: 20px; text-transform: uppercase; letter-spacing: 0.03em; }
    .badge-admin { background: var(--navy); color: var(--yellow); }
    .badge-user { background: #E7ECF7; color: var(--navy); }
    .badge-ok { background: #E4F5EC; color: var(--ok); }
    .badge-fail { background: #FBE9E7; color: var(--danger); }

    .tabbar { display: flex; gap: 4px; border-bottom: 2px solid var(--line); margin-bottom: 16px; flex-wrap: wrap; }
    .tab { padding: 9px 16px; cursor: pointer; font-family: 'Barlow Condensed', sans-serif; font-weight: 600; font-size: 15.5px; color: var(--ink-soft); border-bottom: 3px solid transparent; margin-bottom: -2px; }
    .tab.active { color: var(--navy); border-color: var(--yellow-deep); }

    .qmark { display: inline-flex; align-items: center; justify-content: center; width: 17px; height: 17px; border-radius: 50%; background: #E7ECF7; color: var(--navy); cursor: pointer; vertical-align: middle; }
    .pop { position: absolute; z-index: 30; background: var(--navy-deep); color: #fff; font-size: 12.5px; line-height: 1.4; padding: 10px 12px; border-radius: 8px; max-width: 260px; box-shadow: 0 8px 24px rgba(10,21,48,.35); }

    .topbar { display:flex; align-items:center; justify-content: space-between; padding: 10px 20px; background: var(--navy); }
    .topbar .who { color:#EAF0FF; font-size: 13px; }
    .toast { position: fixed; bottom: 18px; right: 18px; background: var(--navy-deep); color:#fff; padding: 12px 18px; border-radius: 8px; font-size: 13.5px; box-shadow: 0 8px 24px rgba(0,0,0,.3); z-index: 50; display:flex; align-items:center; gap:8px; }
    .alert-banner { display:flex; gap:10px; align-items:flex-start; background:#FFF6DE; border:1.5px solid var(--yellow-deep); border-radius: 8px; padding: 12px 14px; font-size: 13px; color: #6B4A00; }
    .scroll-x { overflow-x: auto; }
    .pill { font-size: 11.5px; padding: 3px 9px; border-radius: 20px; background:#E7ECF7; color: var(--navy); font-weight:600; }
  `}</style>
);

/* ---------- utilidades ---------- */
const pad2 = (n) => String(n).padStart(2, "0");
const fmtDMY = (iso) => { const [y,m,d] = iso.split("-"); return `${d}/${m}/${y}`; };
const rnd = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
const MONTHS = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

function dateRange(startISO, endISO) {
  const out = [];
  let d = new Date(startISO + "T00:00:00");
  const end = new Date(endISO + "T00:00:00");
  if (isNaN(d) || isNaN(end) || d > end) return out;
  while (d <= end) {
    out.push(d.toISOString().slice(0, 10));
    d.setDate(d.getDate() + 1);
  }
  return out.slice(0, 62);
}

function downloadCSV(filename, rows) {
  const csv = rows.map(r => r.map(c => {
    const s = String(c ?? "");
    return /[",;\n]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s;
  }).join(";")).join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function downloadXLSX(filename, rows, sheetName = "Datos") {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

function copyRowsToClipboard(rows) {
  const tsv = rows.map(r => r.join("\t")).join("\n");
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tsv).catch(() => {});
  }
  return tsv;
}

/* ---------- generación de datos simulados ---------- */
const CCKALL_USERS = ["FSilva", "CMorales", "PVenegas"];

function genPrimariosRows(dates) {
  return dates.map(date => {
    const BCA = rnd(0,50), BOR = rnd(0,50), ACA = rnd(0,50), AOR = rnd(0,50), GT222 = rnd(0,50), Nulos = rnd(0,10);
    return {
      date, BCA, BOR, ACA, AOR, GT222, Nulos,
      totBasicas: BCA+BOR, totAvanzadas: ACA+AOR,
      totValidas: BCA+BOR+ACA+AOR+GT222,
      totValidasNulos: BCA+BOR+ACA+AOR+GT222+Nulos,
      centroAsistencial: BCA+ACA, otroResultado: BOR+AOR,
    };
  });
}

function genSecundariosRows(dates) {
  return dates.map(date => {
    const c420 = rnd(0,50), QTA = rnd(0,50), SR = rnd(0,50), GT222 = rnd(0,50), TA = rnd(0,50), Nulos = rnd(0,8);
    return { date, c420, QTA, SR, GT222, TA, Nulos, totFichasNulos: c420+QTA+SR+Nulos, totValidasREM: c420 };
  });
}

const BASES = [
  { id: "B12", tipo: "Básica" }, { id: "B20", tipo: "Básica" }, { id: "B07", tipo: "Avanzada" },
  { id: "B15", tipo: "Básica" }, { id: "B03", tipo: "Avanzada" }, { id: "B18", tipo: "Básica" },
];
function genBuzonRows() {
  return BASES.map((b, i) => {
    const movil = 100 + i * 7 + rnd(1, 6);
    const trasladados = rnd(10,80), nst = rnd(10,80);
    const qtaOm = rnd(10,80), logistica = rnd(10,80);
    const totInt = trasladados + nst, totMov = qtaOm + logistica;
    return { movil, tipo: b.tipo, base: b.id, trasladados, nst, totInt, qtaOm, logistica, totMov, totGeneral: totInt + totMov };
  });
}

const REF_DAY = {
  totalLlamadas: 1711, contestadas: 1615, contAntes10: 1329, contAntes20: 175, contDespues20: 111,
  perdidas: 43, perdAntes10: 21, perdSobre10: 12, perdSobre20: 6, perdSobre30: 2, perdSobre40: 1, perdSobre50: 0, perdSobre60: 1,
  totAbandonadas: 53, abanAntes10: 21, abanSobre10: 32, promAbandono: "0:00:06", nivelServicio: 82, nivelAtencion: 94,
};
function scaleDayFromRef(factor) {
  const s = (v) => Math.max(0, Math.round(v * factor + rnd(-8,8)));
  return {
    totalLlamadas: s(REF_DAY.totalLlamadas), contestadas: s(REF_DAY.contestadas), contAntes10: s(REF_DAY.contAntes10),
    contAntes20: s(REF_DAY.contAntes20), contDespues20: s(REF_DAY.contDespues20), perdidas: s(REF_DAY.perdidas),
    perdAntes10: s(REF_DAY.perdAntes10), perdSobre10: s(REF_DAY.perdSobre10), perdSobre20: s(REF_DAY.perdSobre20),
    perdSobre30: s(REF_DAY.perdSobre30), perdSobre40: s(REF_DAY.perdSobre40), perdSobre50: 0, perdSobre60: rnd(0,1),
    totAbandonadas: s(REF_DAY.totAbandonadas), abanAntes10: s(REF_DAY.abanAntes10), abanSobre10: s(REF_DAY.abanSobre10),
    promAbandono: `0:00:${pad2(rnd(4,12))}`, nivelServicio: Math.min(99, Math.max(60, REF_DAY.nivelServicio + rnd(-6,6))),
    nivelAtencion: Math.min(99, Math.max(70, REF_DAY.nivelAtencion + rnd(-4,4))),
  };
}
function sumTelefonicoRange(dates) {
  const days = dates.map((d, i) => ({ date: d, ...(i === 0 ? REF_DAY : scaleDayFromRef(0.9 + rnd(-15,15)/100)) }));
  const acc = {};
  Object.keys(REF_DAY).forEach(k => {
    if (k === "promAbandono" || k === "nivelServicio" || k === "nivelAtencion") return;
    acc[k] = days.reduce((s, d) => s + d[k], 0);
  });
  acc.nivelServicio = Math.round(days.reduce((s,d)=>s+d.nivelServicio,0) / days.length);
  acc.nivelAtencion = Math.round(days.reduce((s,d)=>s+d.nivelAtencion,0) / days.length);
  acc.promAbandono = "0:00:0" + rnd(4,9);
  return { days, acc };
}

const ETL_LABELS = ["Buzón de Intervención", "Traslados Primarios", "Traslados Secundarios", "Registro Telefónico"];
function genEtlActivity() {
  const now = new Date();
  return ETL_LABELS.map((a, i) => {
    const t = new Date(now.getTime() - i * 3600 * 1000 * (1 + rnd(0,3)));
    const ok = Math.random() > 0.12;
    return { actividad: a, fecha: t.toLocaleDateString("es-CL"), hora: t.toLocaleTimeString("es-CL"), resultado: ok ? "Éxito" : "Fracaso" };
  });
}
function genErrores(n) {
  const fuentes = ["Traslado_Primario_junio.xlsx", "Traslado_Secundario_mayo.xlsx", "Buzon_Intervencion.xlsx", "CCKALL (Centro Regulador)"];
  const tipos = ["Fallas de conexión con CCKALL", "Indisponibilidad de CCKALL", "Errores de autenticación al CCKALL",
    "Credenciales inválidas o expiradas", "Tipo de dato incorrecto en celda numérica", "Formato de fecha inválido (no dd/mm/aa)",
    "Archivo sin extensión .xlsx", "Tiempo de espera agotado", "Indisponibilidad del DW"];
  const rows = [];
  const base = new Date();
  for (let i = 0; i < n; i++) {
    const t = new Date(base.getTime() - i * rnd(1,9) * 3600 * 1000);
    rows.push({ fecha: t.toLocaleDateString("es-CL"), hora: t.toLocaleTimeString("es-CL"),
      fuente: fuentes[rnd(0,fuentes.length-1)], tipo: tipos[rnd(0,tipos.length-1)] });
  }
  return rows;
}

/* ---------- usuarios iniciales del sistema ---------- */
const INITIAL_USERS = [
  { username: "administrador_tic", correo: "jefe.tic@samu.cl", consulta: false, modificacion: false, isAdmin: true, ultimo: "2026-08-07" },
  { username: "estadistica_samu", correo: "estadistica@samu.cl", consulta: true, modificacion: false, isAdmin: false, ultimo: "2026-08-06" },
  { username: "jroman", correo: "j.roman@samu.cl", consulta: true, modificacion: true, isAdmin: false, ultimo: "2026-08-05" },
  { username: "valarcon", correo: "v.alarcon@samu.cl", consulta: false, modificacion: false, isAdmin: false, ultimo: "2026-06-14" },
];

/* ==================== SUBCOMPONENTES ==================== */

function Logo({ size = 46 }) {
  return <img src={LOGO_SRC} alt="Escudo SAMU Metropolitano" style={{ width: size, height: size * 461/431, objectFit: "contain" }} />;
}

function Qmark({ text }) {
  const [open, setOpen] = useState(false);
  return (
    <span style={{ position: "relative", display: "inline-block" }}>
      <span className="qmark" onClick={() => setOpen(o => !o)} onBlur={() => setOpen(false)} tabIndex={0}>
        <HelpCircle size={12} />
      </span>
      {open && <div className="pop" style={{ top: 22, left: 0 }}>{text}</div>}
    </span>
  );
}

function Shell({ children, session, onBack, onLogout, showBack, title }) {
  return (
    <div className="samu-root" style={{ display: "flex", flexDirection: "column" }}>
      <div className="hazard" />
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 22px 10px", background: "var(--navy)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{ background: "#fff", borderRadius: 8, padding: 5, display: "flex" }}><Logo size={56} /></div>
          <div>
            <div className="disp" style={{ color: "#fff", fontSize: 20, fontWeight: 700, lineHeight: 1 }}>SAMU METROPOLITANO</div>
            <div style={{ color: "#B9C6E8", fontSize: 11.5, letterSpacing: "0.06em" }}>{title || "SISTEMA DE GESTIÓN DE DATOS ESTADÍSTICOS"}</div>
          </div>
        </div>
        {session && (
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{ textAlign: "right" }}>
              <div style={{ color: "#EAF0FF", fontSize: 13, fontWeight: 600 }}>{session.username}</div>
              <span className={session.isAdmin ? "badge badge-admin" : "badge badge-user"}>{session.isAdmin ? "Administrador" : "Estadística"}</span>
            </div>
            {showBack && (
              <button className="btn btn-outline btn-sm" style={{ background: "transparent", color: "#fff", borderColor: "#3A4E85" }} onClick={onBack}>
                <ChevronLeft size={15} /> Volver
              </button>
            )}
            <button className="btn btn-primary btn-sm" onClick={onLogout}><LogOut size={14} /> Cerrar sesión</button>
          </div>
        )}
      </div>
      <div style={{ flex: 1, padding: 22, overflow: "auto" }}>{children}</div>
    </div>
  );
}

/* ---------- LOGIN ---------- */
function LoginScreen({ users, onLogin }) {
  const [role, setRole] = useState("estadistica_samu");
  const [password, setPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState("");
  const hasAdmin = users.some(u => u.isAdmin);

  function submit() {
    if (!password) { setError("Ingrese su contraseña."); return; }
    const u = users.find(u => u.username === role);
    onLogin(u);
  }

  return (
    <div className="samu-root" style={{ minHeight: 640, display: "flex", flexDirection: "column" }}>
      <div className="hazard" />
      <div style={{ position: "absolute", top: 22, left: 22, background: "#fff", borderRadius: 8, padding: 4, boxShadow: "0 2px 8px rgba(20,35,71,.15)" }}>
        <Logo size={44} />
      </div>
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", background: "linear-gradient(180deg, var(--navy) 0%, var(--navy-2) 100%)" }}>
        <div className="card" style={{ width: 380 }}>
          <div style={{ textAlign: "center", marginBottom: 18 }}>
            <div className="disp" style={{ fontSize: 24, fontWeight: 800 }}>Iniciar sesión</div>
            <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>Sistema de Gestión de Datos Estadísticos</div>
          </div>

          <label className="label">Usuario</label>
          <select className="select" value={role} onChange={e => setRole(e.target.value)} style={{ marginBottom: 14 }}>
            <option value="estadistica_samu">estadistica_samu</option>
            <option value="administrador_tic">administrador_tic</option>
          </select>

          <label className="label">Contraseña</label>
          <div style={{ position: "relative", marginBottom: 6 }}>
            <input className="input" type={showPw ? "text" : "password"} value={password}
              onChange={e => setPassword(e.target.value)} placeholder="" />
            <span onClick={() => setShowPw(s => !s)} style={{ position: "absolute", right: 10, top: 9, cursor: "pointer", color: "var(--ink-soft)" }}>
              {showPw ? <EyeOff size={16} /> : <Eye size={16} />}
            </span>
          </div>
          {error && <div style={{ color: "var(--danger)", fontSize: 12.5, marginBottom: 8 }}>{error}</div>}

          <button className="btn btn-primary" style={{ width: "100%", justifyContent: "center", marginTop: 10 }} onClick={submit}>
            Ingresar
          </button>

          {!hasAdmin && (
            <button className="btn btn-outline" style={{ width: "100%", justifyContent: "center", marginTop: 10 }}>
              <UserPlus size={15} /> Crear usuario administrador
            </button>
          )}
          <div style={{ marginTop: 16, fontSize: 11, color: "var(--ink-soft)", textAlign: "center" }}>
            Correo institucional @samu.cl requerido para cuentas nuevas.
          </div>
        </div>
      </div>
      <div style={{ textAlign: "center", padding: "10px", color: "#B9C6E8", fontSize: 11, background: "var(--navy-deep)" }}>
        SAMU Metropolitano · Centro Regulador 131
      </div>
    </div>
  );
}

/* ---------- MENÚ PRINCIPAL ---------- */
function MainMenu({ session, users, go }) {
  const etl = useMemo(() => genEtlActivity(), []);
  const inactivos = useMemo(() => {
    const cutoff = new Date("2026-08-07"); cutoff.setDate(cutoff.getDate() - 30);
    return users.filter(u => new Date(u.ultimo) < cutoff);
  }, [users]);

  const items = [
    { key: "consultar", label: "Consultar datos", icon: Search, disabled: !session.consulta, note: !session.consulta ? "Sin permiso de consulta" : "" },
    { key: "dashboard", label: "Ver dashboard", icon: LayoutDashboard },
    { key: "modificar", label: "Modificar datos", icon: Pencil, disabled: !session.modificacion, note: !session.modificacion ? "Sin permiso de modificación" : "" },
  ];
  const adminItems = [
    { key: "usuarios", label: "Ver usuarios registrados", icon: Users },
    { key: "crearUsuario", label: "Crear usuario", icon: UserPlus },
    { key: "eliminarUsuario", label: "Eliminar usuario", icon: UserMinus },
    { key: "modificarUsuario", label: "Modificar usuario", icon: UserCog },
    { key: "errores", label: "Ver errores", icon: AlertTriangle },
    { key: "bitacora", label: "Ver bitácora de usuarios", icon: ClipboardList },
  ];

  return (
    <div>
      {session.isAdmin && inactivos.length > 0 && (
        <div className="alert-banner" style={{ marginBottom: 18 }}>
          <Bell size={18} style={{ flexShrink: 0, marginTop: 1 }} />
          <div>
            <b>{inactivos.length} usuario(s)</b> no han iniciado sesión en más de 30 días: {" "}
            {inactivos.map(u => u.username).join(", ")}. Revise si corresponde eliminarlos del sistema.
          </div>
        </div>
      )}

      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 10 }}>Menú principal</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px,1fr))", gap: 12, marginBottom: 26 }}>
        {items.map(it => (
          <button key={it.key} className="menu-btn" disabled={it.disabled} title={it.note} onClick={() => go(it.key)}>
            <it.icon className="ic" size={22} />
            <div>
              {it.label}
              {it.note && <div style={{ fontSize: 11, fontFamily: "Inter", fontWeight: 400, color: "var(--ink-soft)" }}>{it.note}</div>}
            </div>
          </button>
        ))}
      </div>

      {session.isAdmin && (
        <>
          <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 10 }}>Administración</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px,1fr))", gap: 12, marginBottom: 26 }}>
            {adminItems.map(it => (
              <button key={it.key} className="menu-btn" onClick={() => go(it.key)}>
                <it.icon className="ic" size={22} /> {it.label}
              </button>
            ))}
          </div>
        </>
      )}

      <div className="card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div className="disp" style={{ fontSize: 16, fontWeight: 700 }}>Actividades ETL — últimas 24 horas</div>
          <span className="pill">Actualizado automáticamente</span>
        </div>
        <div className="table-wrap">
          <table>
            <thead><tr><th>Actividad</th><th>Fecha</th><th>Hora</th><th>Resultado</th></tr></thead>
            <tbody>
              {etl.map((r, i) => (
                <tr key={i}>
                  <td style={{ textAlign: "left" }}>{r.actividad}</td>
                  <td>{r.fecha}</td><td>{r.hora}</td>
                  <td><span className={r.resultado === "Éxito" ? "badge badge-ok" : "badge badge-fail"}>{r.resultado}</span></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ---------- TABLA CON EXPORTAR / COPIAR ---------- */
function ExportBar({ rows, filenameBase }) {
  return (
    <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
      <button className="btn btn-outline btn-sm" onClick={() => downloadCSV(`${filenameBase}.csv`, rows)}><Download size={13} /> CSV</button>
      <button className="btn btn-outline btn-sm" onClick={() => downloadXLSX(`${filenameBase}.xlsx`, rows)}><Download size={13} /> XLSX</button>
      <button className="btn btn-outline btn-sm" onClick={() => copyRowsToClipboard(rows)}><Copy size={13} /> Copiar tabla</button>
    </div>
  );
}

/* ---------- CONSULTAR DATOS ---------- */
function InputConIcono({ type, style, ...props }) {
  const inputRef = useRef(null);
  const Icono = type === "time" ? Clock3 : CalendarDays;

  function abrirSelector() {
    const input = inputRef.current;
    if (!input) return;

    input.focus();
    // showPicker abre el selector nativo cuando el navegador lo admite.
    // El click es el respaldo para navegadores que no lo implementan.
    if (typeof input.showPicker === "function") input.showPicker();
    else input.click();
  }

  return (
    <div style={{ position: "relative" }}>
      <input ref={inputRef} type={type} className="input" {...props} style={{ paddingRight: 38, ...style }} />
      <button
        type="button"
        aria-label={type === "time" ? "Abrir selector de hora" : "Abrir selector de fecha"}
        onClick={abrirSelector}
        style={{ position: "absolute", right: 4, top: "50%", transform: "translateY(-50%)", display: "grid", placeItems: "center", width: 30, height: 30, padding: 0, border: 0, background: "transparent", color: "var(--ink-soft)", cursor: "pointer" }}
      >
        <Icono size={16} aria-hidden="true" />
      </button>
    </div>
  );
}

function ConsultarDatos({ log }) {
  const [tipo, setTipo] = useState("primarios");
  const [ini, setIni] = useState("2026-06-01");
  const [fin, setFin] = useState("2026-06-07");
  const [horaIni, setHoraIni] = useState("00:00");
  const [horaFin, setHoraFin] = useState("23:59");
  const [mesIni, setMesIni] = useState({ y: 2026, m: 5 });
  const [mesFin, setMesFin] = useState({ y: 2026, m: 5 });
  const [ran, setRan] = useState(false);

  const dates = useMemo(() => dateRange(ini, fin), [ini, fin, ran]);
  const primarios = useMemo(() => genPrimariosRows(dates), [ran]);
  const secundarios = useMemo(() => genSecundariosRows(dates), [ran]);
  const buzon = useMemo(() => genBuzonRows(), [ran]);
  const tele = useMemo(() => sumTelefonicoRange(dates.length ? dates : [ini]), [ran]);

  const tipos = [
    { key: "primarios", label: "Primarios" },
    { key: "secundarios", label: "Traslados Secundarios" },
    { key: "buzon", label: "Buzón de Intervención" },
    { key: "telefonico", label: "Registro Telefónico" },
  ];

  const registrarConsulta = () => {
    const etiqueta = tipos.find(t => t.key === tipo)?.label || tipo;
    log(`Consulta de datos: ${etiqueta}`);
    setRan(true);
  };

  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Consultar datos</div>

      <div className="card" style={{ marginBottom: 20, textAlign: "left"  }}>
        <label className="label">Tipo de dato</label>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
          {tipos.map(t => (
            <button key={t.key} className={`btn btn-sm ${tipo === t.key ? "btn-primary" : "btn-outline"}`} onClick={() => { setTipo(t.key); setRan(false); }}>
              {t.label}
            </button>
          ))}
        </div>

        {tipo === "buzon" ? (
          <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
            <div>
              <label className="label">Mes inicial</label>
              <div style={{ display: "flex", gap: 6 }}>
                <select className="select" value={mesIni.m} onChange={e => setMesIni(s => ({ ...s, m: +e.target.value }))}>
                  {MONTHS.map((m, i) => <option key={i} value={i}>{m}</option>)}
                </select>
                <select className="select" value={mesIni.y} onChange={e => setMesIni(s => ({ ...s, y: +e.target.value }))}>
                  <option value={2025}>2025</option><option value={2026}>2026</option>
                </select>
              </div>
            </div>
            <div>
              <label className="label">Mes final</label>
              <div style={{ display: "flex", gap: 6 }}>
                <select className="select" value={mesFin.m} onChange={e => setMesFin(s => ({ ...s, m: +e.target.value }))}>
                  {MONTHS.map((m, i) => <option key={i} value={i}>{m}</option>)}
                </select>
                <select className="select" value={mesFin.y} onChange={e => setMesFin(s => ({ ...s, y: +e.target.value }))}>
                  <option value={2025}>2025</option><option value={2026}>2026</option>
                </select>
              </div>
            </div>
          </div>
        ) : (
          <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
            <div>
              <label className="label">Fecha de inicio</label>
              <InputConIcono type="date" value={ini} onChange={e => setIni(e.target.value)} />
            </div>
            <div>
              <label className="label">Fecha de término</label>
              <InputConIcono type="date" value={fin} onChange={e => setFin(e.target.value)} />
            </div>
            {tipo === "telefonico" && (
              <>
                <div>
                  <label className="label">Hora inicial (solo registro telefónico)</label>
                  <InputConIcono type="time" value={horaIni} onChange={e => setHoraIni(e.target.value)} />
                </div>
                <div>
                  <label className="label">Hora final (solo registro telefónico)</label>
                  <InputConIcono type="time" value={horaFin} onChange={e => setHoraFin(e.target.value)} />
                </div>
              </>
            )}
          </div>
        )}

        <button className="btn btn-primary" style={{ marginTop: 16 }} onClick={registrarConsulta}>
          <Search size={15} /> Consultar
        </button>
      </div>

      {ran && tipo === "primarios" && <TablaPrimarios rows={primarios} />}
      {ran && tipo === "secundarios" && <TablaSecundarios rows={secundarios} />}
      {ran && tipo === "buzon" && <TablaBuzon rows={buzon} />}
      {ran && tipo === "telefonico" && <TablasTelefonico data={tele} />}
    </div>
  );
}

function TablaPrimarios({ rows }) {
  const head = ["Fecha","Básicas Centro Asistencial (BCA)","Básicas Otro Resultado (BOR)","Avanzada Centro Asistencial (ACA)","Avanzada Otro Resultado (AOR)","GT 222","Nulos","Totales Básicas","Totales Avanzadas","Total Fichas Válidas","Total Fichas + Nulos","Centro Asist.","Otro Resultado"];
  const csvRows = [head, ...rows.map(r => [fmtDMY(r.date), r.BCA, r.BOR, r.ACA, r.AOR, r.GT222, r.Nulos, r.totBasicas, r.totAvanzadas, r.totValidas, r.totValidasNulos, r.centroAsistencial, r.otroResultado])];
  const remJ = rows.reduce((s,r)=>s+r.BCA+r.BOR+r.ACA+r.AOR+r.GT222,0);
  const remLBasico = rows.reduce((s,r)=>s+r.BCA,0);
  const remLAvanzado = rows.reduce((s,r)=>s+r.ACA,0);
  return (
    <>
      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "18px 0 6px" }}>Tabla 1 — Detalle diario</div>
      <ExportBar rows={csvRows} filenameBase="traslados_primarios" />
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr>{head.map(h=><th key={h}>{h}</th>)}</tr></thead>
          <tbody>{rows.length===0 ? <tr><td colSpan={head.length}>NA — sin datos para el período</td></tr> : rows.map(r=>(
            <tr key={r.date}><td>{fmtDMY(r.date)}</td><td>{r.BCA}</td><td>{r.BOR}</td><td>{r.ACA}</td><td>{r.AOR}</td><td>{r.GT222}</td><td>{r.Nulos}</td>
              <td>{r.totBasicas}</td><td>{r.totAvanzadas}</td><td>{r.totValidas}</td><td>{r.totValidasNulos}</td><td>{r.centroAsistencial}</td><td>{r.otroResultado}</td></tr>
          ))}</tbody>
        </table>
      </div>

      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "18px 0 6px" }}>Tabla 2 — Indicadores REM</div>
      <div className="table-wrap" style={{ maxWidth: 340 }}>
        <table>
          <thead><tr><th>Indicador</th><th>Total período</th></tr></thead>
          <tbody>
            <tr><td style={{textAlign:"left"}}>REM J</td><td>{remJ}</td></tr>
            <tr><td style={{textAlign:"left"}}>REM L Básico</td><td>{remLBasico}</td></tr>
            <tr><td style={{textAlign:"left"}}>REM L Avanzado</td><td>{remLAvanzado}</td></tr>
          </tbody>
        </table>
      </div>
    </>
  );
}

function TablaSecundarios({ rows }) {
  const head = ["Fecha","420","QTA","SR (Sin Registro)","GT 222","TA (Traslado Aéreo)","Nulos","Totales Fichas + Nulos","Totales Válidas (REM)"];
  const csvRows = [head, ...rows.map(r => [fmtDMY(r.date), r.c420, r.QTA, r.SR, r.GT222, r.TA, r.Nulos, r.totFichasNulos, r.totValidasREM])];
  return (
    <>
      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "18px 0 6px" }}>Traslados Secundarios — Detalle diario</div>
      <ExportBar rows={csvRows} filenameBase="traslados_secundarios" />
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr>{head.map(h=><th key={h}>{h}</th>)}</tr></thead>
          <tbody>{rows.length===0 ? <tr><td colSpan={head.length}>NA — sin datos para el período</td></tr> : rows.map(r=>(
            <tr key={r.date}><td>{fmtDMY(r.date)}</td><td>{r.c420}</td><td>{r.QTA}</td><td>{r.SR}</td><td>{r.GT222}</td><td>{r.TA}</td><td>{r.Nulos}</td><td>{r.totFichasNulos}</td><td>{r.totValidasREM}</td></tr>
          ))}</tbody>
        </table>
      </div>
    </>
  );
}

function TablaBuzon({ rows }) {
  const head = ["Móvil","Tipo Ambulancia","Base","Interv. Trasladados","Intervención Pacientes No Trasladados (NST)","Tot. Intervenciones","Mov. QTA-OM","Mov. Logística","Tot. Movimientos","Total General"];
  const csvRows = [head, ...rows.map(r => [r.movil, r.tipo, r.base, r.trasladados, r.nst, r.totInt, r.qtaOm, r.logistica, r.totMov, r.totGeneral])];
  return (
    <>
      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "18px 0 6px" }}>Buzón de Intervención — por móvil</div>
      <ExportBar rows={csvRows} filenameBase="buzon_intervencion" />
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr>{head.map(h=><th key={h}>{h}</th>)}</tr></thead>
          <tbody>{rows.map(r=>(
            <tr key={r.movil}><td>{r.movil}</td><td>{r.tipo}</td><td>{r.base}</td><td>{r.trasladados}</td><td>{r.nst}</td><td>{r.totInt}</td><td>{r.qtaOm}</td><td>{r.logistica}</td><td>{r.totMov}</td><td><b>{r.totGeneral}</b></td></tr>
          ))}</tbody>
        </table>
      </div>
    </>
  );
}

const IND_HELP = {
  totalLlamadas: "Llamadas entrantes con contraparte mayor a 6 dígitos.",
  contestadas: "Estado Número Final = 'contestada' y Duración = 0 (columna Duración usada para cálculos sensibles al tiempo).",
  perdidas: "Llamadas entrantes perdidas con contraparte mayor a 6 dígitos y Tiempo Ring > 0.",
  totAbandonadas: "Tiempo de cola mayor a 4 segundos.",
  nivelServicio: "Porcentaje de llamadas contestadas antes de 10 s contra el total de contestadas.",
  nivelAtencion: "Porcentaje de contestadas contra el total, sin contar no-contestadas antes de 10 s ni perdidas en cola.",
};



const COMUNAS_RM = [
  "Desconocido", "Cerrillos", "Cerro Navia", "Conchalí", "El Bosque", "Estación Central", "Huechuraba",
  "Independencia", "La Cisterna", "La Florida", "La Granja", "La Pintana", "La Reina",
  "Las Condes", "Lo Barnechea", "Lo Espejo", "Lo Prado", "Macul", "Maipú", "Ñuñoa",
  "Pedro Aguirre Cerda", "Peñalolén", "Providencia", "Pudahuel", "Puente Alto", "Quilicura",
  "Quinta Normal", "Recoleta", "Renca", "San Bernardo", "San Joaquín", "San Miguel",
  "San Ramón", "Santiago", "Vitacura", "Pirque", "San José de Maipo", "Colina", "Lampa",
  "Tiltil", "Buin", "Calera de Tango", "Paine", "Alhué", "Curacaví", "María Pinto",
  "Melipilla", "San Pedro", "El Monte", "Isla de Maipo", "Padre Hurtado", "Peñaflor", "Talagante"
];

const INDICADORES_COMUNA = [
  ["totalLlamadas", "Llamadas totales recibidas"], ["contestadas", "Llamadas contestadas"],
  ["contAntes10", "Llamadas contestadas antes 10 segundos"], ["contAntes20", "Llamadas contestadas antes 20 segundos"],
  ["contDespues20", "Contestadas después 20 segundos"], ["perdidas", "Perdidas"],
  ["perdAntes10", "Perdidas antes 10 segundos"], ["perdSobre10", "Perdidas después 10 segundos"],
  ["perdSobre20", "Perdidas después 20 segundos"], ["perdSobre30", "Perdidas después 30 segundos"],
  ["perdSobre40", "Perdidas después de 40 segundos"], ["perdSobre50", "Perdidas después 50 segundos"],
  ["perdSobre60", "Perdidas después 60 segundos"], ["totAbandonadas", "Abandonadas"],
  ["abanAntes10", "Abandonadas antes 10 segundos"], ["abanSobre10", "Abandonadas después 10 segundos"],
];

// Distribuye cada total del período sin perder registros: toda columna suma el indicador general.
function distribuirTotalPorComuna(total, semilla) {
  const pesos = COMUNAS_RM.map((_, i) => 80 + ((i * 37 + semilla * 19) % 73));
  const sumaPesos = pesos.reduce((s, peso) => s + peso, 0);
  const partes = pesos.map((peso, i) => {
    const exacto = total * peso / sumaPesos;
    return { i, valor: Math.floor(exacto), resto: exacto % 1 };
  });
  const pendientes = total - partes.reduce((s, parte) => s + parte.valor, 0);
  partes.sort((a, b) => b.resto - a.resto || a.i - b.i);
  for (let i = 0; i < pendientes; i++) partes[i].valor++;
  return partes.sort((a, b) => a.i - b.i).map(parte => parte.valor);
}

function construirFilasPorComuna(acc) {
  const distribuciones = Object.fromEntries(INDICADORES_COMUNA.map(([key], i) => [key, distribuirTotalPorComuna(acc[key], i + 1)]));
  return COMUNAS_RM.map((comuna, i) => ({ comuna, ...Object.fromEntries(INDICADORES_COMUNA.map(([key]) => [key, distribuciones[key][i]])) }));
}

const TIEMPOS_LLAMADAS_131 = [
  { categoria: "> 0 s y ≤ 30 s", etiqueta: ["> 0 s", "≤ 30 s"], min: 5, q1: 12, q2: 19, q3: 25, max: 30, media: 19, outliers: [2] },
  { categoria: "> 30 s y ≤ 60 s", etiqueta: ["> 30 s", "≤ 60 s"], min: 31, q1: 38, q2: 45, q3: 52, max: 59, media: 46, outliers: [60] },
  { categoria: "> 60 s", etiqueta: ["> 60 s"], min: 62, q1: 74, q2: 92, q3: 122, max: 157, media: 98, outliers: [173, 186] },
];

function formatoMMSS(segundos) {
  return `${pad2(Math.floor(segundos / 60))}:${pad2(segundos % 60)}`;
}

function DiagramaCajasTiempos131() {
  const ancho = 760, alto = 390, margen = { arriba: 24, derecha: 20, abajo: 70, izquierda: 70 };
  const maximoEje = 210;
  const y = valor => margen.arriba + (maximoEje - valor) * (alto - margen.arriba - margen.abajo) / maximoEje;
  const centros = [220, 410, 600];
  const ticks = [0, 30, 60, 90, 120, 150, 180, 210];
  const colorCaja = "#7fc3c2";
  const colorMedia = "#870c0e";

  return (
    <div className="card" style={{ marginTop: 22, padding: "18px 16px 12px", overflowX: "auto" }}>
      <div className="disp" style={{ fontSize: 16, fontWeight: 700, marginBottom: 2 }}>Tiempos de llamadas recibidas al 131</div>
      <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginBottom: 10 }}>Distribución por duración de llamada. La línea roja representa la media.</div>
      <svg viewBox={`0 0 ${ancho} ${alto}`} role="img" aria-label="Tres diagramas de caja y bigotes verticales de tiempos de llamadas recibidas al 131" style={{ display: "block", width: "100%", minWidth: 620, height: "auto" }}>
        {ticks.map(tick => <g key={tick}>
          <line x1={margen.izquierda} x2={ancho - margen.derecha} y1={y(tick)} y2={y(tick)} stroke="#DCE1EE" strokeDasharray="3 3" />
          <text x={margen.izquierda - 10} y={y(tick) + 4} textAnchor="end" fill="#4C5876" fontSize="12">{formatoMMSS(tick)}</text>
        </g>)}
        <line x1={margen.izquierda} x2={margen.izquierda} y1={margen.arriba} y2={alto - margen.abajo} stroke="#142347" />
        <line x1={margen.izquierda} x2={ancho - margen.derecha} y1={alto - margen.abajo} y2={alto - margen.abajo} stroke="#142347" />
        <text transform={`translate(18 ${(alto - margen.abajo + margen.arriba) / 2}) rotate(-90)`} textAnchor="middle" fill="#142347" fontSize="13" fontWeight="600">Tiempo (mm:ss)</text>
        {TIEMPOS_LLAMADAS_131.map((item, i) => {
          const cx = centros[i], anchoCaja = 72, tapa = 18;
          return <g key={item.categoria}>
            <line x1={cx} x2={cx} y1={y(item.max)} y2={y(item.q3)} stroke="#142347" strokeWidth="1.5" />
            <line x1={cx - tapa} x2={cx + tapa} y1={y(item.max)} y2={y(item.max)} stroke="#142347" strokeWidth="1.5" />
            <line x1={cx} x2={cx} y1={y(item.q1)} y2={y(item.min)} stroke="#142347" strokeWidth="1.5" />
            <line x1={cx - tapa} x2={cx + tapa} y1={y(item.min)} y2={y(item.min)} stroke="#142347" strokeWidth="1.5" />
            <rect x={cx - anchoCaja / 2} y={y(item.q3)} width={anchoCaja} height={y(item.q1) - y(item.q3)} fill={colorCaja} fillOpacity="0.82" stroke="#142347" strokeWidth="1.5" />
            <line x1={cx - anchoCaja / 2} x2={cx + anchoCaja / 2} y1={y(item.q2)} y2={y(item.q2)} stroke="#142347" strokeWidth="2" />
            <line x1={cx - anchoCaja / 2} x2={cx + anchoCaja / 2} y1={y(item.media)} y2={y(item.media)} stroke={colorMedia} strokeWidth="3" />
            {item.outliers.map((outlier, j) => <circle key={j} cx={cx + (j ? 12 : 0)} cy={y(outlier)} r="4" fill="#142347" />)}
            <text x={cx} y={alto - 42} textAnchor="middle" fill="#142347" fontSize="12" fontWeight="600">
              {item.etiqueta.map((linea, j) => <tspan key={linea} x={cx} dy={j ? 14 : 0}>{linea}</tspan>)}
            </text>
          </g>;
        })}
        <text x={(margen.izquierda + ancho - margen.derecha) / 2} y={alto - 8} textAnchor="middle" fill="#142347" fontSize="13" fontWeight="600">Categoría de duración</text>
      </svg>
      <div style={{ display: "flex", gap: 16, justifyContent: "center", fontSize: 12, color: "var(--ink-soft)", marginTop: 3 }}>
        <span><i style={{ display: "inline-block", width: 11, height: 11, marginRight: 5, verticalAlign: "-1px", background: colorCaja, border: "1px solid #142347" }} />Caja y bigotes</span>
        <span><i style={{ display: "inline-block", width: 17, borderTop: `3px solid ${colorMedia}`, marginRight: 5, verticalAlign: "3px" }} />Media</span>
        <span><i style={{ display: "inline-block", width: 7, height: 7, marginRight: 5, verticalAlign: "1px", borderRadius: "50%", background: "#142347" }} />Outlier</span>
      </div>
    </div>
  );
}

function TablasTelefonico({ data }) {
  const { days, acc } = data;
  const filasPorComuna = useMemo(() => construirFilasPorComuna(acc), [acc]);
  const repRows = [
    ["Indicador","Valor"],
    ["Total llamadas", acc.totalLlamadas], ["Llamadas contestadas", acc.contestadas],
    ["Contestadas antes de 10 s", acc.contAntes10], ["Contestadas antes de 20 s", acc.contAntes20],
    ["Contestadas después de 20 s", acc.contDespues20], ["Llamadas perdidas", acc.perdidas],
    ["Perdidas antes 10 s", acc.perdAntes10], ["Perdidas sobre 10 s", acc.perdSobre10], ["Perdidas sobre 20 s", acc.perdSobre20],
    ["Perdidas sobre 30 s", acc.perdSobre30], ["Perdidas sobre 40 s", acc.perdSobre40], ["Perdidas sobre 50 s", acc.perdSobre50],
    ["Perdidas sobre 60 s", acc.perdSobre60], ["Total abandonadas", acc.totAbandonadas],
    ["Abandonadas antes 10 s", acc.abanAntes10], ["Abandonadas sobre 10 s", acc.abanSobre10],
    ["Promedio tiempo abandono", acc.promAbandono], ["Nivel servicio", acc.nivelServicio + "%"], ["Nivel atención", acc.nivelAtencion + "%"],
  ];

  const perOperador = CCKALL_USERS.map(u => {
    const contestadas = rnd(300, 560), perdidas = rnd(10,60);
    return { u, contestadas, perdidas, prom: `0:0${rnd(1,4)}:${pad2(rnd(0,59))}`, a10: rnd(200,contestadas), a20: rnd(0,60), sobre10: rnd(10,80) };
  });

  const cortadas = { porOperador: rnd(80, 260), porCliente: rnd(400, 900), perdidas: acc.perdidas, abandonadas: acc.totAbandonadas };
  const pctOperador = Math.round(cortadas.porOperador * 100 / acc.contestadas);
  const pctCliente = Math.round(cortadas.porCliente * 100 / acc.contestadas);

  const desgloseHoras = Array.from({ length: 24 }, (_, h) => {
    const total = rnd(20, 110);
    const contestadas = Math.round(total * (0.85 + rnd(-8,8)/100));
    return { h, total, contestadas, perdidas: total - contestadas };
  });

  const resumenTiempos = [
    ["Categoría", "Primer cuartil", "Segundo cuartil", "Tercer cuartil", "Media", "Mínimo", "Máximo"],
    ...TIEMPOS_LLAMADAS_131.map(t => [t.categoria, formatoMMSS(t.q1), formatoMMSS(t.q2), formatoMMSS(t.q3), formatoMMSS(t.media), formatoMMSS(t.min), formatoMMSS(t.max)]),
  ];


  return (
    <div>
      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "18px 0 6px", display:"flex", alignItems:"center", gap:6 }}>
        Módulo de Reporte de Llamadas <Qmark text={IND_HELP.totalLlamadas} />
      </div>
      <ExportBar rows={repRows} filenameBase="registro_telefonico_reporte" />
      <div className="table-wrap" style={{ maxWidth: 420 }}>
        <table><tbody>
          {repRows.slice(1).map(([k,v]) => <tr key={k}><td style={{textAlign:"left", fontWeight:600}}>{k}</td><td>{v}</td></tr>)}
        </tbody></table>
      </div>

      <DiagramaCajasTiempos131 />
      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "22px 0 6px" }}>Resumen estadístico — tiempos de llamadas al 131</div>
      <ExportBar rows={resumenTiempos} filenameBase="registro_telefonico_tiempos_131" />
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr>{resumenTiempos[0].map(encabezado => <th key={encabezado}>{encabezado}</th>)}</tr></thead>
          <tbody>{resumenTiempos.slice(1).map(fila => <tr key={fila[0]}>{fila.map((celda, i) => <td key={i} style={i === 0 ? { textAlign: "left", fontWeight: 600 } : undefined}>{celda}</td>)}</tr>)}</tbody>
        </table>
      </div>

      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "22px 0 6px" }}>Indicadores por Usuario (CCKALL)</div>
      <ExportBar rows={[["Usuario","Contestadas","Perdidas","Prom. Hablado","Antes 10s","Antes 20s","Sobre 10s"], ...perOperador.map(o=>[o.u,o.contestadas,o.perdidas,o.prom,o.a10,o.a20,o.sobre10])]} filenameBase="registro_telefonico_usuarios" />
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr><th>Usuario</th><th>Contestadas</th><th>Perdidas</th><th>Prom. Hablado</th><th>Antes 10s</th><th>Antes 20s</th><th>Sobre 10s</th></tr></thead>
          <tbody>{perOperador.map(o=>(
            <tr key={o.u}><td>{o.u}</td><td>{o.contestadas}</td><td>{o.perdidas}</td><td>{o.prom}</td><td>{o.a10}</td><td>{o.a20}</td><td>{o.sobre10}</td></tr>
          ))}</tbody>
        </table>
      </div>

      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "22px 0 6px" }}>Cortadas por Operador / Cliente</div>
      <ExportBar rows={[["Indicador","Valor"],["Contestadas cortadas por operador",cortadas.porOperador],["% sobre contestadas",pctOperador+"%"],["Contestadas cortadas por cliente",cortadas.porCliente],["% sobre contestadas",pctCliente+"%"],["Perdidas",cortadas.perdidas],["Abandonadas",cortadas.abandonadas]]} filenameBase="registro_telefonico_cortadas" />
      <div className="table-wrap" style={{ maxWidth: 380 }}>
        <table><tbody>
          <tr><td style={{textAlign:"left"}}>Contestadas cortadas por operador</td><td>{cortadas.porOperador}</td></tr>
          <tr><td style={{textAlign:"left"}}>% sobre contestadas</td><td>{pctOperador}%</td></tr>
          <tr><td style={{textAlign:"left"}}>Contestadas cortadas por cliente</td><td>{cortadas.porCliente}</td></tr>
          <tr><td style={{textAlign:"left"}}>% sobre contestadas</td><td>{pctCliente}%</td></tr>
          <tr><td style={{textAlign:"left"}}>Perdidas</td><td>{cortadas.perdidas}</td></tr>
          <tr><td style={{textAlign:"left"}}>Abandonadas</td><td>{cortadas.abandonadas}</td></tr>
        </tbody></table>
      </div>
      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "22px 0 6px" }}>Desglose por Hora — {fmtDMY(days[days.length-1].date)}</div>
      <ExportBar rows={[["Hora","Total","Contestadas","Perdidas"], ...desgloseHoras.map(r=>[`${pad2(r.h)}:00`, r.total, r.contestadas, r.perdidas])]} filenameBase="registro_telefonico_horas" />
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr><th>Hora</th><th>Total</th><th>Contestadas</th><th>Perdidas</th></tr></thead>
          <tbody>{desgloseHoras.map(r=>(
            <tr key={r.h}><td>{pad2(r.h)}:00</td><td>{r.total}</td><td>{r.contestadas}</td><td>{r.perdidas}</td></tr>
          ))}</tbody>
        </table>
      </div>

      <div className="disp" style={{ fontSize: 15, fontWeight: 700, margin: "22px 0 6px" }}>Indicadores por Comuna — Región Metropolitana</div>
      <ExportBar
        rows={[["Comuna", ...INDICADORES_COMUNA.map(([, label]) => label)], ...filasPorComuna.map(fila => [fila.comuna, ...INDICADORES_COMUNA.map(([key]) => fila[key])])]}
        filenameBase="registro_telefonico_comunas"
      />
      <div className="table-wrap scroll-x" style={{ maxHeight: 560, overflowY: "auto" }}>
        <table>
          <thead><tr><th>Comuna</th>{INDICADORES_COMUNA.map(([, label]) => <th key={label}>{label}</th>)}</tr></thead>
          <tbody>{filasPorComuna.map(fila => (
            <tr key={fila.comuna}>
              <td style={{ textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>{fila.comuna}</td>
              {INDICADORES_COMUNA.map(([key]) => <td key={key}>{fila[key]}</td>)}
            </tr>
          ))}</tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------- MODIFICAR DATOS ---------- */
function ModificarDatos({ log }) {
  const [tipo, setTipo] = useState("Traslado Primario");
  const [fecha, setFecha] = useState("2026-06-01");
  
  const camposPorTipo = {
    "Traslado Primario": ["Básicas Centro Asistencial (BCA)","Básicas Otro Resultado (BOR)","Avanzada Centro Asistencial (ACA)","Avanzada Otro Resultado (AOR)","GT 222","Nulos"],
    "Traslado Secundario": ["420","QTA","SR (Sin Registro)","GT 222","TA (Traslado Aéreo)","Nulos"],
    "Buzón de Intervención": ["Interv. Trasladados","Intervención Pacientes No Trasladados (NST)","Mov. QTA-OM","Mov. Logística"],
    "Registro Telefónico": ["Total llamadas", "Llamadas contestadas", "Llamadas perdidas", "Total abandonadas"],
  };
  
  const [campo, setCampo] = useState(camposPorTipo["Traslado Primario"][0]);
  const [nuevoValor, setNuevoValor] = useState("");
  const [resultado, setResultado] = useState(null);

  function guardar() {
    if (nuevoValor === "" || isNaN(Number(nuevoValor))) { setResultado({ ok: false, msg: "Debe ingresar un valor numérico válido." }); return; }
    const anterior = rnd(0, 50);
    setResultado({ ok: true, anterior, nuevo: Number(nuevoValor) });
    log(`Modificación de Datos de ${tipo.replace("Traslado ", "Traslado ")}`);
  }

  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Modificar datos</div>
      <div className="card" style={{ maxWidth: 480 }}>
        <label className="label">Tipo de dato</label>
        <select className="select" value={tipo} onChange={e => { setTipo(e.target.value); setCampo(camposPorTipo[e.target.value][0]); }} style={{ marginBottom: 14 }}>
          {Object.keys(camposPorTipo).map(t => <option key={t}>{t}</option>)}
        </select>

        <label className="label">Fecha</label>
        <InputConIcono type="date" value={fecha} onChange={e => setFecha(e.target.value)} style={{ marginBottom: 14 }} />

        <label className="label">Campo / procedimiento</label>
        <select className="select" value={campo} onChange={e => setCampo(e.target.value)} style={{ marginBottom: 14 }}>
          {camposPorTipo[tipo].map(c => <option key={c}>{c}</option>)}
        </select>

        <label className="label">Nuevo valor</label>
        <input className="input" value={nuevoValor} onChange={e => setNuevoValor(e.target.value)} placeholder="Ej: 24" style={{ marginBottom: 16 }} />

        <button className="btn btn-primary" onClick={guardar}>Guardar cambio</button>

        {resultado && (
          resultado.ok ? (
            <div style={{ marginTop: 16, background: "#E4F5EC", border: "1.5px solid var(--ok)", borderRadius: 8, padding: 12, fontSize: 13 }}>
              <b>Modificación exitosa.</b><br />
              {tipo} · {campo} · {fmtDMY(fecha)}<br />
              Valor anterior: <b>{resultado.anterior}</b> → Nuevo valor: <b>{resultado.nuevo}</b>
            </div>
          ) : (
            <div style={{ marginTop: 16, background: "#FBE9E7", border: "1.5px solid var(--danger)", borderRadius: 8, padding: 12, fontSize: 13 }}>
              <b>Error:</b> {resultado.msg}
            </div>
          )
        )}
      </div>
    </div>
  );
}

/* ---------- DASHBOARD ---------- */
function Dashboard() {
  const [tab, setTab] = useState("tel");
  const tabs = [
    { key: "tel", label: "Registro Telefónico" },
    { key: "prim", label: "Traslados Primarios" },
    { key: "sec", label: "Traslado Secundario" },
    { key: "buz", label: "Buzón de Intervención" },
  ];
  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Dashboard</div>
      <div className="tabbar">
        {tabs.map(t => <div key={t.key} className={`tab ${tab===t.key?"active":""}`} onClick={()=>setTab(t.key)}>{t.label}</div>)}
      </div>
      {tab === "tel" && <DashTelefonico />}
      {tab === "prim" && <DashTraslados 
        titulo="Traslados Primarios" 
        tipos={["Básicas Centro Asistencial (BCA)","Básicas Otro Resultado (BOR)","Avanzada Centro Asistencial (ACA)","Avanzada Otro Resultado (AOR)","GT 222","Nulos"]} 
        totales={["Totales Básicas", "Totales Avanzadas", "Total Fichas Válidas", "Total Fichas + Nulos"]}
      />}
      {tab === "sec" && <DashTraslados 
        titulo="Traslado Secundario" 
        tipos={["420","QTA","SR (Sin Registro)","GT 222","TA (Traslado Aéreo)","Nulos"]} 
        totales={["Totales Válidas (REM)", "Totales Fichas + Nulos"]}
      />}
      {tab === "buz" && <DashBuzon />}
    </div>
  );
}

function DashTelefonico() {
  const [mes, setMes] = useState(5);
  const [year1, setYear1] = useState("2025");
  const [year2, setYear2] = useState("2026");
  const nivelServicioDias = useMemo(() => Array.from({ length: 30 }, (_, i) => ({ dia: i + 1, valor: Math.min(99, Math.max(55, 82 + rnd(-14,12))) })), [mes]);
  const comparativoAnual = useMemo(() => [
    { mes: "Enero", 2025: 368, 2026: 1053 }, { mes: "Febrero", 2025: 661, 2026: 1119 },
    { mes: "Marzo", 2025: 510, 2026: 1098 }, { mes: "Abril", 2025: 699, 2026: 556 },
    { mes: "Mayo", 2025: 350, 2026: 335 }, { mes: "Junio", 2025: 1080, 2026: 307 },
    { mes: "Julio", 2025: 528, 2026: 527 },
  ], []);
  const compMensual = useMemo(() => MONTHS.map((m, i) => {
    const fila = comparativoAnual[i];
    return { mes: m.slice(0,3), actual: fila?.[2026] ?? null, anterior: fila?.[2025] ?? null };
  }), [comparativoAnual]);
  const porHora = useMemo(() => Array.from({ length: 24 }, (_, h) => {
    const recibidas = rnd(10, 110);
    return { h: `${pad2(h)}:00`, recibidas, contestadas: rnd(Math.round(recibidas * 0.65), recibidas) };
  }), []);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div className="disp" style={{ fontSize: 15, fontWeight: 700, display:"flex", alignItems:"center", gap:6 }}>
            Nivel de servicio por día <Qmark text={IND_HELP.nivelServicio} />
          </div>
          <select className="select" style={{ width: 160 }} value={mes} onChange={e=>setMes(+e.target.value)}>
            {MONTHS.map((m,i)=><option key={i} value={i}>{m}</option>)}
          </select>
        </div>
        <ResponsiveContainer width="100%" height={220}>
          <LineChart data={nivelServicioDias}>
            <CartesianGrid strokeDasharray="3 3" stroke="#E2E6F0" />
            <XAxis dataKey="dia" fontSize={11} /><YAxis fontSize={11} unit="%" domain={[0,100]} />
            <RTooltip /> <Line type="monotone" dataKey="valor" name="Nivel de servicio" stroke="#142347" strokeWidth={2} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <div className="card">
        <div className="disp" style={{ fontSize: 15, fontWeight: 700, marginBottom: 8 }}>Llamadas recibidas por mes — 2026 vs. 2025</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={compMensual}>
            <CartesianGrid strokeDasharray="3 3" stroke="#E2E6F0" />
            <XAxis dataKey="mes" fontSize={11} /><YAxis fontSize={11} />
            <RTooltip /><Legend />
            <Bar dataKey="actual" name="2026" fill="#FFC933" />
            <Bar dataKey="anterior" name="2025" fill="#142347" />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="card">
        <div className="disp" style={{ fontSize: 15, fontWeight: 700, marginBottom: 8 }}>Comparativo anual de llamadas totales recibidas al 131</div>
        <div className="table-wrap scroll-x">
          <table>
            <thead><tr>
              <th>Mes</th>
              {[year1, year2].map((year, index) => (
                <th key={index}><select value={year} onChange={e => index === 0 ? setYear1(e.target.value) : setYear2(e.target.value)} style={{ background: "transparent", color: "#fff", border: "1px solid #3A4E85", borderRadius: 4, outline: "none", padding: "2px 4px" }}>
                  <option value="2025" style={{ color: "#000" }}>2025</option><option value="2026" style={{ color: "#000" }}>2026</option>
                </select></th>
              ))}
              <th>% cambio</th>
            </tr></thead>
            <tbody>{comparativoAnual.map(fila => {
              const inicial = fila[year1], final = fila[year2];
              const cambio = (final - inicial) * 100 / inicial;
              return <tr key={fila.mes}><td style={{ textAlign: "left" }}>{fila.mes}</td><td>{inicial}</td><td>{final}</td><td style={{ color: cambio >= 0 ? "var(--ok)" : "var(--danger)", fontWeight: 600 }}>{`${cambio >= 0 ? "+" : ""}${cambio.toFixed(1)}%`}</td></tr>;
            })}</tbody>
          </table>
        </div>
      </div>

      <div className="card">
        <div className="disp" style={{ fontSize: 15, fontWeight: 700, marginBottom: 8 }}>Llamadas por hora — día anterior</div>
        <ResponsiveContainer width="100%" height={220}>
          <LineChart data={porHora}>
            <CartesianGrid strokeDasharray="3 3" stroke="#E2E6F0" />
            <XAxis dataKey="h" fontSize={10} interval={2} /><YAxis fontSize={11} />
            <RTooltip /><Legend />
            <Line type="monotone" dataKey="recibidas" name="Recibidas" stroke="#142347" strokeWidth={2} dot={false} />
            <Line type="monotone" dataKey="contestadas" name="Contestadas" stroke="#E8A800" strokeWidth={2} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function DashTraslados({ titulo, tipos, totales }) {
  const allChartOptions = [...tipos, ...totales];
  const [selChart, setSelChart] = useState(allChartOptions[0]);
  const [selTable, setSelTable] = useState(tipos[0]);
  const [year1, setYear1] = useState("2025");
  const [year2, setYear2] = useState("2026");

  const evol = useMemo(() => Array.from({ length: 12 }, (_, i) => {
    const row = { mes: MONTHS[i].slice(0,3) };
    allChartOptions.forEach(t => row[t] = rnd(200,1400));
    return row;
  }), [titulo, allChartOptions]);

  const comparativo = useMemo(() => MONTHS.map(m => {
    const anterior = rnd(300,1200), actual = rnd(300,1200);
    return { mes: m, anterior, actual, pct: Math.round((actual-anterior)*1000/anterior)/10 };
  }), [selTable, year1, year2]);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div className="disp" style={{ fontSize: 15, fontWeight: 700 }}>{titulo} — evolución mensual por tipo</div>
          <select className="select" style={{ width: 320 }} value={selChart} onChange={e=>setSelChart(e.target.value)}>
            <optgroup label="Tipos de procedimiento">
              {tipos.map(t=><option key={t}>{t}</option>)}
            </optgroup>
            <optgroup label="Totales">
              {totales.map(t=><option key={t}>{t}</option>)}
            </optgroup>
          </select>
        </div>
        <ResponsiveContainer width="100%" height={240}>
          <LineChart data={evol}>
            <CartesianGrid strokeDasharray="3 3" stroke="#E2E6F0" />
            <XAxis dataKey="mes" fontSize={11} /><YAxis fontSize={11} />
            <RTooltip /><Legend />
            <Line type="monotone" dataKey={selChart} stroke="#142347" strokeWidth={2.5} dot={{r:3}} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <div className="card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div className="disp" style={{ fontSize: 15, fontWeight: 700 }}>Tabla comparativa anual</div>
          <select className="select" style={{ width: 320 }} value={selTable} onChange={e=>setSelTable(e.target.value)}>
            {tipos.map(t=><option key={t}>{t}</option>)}
          </select>
        </div>
        <ExportBar rows={[["Mes", year1, year2, "% cambio"], ...comparativo.map(c=>[c.mes,c.anterior,c.actual,(isFinite(c.pct)?c.pct+"%":"NA")])]} filenameBase={`comparativo_${titulo.replace(/\s/g,"_")}`} />
        <div className="table-wrap scroll-x">
          <table>
            <thead>
              <tr>
                <th>Mes</th>
                <th>
                  <select value={year1} onChange={e=>setYear1(e.target.value)} style={{background:"transparent", color:"#fff", border:"1px solid #3A4E85", borderRadius:4, outline:"none", padding: "2px 4px"}}>
                    <option value="2024" style={{color:"#000"}}>2024</option>
                    <option value="2025" style={{color:"#000"}}>2025</option>
                    <option value="2026" style={{color:"#000"}}>2026</option>
                  </select>
                </th>
                <th>
                  <select value={year2} onChange={e=>setYear2(e.target.value)} style={{background:"transparent", color:"#fff", border:"1px solid #3A4E85", borderRadius:4, outline:"none", padding: "2px 4px"}}>
                    <option value="2024" style={{color:"#000"}}>2024</option>
                    <option value="2025" style={{color:"#000"}}>2025</option>
                    <option value="2026" style={{color:"#000"}}>2026</option>
                  </select>
                </th>
                <th>% cambio</th>
              </tr>
            </thead>
            <tbody>{comparativo.map(c=>(
              <tr key={c.mes}><td style={{textAlign:"left"}}>{c.mes}</td><td>{c.anterior}</td><td>{c.actual}</td>
                <td style={{ color: c.pct>=0 ? "var(--ok)" : "var(--danger)", fontWeight:600 }}>{isFinite(c.pct)?`${c.pct>0?"+":""}${c.pct}%`:"NA"}</td></tr>
            ))}</tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function DashBuzon() {
  const IDS_MOVILES = [
    "112", "113", "114", "117/137", "119",
    "126", "127", "131", "133", "134", "135", "136",
    "139", "140", "142", "143", "144", "145", "146", "147",
    "148", "149", "150", "151", "152", "153", "154", "155",
    "156", "157", "158", "159", "160", "161", "162", "163",
    "164", "166", "167", "168", "169"
  ];

  const [movil, setMovil] = useState(IDS_MOVILES[0]);

  const gen = () =>
    Array.from({ length: 12 }, (_, i) => ({
      mes: MONTHS[i].slice(0, 3),
      valor: rnd(10, 80)
    }));

  const trasladados = useMemo(() => gen(), [movil]);
  const nst = useMemo(() => gen(), [movil]);
  const qtaOm = useMemo(() => gen(), [movil]);
  const logistica = useMemo(() => gen(), [movil]);

  const chart = (data, label, color) => (
    <div className="card">
      <div
        className="disp"
        style={{
          fontSize: 14,
          fontWeight: 700,
          marginBottom: 8
        }}
      >
        {label}
      </div>

      <ResponsiveContainer width="100%" height={170}>
        <LineChart data={data}>
          <CartesianGrid
            strokeDasharray="3 3"
            stroke="#E2E6F0"
          />

          <XAxis
            dataKey="mes"
            fontSize={10}
          />

          <YAxis
            fontSize={10}
          />

          <RTooltip />

          <Line
            type="monotone"
            dataKey="valor"
            stroke={color}
            strokeWidth={2}
            dot={{ r: 3 }}
          />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );

  return (
    <div>
      <div
        className="card"
        style={{
          marginBottom: 16
        }}
      >
        <label className="label">
          Filtrar por ID de móvil
        </label>

        <select
          className="select"
          style={{ width: 200 }}
          value={movil}
          onChange={e => setMovil(e.target.value)}
        >
          {IDS_MOVILES.map(id => (
            <option
              key={id}
              value={id}
            >
              {id}
            </option>
          ))}
        </select>
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns:
            "repeat(auto-fit, minmax(280px,1fr))",
          gap: 16
        }}
      >
        {chart(
          trasladados,
          "Intervenciones — Pacientes Trasladados",
          "#142347"
        )}

        {chart(
          nst,
          "Intervenciones — No Trasladados (NST)",
          "#C0362C"
        )}

        {chart(
          qtaOm,
          "Movimientos — Sin Contacto (QTA-OM)",
          "#E8A800"
        )}

        {chart(
          logistica,
          "Movimientos — Logística",
          "#1E8E5A"
        )}
      </div>
    </div>
  );
}

/* ---------- ADMIN: USUARIOS ---------- */
function VerUsuarios({ users }) {
  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Usuarios registrados</div>
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr><th>Nombre de usuario</th><th>Correo SAMU</th><th>Consulta de datos</th><th>Modificación de datos</th><th>Rol</th><th>Último ingreso</th></tr></thead>
          <tbody>{users.map(u=>(
            <tr key={u.username}>
              <td style={{textAlign:"left"}}>{u.username}</td><td style={{textAlign:"left"}}>{u.correo}</td>
              <td>{u.consulta ? <Check size={14} color="#1E8E5A"/> : <X size={14} color="#C0362C"/>}</td>
              <td>{u.modificacion ? <Check size={14} color="#1E8E5A"/> : <X size={14} color="#C0362C"/>}</td>
              <td><span className={u.isAdmin?"badge badge-admin":"badge badge-user"}>{u.isAdmin?"Admin":"Estadística"}</span></td>
              <td>{fmtDMY(u.ultimo)}</td>
            </tr>
          ))}</tbody>
        </table>
      </div>
    </div>
  );
}

function CrearUsuario({ onCreate, forceAdmin, toast }) {
  const [username, setUsername] = useState("");
  const [correo, setCorreo] = useState("");
  const [pw1, setPw1] = useState("");
  const [pw2, setPw2] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState("");

  function submit() {
    if (!username.trim()) return setError("Ingrese un nombre de usuario.");
    if (!/^[\w.\-]+@samu\.cl$/i.test(correo)) return setError("El correo debe ser institucional (@samu.cl).");
    if (pw1.length < 6) return setError("La contraseña debe tener al menos 6 caracteres.");
    if (pw1 !== pw2) return setError("Las contraseñas no coinciden.");
    setError("");
    onCreate({ username: username.trim(), correo, consulta: false, modificacion: false, isAdmin: !!forceAdmin, ultimo: "2026-08-07" });
    setUsername(""); setCorreo(""); setPw1(""); setPw2("");
    toast(`Usuario "${username.trim()}" creado sin permisos asignados.`);
  }

  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>{forceAdmin ? "Crear usuario administrador" : "Crear usuario"}</div>
      <div className="card" style={{ maxWidth: 440 }}>
        <label className="label">Nombre de usuario</label>
        <input className="input" value={username} onChange={e=>setUsername(e.target.value)} style={{ marginBottom: 14 }} />

        <label className="label">Correo SAMU (@samu.cl)</label>
        <input className="input" value={correo} onChange={e=>setCorreo(e.target.value)} placeholder="nombre.apellido@samu.cl" style={{ marginBottom: 14 }} />

        <label className="label">Contraseña</label>
        <div style={{ position: "relative", marginBottom: 14 }}>
          <input className="input" type={showPw?"text":"password"} value={pw1} onChange={e=>setPw1(e.target.value)} />
          <span onClick={()=>setShowPw(s=>!s)} style={{ position:"absolute", right:10, top:9, cursor:"pointer", color:"var(--ink-soft)" }}>{showPw?<EyeOff size={16}/>:<Eye size={16}/>}</span>
        </div>

        <label className="label">Confirmar contraseña</label>
        <input className="input" type={showPw?"text":"password"} value={pw2} onChange={e=>setPw2(e.target.value)} style={{ marginBottom: 14 }} />

        {error && <div style={{ color: "var(--danger)", fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
        <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginBottom: 14 }}>El usuario se creará sin permisos de consulta ni modificación. Un administrador deberá asignarlos.</div>
        <button className="btn btn-primary" onClick={submit}><UserPlus size={15}/> {forceAdmin ? "Crear administrador" : "Crear usuario"}</button>
      </div>
    </div>
  );
}

function EliminarUsuario({ users, onDelete, sessionUsername }) {
  const [pending, setPending] = useState(null);
  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Eliminar usuario</div>
      <div className="table-wrap scroll-x" style={{ maxWidth: 620 }}>
        <table>
          <thead><tr><th>Usuario</th><th>Correo</th><th>Acción</th></tr></thead>
          <tbody>{users.map(u=>(
            <tr key={u.username}>
              <td style={{textAlign:"left"}}>{u.username}</td><td style={{textAlign:"left"}}>{u.correo}</td>
              <td>
                {pending===u.username ? (
                  <span style={{ display:"flex", gap:6, justifyContent:"center" }}>
                    <button className="btn btn-danger btn-sm" onClick={()=>{onDelete(u.username); setPending(null);}}>Confirmar</button>
                    <button className="btn btn-outline btn-sm" onClick={()=>setPending(null)}>Cancelar</button>
                  </span>
                ) : (
                  <button className="btn btn-outline btn-sm" disabled={u.username===sessionUsername}
                    onClick={()=>setPending(u.username)}><Trash2 size={13}/> Eliminar</button>
                )}
              </td>
            </tr>
          ))}</tbody>
        </table>
      </div>
    </div>
  );
}

function ModificarUsuario({ users, onUpdate, sessionUsername, toast }) {
  const [sel, setSel] = useState(users[0]?.username || "");
  const u = users.find(x=>x.username===sel);
  const isSelf = sel === sessionUsername;

  function toggle(field) {
    onUpdate(sel, { [field]: !u[field] });
  }

  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Modificar usuario</div>
      <div className="card" style={{ maxWidth: 460 }}>
        <label className="label">Usuario</label>
        <select className="select" value={sel} onChange={e=>setSel(e.target.value)} style={{ marginBottom: 18 }}>
          {users.map(x=><option key={x.username} value={x.username}>{x.username}</option>)}
        </select>

        {u && (
          <>
            <div style={{ fontSize: 13, marginBottom: 10 }}>{u.correo}</div>
            {isSelf && <div className="alert-banner" style={{ marginBottom: 14 }}><AlertTriangle size={15}/> Un administrador no puede modificar sus propios permisos.</div>}
            <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, opacity: isSelf?0.5:1 }}>
              <input type="checkbox" checked={u.consulta} disabled={isSelf} onChange={()=>toggle("consulta")} /> Permiso de consulta de datos
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, opacity: isSelf?0.5:1 }}>
              <input type="checkbox" checked={u.modificacion} disabled={isSelf} onChange={()=>toggle("modificacion")} /> Permiso de modificación de datos
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16, opacity: isSelf?0.5:1 }}>
              <input type="checkbox" checked={u.isAdmin} disabled={isSelf} onChange={()=>{toggle("isAdmin"); toast(`Privilegios de administrador ${!u.isAdmin?"otorgados a":"revocados de"} ${u.username}.`);}} /> Rol Administrador
            </label>
            {/* <div style={{ fontSize: 11, color: "var(--ink-soft)" }}>Todos los cambios quedan registrados en la bitácora de auditoría.</div> */}
          </>
        )}
      </div>
    </div>
  );
}

/* ---------- VER ERRORES ---------- */
function VerErrores() {
  const rows = useMemo(() => genErrores(134), []);
  const perPage = 50;
  const [page, setPage] = useState(1);
  const totalPages = Math.ceil(rows.length / perPage);
  const slice = rows.slice((page-1)*perPage, page*perPage);
  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Errores del sistema (ETL / fuentes de datos)</div>
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr><th>Fecha</th><th>Hora</th><th>Fuente del error</th><th>Tipo de error</th></tr></thead>
          <tbody>{slice.map((r,i)=>(
            <tr key={i}><td>{r.fecha}</td><td>{r.hora}</td><td style={{textAlign:"left"}}>{r.fuente}</td><td style={{textAlign:"left"}}>{r.tipo}</td></tr>
          ))}</tbody>
        </table>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 12 }}>
        <span style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>Página {page} de {totalPages} · {rows.length} registros</span>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="btn btn-outline btn-sm" disabled={page===1} onClick={()=>setPage(p=>p-1)}>Anterior</button>
          <button className="btn btn-outline btn-sm" disabled={page===totalPages} onClick={()=>setPage(p=>p+1)}>Siguiente</button>
        </div>
      </div>
    </div>
  );
}

/* ---------- BITÁCORA ---------- */
function Bitacora({ entries }) {
  return (
    <div>
      <div className="disp" style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Bitácora de usuarios del sistema</div>
      <div className="table-wrap scroll-x">
        <table>
          <thead><tr><th>Fecha</th><th>Hora</th><th>Usuario</th><th>Actividad</th></tr></thead>
          <tbody>{entries.map((e,i)=>(
            <tr key={i}><td>{e.fecha}</td><td>{e.hora}</td><td>{e.usuario}</td><td style={{textAlign:"left"}}>{e.actividad}</td></tr>
          ))}</tbody>
        </table>
      </div>
    </div>
  );
}

/* ==================== APP ==================== */
export default function App() {
  const [users, setUsers] = useState(INITIAL_USERS);
  const [session, setSession] = useState(null);
  const [stack, setStack] = useState(["login"]);
  const [bitacora, setBitacora] = useState([]);
  const [toastMsg, setToastMsg] = useState(null);

  const screen = stack[stack.length - 1];
  const go = (s) => setStack(prev => [...prev, s]);
  const back = () => setStack(prev => prev.length > 1 ? prev.slice(0, -1) : prev);

  const toast = useCallback((msg) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(null), 3200);
  }, []);

  const logAct = useCallback((actividad) => {
    const now = new Date();
    setBitacora(prev => [{ fecha: now.toLocaleDateString("es-CL"), hora: now.toLocaleTimeString("es-CL"), usuario: session?.username || "-", actividad }, ...prev]);
  }, [session]);

  function handleLogin(u) {
    setSession(u);
    setStack(["main"]);
    setTimeout(() => {
      const now = new Date();
      setBitacora(prev => [{ fecha: now.toLocaleDateString("es-CL"), hora: now.toLocaleTimeString("es-CL"), usuario: u.username, actividad: "Inicio de Sesión" }, ...prev]);
    }, 0);
  }
  function handleLogout() {
    if (session) {
      const now = new Date();
      setBitacora(prev => [{ fecha: now.toLocaleDateString("es-CL"), hora: now.toLocaleTimeString("es-CL"), usuario: session.username, actividad: "Cierre de Sesión" }, ...prev]);
    }
    setSession(null);
    setStack(["login"]);
  }
  function updateUser(username, patch) {
    setUsers(prev => prev.map(u => u.username === username ? { ...u, ...patch } : u));
    if (session?.username === username) setSession(s => ({ ...s, ...patch }));
  }
  function deleteUser(username) {
    setUsers(prev => prev.filter(u => u.username !== username));
    toast(`Usuario "${username}" eliminado.`);
  }
  function createUser(u) {
    setUsers(prev => [...prev, u]);
  }

  function navGo(key) {
    const labelMap = { dashboard: "Ver Dashboard" };
    if (labelMap[key]) logAct(labelMap[key]);
    go(key);
  }

  if (!session) {
    return <><GlobalStyle /><LoginScreen users={users} onLogin={handleLogin} /></>;
  }

  const titleMap = {
    main: "MENÚ PRINCIPAL", consultar: "CONSULTAR DATOS", dashboard: "DASHBOARD", modificar: "MODIFICAR DATOS",
    usuarios: "USUARIOS REGISTRADOS", crearUsuario: "CREAR USUARIO", eliminarUsuario: "ELIMINAR USUARIO",
    modificarUsuario: "MODIFICAR USUARIO", errores: "ERRORES DEL SISTEMA", bitacora: "BITÁCORA",
  };

  return (
    <>
      <GlobalStyle />
      <Shell session={session} onBack={back} onLogout={handleLogout} showBack={screen !== "main"} title={titleMap[screen]}>
        {screen === "main" && <MainMenu session={session} users={users} go={navGo} />}
        {screen === "consultar" && <ConsultarDatos log={logAct} />}
        {screen === "dashboard" && <Dashboard />}
        {screen === "modificar" && <ModificarDatos log={logAct} />}
        {screen === "usuarios" && <VerUsuarios users={users} />}
        {screen === "crearUsuario" && <CrearUsuario onCreate={createUser} toast={toast} />}
        {screen === "eliminarUsuario" && <EliminarUsuario users={users} onDelete={deleteUser} sessionUsername={session.username} />}
        {screen === "modificarUsuario" && <ModificarUsuario users={users} onUpdate={updateUser} sessionUsername={session.username} toast={toast} />}
        {screen === "errores" && <VerErrores />}
        {screen === "bitacora" && <Bitacora entries={bitacora} />}
      </Shell>
      {toastMsg && <div className="toast"><ShieldCheck size={16} color="#FFC933" /> {toastMsg}</div>}
    </>
  );
}
