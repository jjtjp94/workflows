import { useState, useRef, useCallback, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════════════════
   CONSTANTS
   ═══════════════════════════════════════════════════════════════════════════ */
const PX_PER_HOUR = 90;
const LANE_H = 40;
const ROW_PAD = 6;
const MIN_ROW_H = LANE_H + ROW_PAD * 2;
const ROLE_W = 170;
const HEADER_H = 48;
const TOTAL_MINUTES = 1440;
const TOTAL_W = (TOTAL_MINUTES / 60) * PX_PER_HOUR;
const RESIZE_SNAP = 15; // always 15-min resize regardless of grid

const LOGO_URL = "https://play-lh.googleusercontent.com/EJBaGftC07UbvTwLiZ6edqzV84F9S78NVsVKSyzcGolDkDpkA93g3xsdbBezlz171w=w240-h480-rw";

const minToPx = (m) => (m / 60) * PX_PER_HOUR;
const pxToMin = (px) => (px / PX_PER_HOUR) * 60;
const snapMin = (m, s) => Math.round(m / s) * s;
const fmtTime = (m) => {
  const h = Math.floor(((m % TOTAL_MINUTES) + TOTAL_MINUTES) % TOTAL_MINUTES / 60);
  const mm = Math.floor(m % 60);
  const ap = h >= 12 ? "p" : "a";
  const h12 = h === 0 ? 12 : h > 12 ? h - 12 : h;
  return mm === 0 ? `${h12}${ap}` : `${h12}:${String(mm).padStart(2, "0")}${ap}`;
};
const fmtTimeFull = (m) => {
  const h = Math.floor(((m % TOTAL_MINUTES) + TOTAL_MINUTES) % TOTAL_MINUTES / 60);
  const mm = Math.floor(m % 60);
  const ap = h >= 12 ? "PM" : "AM";
  const h12 = h === 0 ? 12 : h > 12 ? h - 12 : h;
  return `${h12}:${String(mm).padStart(2, "0")} ${ap}`;
};
const fmtDur = (m) => {
  if (m < 60) return `${m}m`;
  const h = Math.floor(m / 60);
  const r = m % 60;
};

let _uid = 0;
const uid = () => `_${++_uid}_${Date.now().toString(36)}`;

const DURATIONS = [15, 30, 45, 60, 90, 120, 180, 240];

const COLORS = [
  "#3b82f6", "#8b5cf6", "#06b6d4", "#10b981", "#f59e0b", "#ef4444",
  "#ec4899", "#0ea5e9", "#6366f1", "#dc2626", "#f97316", "#14b8a6",
  "#84cc16", "#e11d48", "#a855f7", "#6b7280",
];

const INIT_TYPES = [
  { id: "t1", name: "Cart Fill", color: "#3b82f6", dur: 60 },
  { id: "t2", name: "TCT", color: "#8b5cf6", dur: 60 },
  { id: "t3", name: "Pyxis Pull", color: "#06b6d4", dur: 60 },
  { id: "t4", name: "Triage", color: "#f59e0b", dur: 60 },
  { id: "t5", name: "IV Room", color: "#10b981", dur: 120 },
  { id: "t6", name: "Chemo", color: "#ef4444", dur: 90 },
  { id: "t7", name: "Pending Queue", color: "#ec4899", dur: 60 },
  { id: "t8", name: "Lunch", color: "#6b7280", dur: 30 },
  { id: "t9", name: "Break", color: "#9ca3af", dur: 15 },
  { id: "t10", name: "Training", color: "#a855f7", dur: 60 },
  { id: "t11", name: "Order Entry", color: "#0ea5e9", dur: 60 },
  { id: "t12", name: "Controlled Substance", color: "#dc2626", dur: 60 },
];

const INIT_ROLES = [
  { id: "r1", name: "Pharmacist 1" },
  { id: "r2", name: "Pharmacist 2" },
  { id: "r3", name: "Tech 1 – Dispensing" },
  { id: "r4", name: "Tech 2 – IV Room" },
  { id: "r5", name: "Tech 3 – Pyxis" },
  { id: "r6", name: "Tech 4 – Triage" },
];

const JUMPS = [0, 3, 6, 9, 12, 15, 18, 21];

const calcLanes = (tasks) => {
  if (!tasks.length) return {};
  const sorted = [...tasks].sort((a, b) => a.start - b.start);
  const lanes = {};
  const active = [];
  for (const t of sorted) {
    const kept = [];
    for (const a of active) { if (a.end > t.start) kept.push(a); }
    active.length = 0;
    active.push(...kept);
    const used = new Set(active.map((a) => a.lane));
    let lane = 0;
    while (used.has(lane)) lane++;
    lanes[t.id] = lane;
    active.push({ end: t.start + t.dur, lane });
  }
  return lanes;
};

/* ═══════════════════════════════════════════════════════════════════════════
   STORAGE HELPERS
   ═══════════════════════════════════════════════════════════════════════════ */
const STORAGE_KEY = "pharmacy-scheduler-data";

const saveToStorage = async (data) => {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch (e) { /* silent */ }
};

const loadFromStorage = async () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) { /* silent */ }
  return null;
};

const clearStorage = async () => {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch (e) { /* silent */ }
};

/* ═══════════════════════════════════════════════════════════════════════════
   EXCEL EXPORT
   ═══════════════════════════════════════════════════════════════════════════ */
function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return { r: parseInt(h.substring(0, 2), 16), g: parseInt(h.substring(2, 4), 16), b: parseInt(h.substring(4, 6), 16) };
}

function exportToExcel(roles, tasks, scenarioName) {
  const SLOTS = 96; // 24 hours * 4 (15-min each)
  const COL_W = 2.3;

  // Build header row: first cell = "Role", then 96 time labels
  const headerRow = ["Role"];
  for (let i = 0; i < SLOTS; i++) {
    headerRow.push(fmtTimeFull(i * 15));
  }

  // Build data rows
  const dataRows = [];
  for (const role of roles) {
    const row = [role.name];
    const roleTasks = tasks.filter((t) => t.roleId === role.id);

    for (let i = 0; i < SLOTS; i++) {
      const slotStart = i * 15;
      const slotEnd = slotStart + 15;
      // Find task that covers this slot
      const task = roleTasks.find((t) => t.start < slotEnd && t.start + t.dur > slotStart);
      row.push(task ? task.name : "");
    }
    dataRows.push({ role, cells: row, roleTasks });
  }

  // Create worksheet
  const wsData = [headerRow, ...dataRows.map((r) => r.cells)];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Column widths
  ws["!cols"] = [{ wch: 20 }]; // Role column
  for (let i = 0; i < SLOTS; i++) {
    ws["!cols"].push({ wch: COL_W });
  }

  // Apply cell styling (colors)
  for (let rIdx = 0; rIdx < dataRows.length; rIdx++) {
    const { roleTasks } = dataRows[rIdx];
    for (let c = 0; c < SLOTS; c++) {
      const slotStart = c * 15;
      const slotEnd = slotStart + 15;
      const task = roleTasks.find((t) => t.start < slotEnd && t.start + t.dur > slotStart);
      if (task) {
        const cellRef = XLSX.utils.encode_cell({ r: rIdx + 1, c: c + 1 });
        if (!ws[cellRef]) ws[cellRef] = { v: task.name, t: "s" };
        const rgb = hexToRgb(task.color);
        ws[cellRef].s = {
          fill: { fgColor: { rgb: task.color.replace("#", "") } },
          font: { color: { rgb: "FFFFFF" }, bold: true, sz: 8 },
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "thin", color: { rgb: "CCCCCC" } },
            bottom: { style: "thin", color: { rgb: "CCCCCC" } },
            left: { style: "thin", color: { rgb: "CCCCCC" } },
            right: { style: "thin", color: { rgb: "CCCCCC" } },
          },
        };
      }
    }
  }

  // Style header row
  for (let c = 0; c <= SLOTS; c++) {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[cellRef]) {
      ws[cellRef].s = {
        fill: { fgColor: { rgb: "1e293b" } },
        font: { color: { rgb: "FFFFFF" }, bold: true, sz: c === 0 ? 10 : 7 },
        alignment: { horizontal: "center", vertical: "center", textRotation: c > 0 ? 90 : 0 },
        border: {
          bottom: { style: "medium", color: { rgb: "000000" } },
        },
      };
    }
  }

  // Style role name column
  for (let r = 1; r <= roles.length; r++) {
    const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
    if (ws[cellRef]) {
      ws[cellRef].s = {
        fill: { fgColor: { rgb: "f1f5f9" } },
        font: { bold: true, sz: 10 },
        alignment: { vertical: "center" },
        border: {
          right: { style: "medium", color: { rgb: "94a3b8" } },
          bottom: { style: "thin", color: { rgb: "e2e8f0" } },
        },
      };
    }
  }

  // Row heights
  ws["!rows"] = [{ hpt: 60 }]; // header taller for rotated text
  for (let i = 0; i < roles.length; i++) {
    ws["!rows"].push({ hpt: 28 });
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, scenarioName || "Schedule");
  XLSX.writeFile(wb, `${(scenarioName || "Schedule").replace(/[^a-zA-Z0-9]/g, "_")}.xlsx`);
}

/* ═══════════════════════════════════════════════════════════════════════════
   MAIN COMPONENT
   ═══════════════════════════════════════════════════════════════════════════ */
export default function PharmacyScheduler() {
  const [types, setTypes] = useState(INIT_TYPES);
  const [roles, setRoles] = useState(INIT_ROLES);
  const [tasks, setTasks] = useState([]);
  const [scale, setScale] = useState(60);
  const [scenarios, setScenarios] = useState([{ id: "s1", name: "Default", tasks: [] }]);
  const [activeSc, setActiveSc] = useState("s1");
  const [editScId, setEditScId] = useState(null);
  const [editScName, setEditScName] = useState("");
  const [loaded, setLoaded] = useState(false);

  const [editTypeId, setEditTypeId] = useState(null);
  const [durPickerId, setDurPickerId] = useState(null);
  const [colorPickerId, setColorPickerId] = useState(null);
  const [addingRole, setAddingRole] = useState(false);
  const [newRoleName, setNewRoleName] = useState("");
  const [editRoleId, setEditRoleId] = useState(null);
  const [editRoleName, setEditRoleName] = useState("");
  const [hoverRoleId, setHoverRoleId] = useState(null);
  const [showResetConfirm, setShowResetConfirm] = useState(false);

  const drag = useRef({ active: false });
  const [ghost, setGhost] = useState(null);
  const scrollRef = useRef(null);
  const boardRef = useRef(null);

  const tasksRef = useRef(tasks);
  tasksRef.current = tasks;
  const rolesRef = useRef(roles);
  rolesRef.current = roles;
  const rowHeightsRef = useRef({});
  const scaleRef = useRef(scale);
  scaleRef.current = scale;
  const activeScRef = useRef(activeSc);
  activeScRef.current = activeSc;

  /* ─── LOAD from persistent storage on mount ──────────────────────── */
  useEffect(() => {
    (async () => {
      const data = await loadFromStorage();
      if (data) {
        if (data.types) setTypes(data.types);
        if (data.roles) setRoles(data.roles);
        if (data.scenarios) {
          setScenarios(data.scenarios);
          const scId = data.activeSc || data.scenarios[0]?.id;
          if (scId) {
            setActiveSc(scId);
            const sc = data.scenarios.find((s) => s.id === scId);
            if (sc) setTasks(sc.tasks);
          }
        }
        if (data.scale) setScale(data.scale);
      }
      setLoaded(true);
    })();
  }, []);

  /* ─── SAVE to persistent storage on every change ─────────────────── */
  useEffect(() => {
    if (!loaded) return;
    saveToStorage({ types, roles, scenarios, activeSc, scale });
  }, [types, roles, scenarios, activeSc, scale, loaded]);

  /* ─── scenario sync ──────────────────────────────────────────────── */
  useEffect(() => {
    const sc = scenarios.find((s) => s.id === activeSc);
    if (sc) setTasks(sc.tasks);
  }, [activeSc, scenarios]);

  const updateTaskInPlace = useCallback((taskId, patch) => {
    setTasks((prev) => prev.map((t) => (t.id === taskId ? { ...t, ...patch } : t)));
    const scId = activeScRef.current;
    setScenarios((prev) =>
      prev.map((s) => s.id === scId ? { ...s, tasks: s.tasks.map((t) => (t.id === taskId ? { ...t, ...patch } : t)) } : s)
    );
  }, []);

  /* ─── row heights ────────────────────────────────────────────────── */
  const rowHeights = useMemo(() => {
    const h = {};
    for (const r of roles) {
      const rt = tasks.filter((t) => t.roleId === r.id);
      const lanes = calcLanes(rt);
      const max = rt.length > 0 ? Math.max(...Object.values(lanes)) + 1 : 1;
      h[r.id] = Math.max(MIN_ROW_H, max * LANE_H + ROW_PAD * 2);
    }
    rowHeightsRef.current = h;
    return h;
  }, [tasks, roles]);

  /* ─── helpers ────────────────────────────────────────────────────── */
  const roleFromY = useCallback((yInBoard) => {
    let cum = 0;
    const rl = rolesRef.current;
    const rh = rowHeightsRef.current;
    for (const r of rl) {
      const h = rh[r.id] || MIN_ROW_H;
      if (yInBoard < cum + h) return r.id;
      cum += h;
    }
    return rl.length > 0 ? rl[rl.length - 1].id : null;
  }, []);

  const getBoardCoords = useCallback((e) => {
    if (!scrollRef.current) return { x: 0, y: 0 };
    const rect = scrollRef.current.getBoundingClientRect();
    return {
      x: e.clientX - rect.left + scrollRef.current.scrollLeft - ROLE_W,
      y: e.clientY - rect.top + scrollRef.current.scrollTop - HEADER_H,
    };
  }, []);

  /* ═══════════════════════════════════════════════════════════════════
     DRAG SYSTEM
     ═══════════════════════════════════════════════════════════════════ */

  const onLibMouseDown = useCallback((e, type) => {
    e.preventDefault();
    drag.current = { active: true, mode: "library", type };
    setGhost({ x: e.clientX, y: e.clientY, name: type.name, color: type.color, dur: type.dur });

    const onMove = (me) => {
      setGhost({ x: me.clientX, y: me.clientY, name: type.name, color: type.color, dur: type.dur });
      const { y } = getBoardCoords(me);
      const rid = roleFromY(y);
      setHoverRoleId(y >= 0 ? rid : null);
    };

    const onUp = (ue) => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      setGhost(null);
      setHoverRoleId(null);
      if (!drag.current.active) return;
      drag.current.active = false;

      const { x, y } = getBoardCoords(ue);
      const rid = roleFromY(y);
      if (y < 0 || !rid) return;

      const sc = scaleRef.current;
      const minute = Math.max(0, Math.min(snapMin(pxToMin(x) - type.dur / 2, sc), TOTAL_MINUTES - type.dur));
      const newTask = { id: uid(), roleId: rid, start: minute, dur: type.dur, name: type.name, color: type.color };
      const newTasks = [...tasksRef.current, newTask];
      setTasks(newTasks);
      const scId = activeScRef.current;
      setScenarios((p) => p.map((s) => (s.id === scId ? { ...s, tasks: newTasks } : s)));
    };

    document.body.style.cursor = "grabbing";
    document.body.style.userSelect = "none";
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }, [getBoardCoords, roleFromY]);

  const onTaskMouseDown = useCallback((e, task) => {
    e.preventDefault();
    e.stopPropagation();
    const { x } = getBoardCoords(e);
    const offsetMin = pxToMin(x) - task.start;
    drag.current = { active: true, mode: "move", taskId: task.id, offsetMin };

    const onMove = (me) => {
      const { x: mx, y: my } = getBoardCoords(me);
      const rid = roleFromY(my);
      const sc = scaleRef.current;
      const dur = tasksRef.current.find((t) => t.id === task.id)?.dur || task.dur;
      const newStart = snapMin(Math.max(0, Math.min(pxToMin(mx) - offsetMin, TOTAL_MINUTES - dur)), sc);
      updateTaskInPlace(task.id, { start: newStart, roleId: rid || task.roleId });
    };

    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      drag.current.active = false;
    };

    document.body.style.cursor = "grabbing";
    document.body.style.userSelect = "none";
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }, [getBoardCoords, roleFromY, updateTaskInPlace]);

  /* RESIZE RIGHT — always snaps to 15 min */
  const onResizeRight = useCallback((e, task) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const origDur = task.dur;
    const taskStart = task.start;

    const onMove = (me) => {
      const dx = me.clientX - startX;
      const newDur = snapMin(Math.max(RESIZE_SNAP, Math.min(origDur + pxToMin(dx), TOTAL_MINUTES - taskStart)), RESIZE_SNAP);
      updateTaskInPlace(task.id, { dur: newDur });
    };

    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    };

    document.body.style.cursor = "ew-resize";
    document.body.style.userSelect = "none";
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }, [updateTaskInPlace]);

  /* RESIZE LEFT — always snaps to 15 min */
  const onResizeLeft = useCallback((e, task) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const origStart = task.start;
    const origEnd = task.start + task.dur;

    const onMove = (me) => {
      const dx = me.clientX - startX;
      let ns = snapMin(origStart + pxToMin(dx), RESIZE_SNAP);
      ns = Math.max(0, Math.min(ns, origEnd - RESIZE_SNAP));
      updateTaskInPlace(task.id, { start: ns, dur: origEnd - ns });
    };

    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    };

    document.body.style.cursor = "ew-resize";
    document.body.style.userSelect = "none";
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }, [updateTaskInPlace]);

  /* ─── task delete ───────────────────────────────────────────────── */
  const deleteTask = useCallback((id) => {
    const newTasks = tasksRef.current.filter((t) => t.id !== id);
    setTasks(newTasks);
    const scId = activeScRef.current;
    setScenarios((p) => p.map((s) => (s.id === scId ? { ...s, tasks: newTasks } : s)));
  }, []);

  /* ─── role CRUD ─────────────────────────────────────────────────── */
  const addRole = () => {
    if (!newRoleName.trim()) return;
    setRoles((p) => [...p, { id: uid(), name: newRoleName.trim() }]);
    setNewRoleName("");
    setAddingRole(false);
  };
  const deleteRole = (id) => {
    setRoles((p) => p.filter((r) => r.id !== id));
    const newTasks = tasksRef.current.filter((t) => t.roleId !== id);
    setTasks(newTasks);
    const scId = activeScRef.current;
    setScenarios((p) => p.map((s) => (s.id === scId ? { ...s, tasks: newTasks } : s)));
  };
  const confirmRenameRole = () => {
    if (editRoleName.trim() && editRoleId) {
      setRoles((p) => p.map((r) => (r.id === editRoleId ? { ...r, name: editRoleName.trim() } : r)));
    }
    setEditRoleId(null);
  };

  /* ─── scenario CRUD ─────────────────────────────────────────────── */
  const addScenario = () => {
    const id = uid();
    setScenarios((p) => [...p, { id, name: `Scenario ${p.length + 1}`, tasks: [] }]);
    setActiveSc(id);
  };
  const dupScenario = () => {
    const cur = scenarios.find((s) => s.id === activeSc);
    if (!cur) return;
    const id = uid();
    setScenarios((p) => [...p, { id, name: `${cur.name} (copy)`, tasks: cur.tasks.map((t) => ({ ...t, id: uid() })) }]);
    setActiveSc(id);
  };
  const delScenario = (id) => {
    if (scenarios.length <= 1) return;
    const rem = scenarios.filter((s) => s.id !== id);
    setScenarios(rem);
    if (activeSc === id) setActiveSc(rem[0].id);
  };

  /* ─── RESET everything ──────────────────────────────────────────── */
  const resetAll = async () => {
    setTypes(INIT_TYPES);
    setRoles(INIT_ROLES);
    setTasks([]);
    setScenarios([{ id: "s1", name: "Default", tasks: [] }]);
    setActiveSc("s1");
    setScale(60);
    setShowResetConfirm(false);
    await clearStorage();
  };

  /* ─── jump to hour ──────────────────────────────────────────────── */
  const jumpTo = (hour) => {
    if (scrollRef.current) scrollRef.current.scrollTo({ left: minToPx(hour * 60), behavior: "smooth" });
  };

  /* ─── type CRUD ─────────────────────────────────────────────────── */
  const updateType = (id, patch) => setTypes((p) => p.map((t) => (t.id === id ? { ...t, ...patch } : t)));
  const deleteType = (id) => setTypes((p) => p.filter((t) => t.id !== id));
  const addType = () => {
    const id = uid();
    setTypes((p) => [...p, { id, name: "New Task", color: COLORS[p.length % COLORS.length], dur: scale }]);
    setEditTypeId(id);
  };

  /* ─── export current scenario ───────────────────────────────────── */
  const handleExport = () => {
    const sc = scenarios.find((s) => s.id === activeSc);
    exportToExcel(roles, tasks, sc?.name || "Schedule");
  };

  /* ─── close popups ──────────────────────────────────────────────── */
  useEffect(() => {
    const handler = (e) => {
      if (!e.target.closest("[data-popup]")) { setDurPickerId(null); setColorPickerId(null); }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  /* ═══════════════════════════════════════════════════════════════════
     RENDER
     ═══════════════════════════════════════════════════════════════════ */

  const ticks = useMemo(() => {
    const arr = [];
    for (let m = 0; m < TOTAL_MINUTES; m += scale) arr.push({ m, isHour: m % 60 === 0 });
    return arr;
  }, [scale]);

  if (!loaded) {
    return (
      <div style={{ height: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans', system-ui, sans-serif", color: "#64748b", background: "#f0f2f5" }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 32, marginBottom: 12 }}>⚡</div>
          <div style={{ fontSize: 14, fontWeight: 500 }}>Loading scheduler...</div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ fontFamily: "'DM Sans', system-ui, -apple-system, sans-serif", height: "100vh", display: "flex", flexDirection: "column", background: "#f0f2f5", color: "#1e293b" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap');
        * { box-sizing: border-box; margin: 0; }
        ::-webkit-scrollbar { height: 10px; width: 10px; }
        ::-webkit-scrollbar-track { background: #e2e8f0; border-radius: 5px; }
        ::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 5px; }
        ::-webkit-scrollbar-thumb:hover { background: #64748b; }
        .task-block:hover .del-btn { opacity: 1 !important; }
        .role-row:hover .role-del { opacity: 1 !important; }
        .type-item:hover .type-actions { opacity: 1 !important; }
        .type-item:hover { background: #f1f5f9 !important; }
        .jump-btn:hover { background: #3b82f6 !important; color: #fff !important; border-color: #3b82f6 !important; }
        .hdr-btn:hover { background: rgba(255,255,255,0.12) !important; }
      `}</style>

      {/* ═══ HEADER ═══════════════════════════════════════════════════ */}
      <div style={{ background: "linear-gradient(135deg, #1e293b 0%, #0f172a 100%)", color: "#fff", padding: "10px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0, boxShadow: "0 2px 12px rgba(0,0,0,0.2)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <img src={LOGO_URL} alt="Logo" style={{ width: 36, height: 36, borderRadius: 10, objectFit: "cover", boxShadow: "0 2px 8px rgba(0,0,0,0.3)" }} />
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, letterSpacing: "-0.02em" }}>Workflow Scheduler</div>
            <div style={{ fontSize: 11, color: "#94a3b8", marginTop: 1 }}>Drag · Resize (15m snap) · Overlap · Full 24hr</div>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {/* Timescale */}
          <div style={{ display: "flex", alignItems: "center", gap: 4, background: "rgba(255,255,255,0.08)", borderRadius: 10, padding: 3, border: "1px solid rgba(255,255,255,0.06)" }}>
            <span style={{ fontSize: 11, color: "#94a3b8", padding: "0 8px", fontWeight: 500 }}>Grid</span>
            {[60, 30, 15].map((s) => (
              <button key={s} onClick={() => setScale(s)}
                style={{ padding: "5px 14px", fontSize: 12, fontWeight: 600, borderRadius: 7, border: "none", cursor: "pointer", transition: "all 0.2s",
                  background: scale === s ? "#3b82f6" : "transparent", color: scale === s ? "#fff" : "#94a3b8",
                  boxShadow: scale === s ? "0 2px 8px rgba(59,130,246,0.3)" : "none" }}>
                {s === 60 ? "1 hr" : `${s}m`}
              </button>
            ))}
          </div>

          {/* Export */}
          <button className="hdr-btn" onClick={handleExport}
            style={{ padding: "6px 14px", fontSize: 12, fontWeight: 600, borderRadius: 8, border: "1px solid rgba(255,255,255,0.12)", background: "rgba(255,255,255,0.06)", color: "#cbd5e1", cursor: "pointer", transition: "all 0.15s", display: "flex", alignItems: "center", gap: 6 }}>
            📥 Export XLSX
          </button>

          {/* Reset */}
          <div style={{ position: "relative" }}>
            <button className="hdr-btn" onClick={() => setShowResetConfirm(!showResetConfirm)}
              style={{ padding: "6px 14px", fontSize: 12, fontWeight: 600, borderRadius: 8, border: "1px solid rgba(239,68,68,0.3)", background: "rgba(239,68,68,0.1)", color: "#fca5a5", cursor: "pointer", transition: "all 0.15s", display: "flex", alignItems: "center", gap: 6 }}>
              ↺ Reset
            </button>
            {showResetConfirm && (
              <div style={{ position: "absolute", top: "calc(100% + 8px)", right: 0, background: "#fff", borderRadius: 12, boxShadow: "0 10px 40px rgba(0,0,0,0.2)", border: "1px solid #e2e8f0", padding: 16, width: 240, zIndex: 999 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: "#1e293b", marginBottom: 6 }}>Reset everything?</div>
                <div style={{ fontSize: 12, color: "#64748b", marginBottom: 14, lineHeight: 1.5 }}>This will clear all tasks, roles, scenarios, and saved data. This cannot be undone.</div>
                <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                  <button onClick={() => setShowResetConfirm(false)}
                    style={{ padding: "6px 14px", fontSize: 12, fontWeight: 500, borderRadius: 7, border: "1px solid #e2e8f0", background: "#fff", color: "#64748b", cursor: "pointer" }}>Cancel</button>
                  <button onClick={resetAll}
                    style={{ padding: "6px 14px", fontSize: 12, fontWeight: 600, borderRadius: 7, border: "none", background: "#ef4444", color: "#fff", cursor: "pointer", boxShadow: "0 2px 6px rgba(239,68,68,0.3)" }}>Reset All</button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* ═══ SCENARIO BAR ═════════════════════════════════════════════ */}
      <div style={{ background: "#1e293b", padding: "5px 20px 7px", display: "flex", alignItems: "center", gap: 8, flexShrink: 0, borderBottom: "1px solid rgba(255,255,255,0.05)" }}>
        <span style={{ fontSize: 10, color: "#475569", fontWeight: 700, marginRight: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>Scenarios</span>
        {scenarios.map((s) => (
          <div key={s.id} onClick={() => setActiveSc(s.id)}
            style={{ display: "flex", alignItems: "center", gap: 6, padding: "4px 12px", borderRadius: 7, fontSize: 12, fontWeight: 500, cursor: "pointer", transition: "all 0.15s",
              background: s.id === activeSc ? "rgba(59,130,246,0.9)" : "rgba(255,255,255,0.06)", color: s.id === activeSc ? "#fff" : "#94a3b8",
              boxShadow: s.id === activeSc ? "0 1px 6px rgba(59,130,246,0.25)" : "none" }}>
            {editScId === s.id ? (
              <input value={editScName} onChange={(e) => setEditScName(e.target.value)} autoFocus
                style={{ background: "transparent", border: "none", borderBottom: "1px solid rgba(255,255,255,0.5)", color: "#fff", fontSize: 12, width: 90, outline: "none", padding: 0 }}
                onBlur={() => { setScenarios((p) => p.map((x) => (x.id === s.id ? { ...x, name: editScName || x.name } : x))); setEditScId(null); }}
                onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
                onClick={(e) => e.stopPropagation()} />
            ) : (
              <span onDoubleClick={(e) => { e.stopPropagation(); setEditScId(s.id); setEditScName(s.name); }}>{s.name}</span>
            )}
            {scenarios.length > 1 && (
              <span onClick={(e) => { e.stopPropagation(); delScenario(s.id); }}
                style={{ cursor: "pointer", opacity: 0.5, fontSize: 15, lineHeight: 1, marginLeft: 2 }}>×</span>
            )}
          </div>
        ))}
        <button onClick={addScenario} style={{ padding: "4px 10px", fontSize: 11, fontWeight: 500, border: "1px dashed rgba(255,255,255,0.15)", borderRadius: 7, background: "transparent", color: "#64748b", cursor: "pointer" }}>+ New</button>
        <button onClick={dupScenario} style={{ padding: "4px 10px", fontSize: 11, fontWeight: 500, border: "1px dashed rgba(255,255,255,0.15)", borderRadius: 7, background: "transparent", color: "#64748b", cursor: "pointer" }}>⧉ Duplicate</button>
      </div>

      {/* ═══ MAIN AREA ════════════════════════════════════════════════ */}
      <div style={{ flex: 1, display: "flex", overflow: "hidden" }}>

        {/* ─── SIDEBAR ───────────────────────────────────────────────── */}
        <div style={{ width: 272, flexShrink: 0, background: "#fff", borderRight: "1px solid #e2e8f0", display: "flex", flexDirection: "column", boxShadow: "2px 0 12px rgba(0,0,0,0.03)" }}>
          <div style={{ padding: "14px 16px 12px", borderBottom: "1px solid #f1f5f9", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.06em" }}>Task Library</span>
            <button onClick={addType} style={{ fontSize: 12, fontWeight: 600, color: "#3b82f6", background: "#eff6ff", border: "none", borderRadius: 7, padding: "5px 12px", cursor: "pointer" }}>+ Add</button>
          </div>

          <div style={{ flex: 1, overflowY: "auto", padding: "8px 8px" }}>
            {types.map((t) => (
              <div key={t.id} className="type-item"
                style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 10px", borderRadius: 9, marginBottom: 3, cursor: "grab", transition: "background 0.12s", position: "relative", background: "#fafbfc" }}
                onMouseDown={(e) => {
                  if (e.target.closest("[data-popup]") || e.target.closest("[data-editable]") || e.target.tagName === "INPUT" || e.target.tagName === "BUTTON") return;
                  onLibMouseDown(e, t);
                }}>
                <div style={{ color: "#cbd5e1", fontSize: 14, lineHeight: 1, flexShrink: 0, cursor: "grab" }}>⠿</div>

                <div data-popup style={{ position: "relative", flexShrink: 0 }}>
                  <div onClick={(e) => { e.stopPropagation(); setColorPickerId(colorPickerId === t.id ? null : t.id); setDurPickerId(null); }}
                    style={{ width: 16, height: 16, borderRadius: 5, background: t.color, cursor: "pointer", border: "2px solid " + t.color + "33" }} />
                  {colorPickerId === t.id && (
                    <div data-popup style={{ position: "absolute", top: 24, left: -4, zIndex: 99, background: "#fff", borderRadius: 12, boxShadow: "0 10px 40px rgba(0,0,0,0.15)", border: "1px solid #e2e8f0", padding: 12, width: 170 }}
                      onClick={(e) => e.stopPropagation()}>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 7 }}>
                        {COLORS.map((c) => (
                          <div key={c} onClick={(e) => { e.stopPropagation(); updateType(t.id, { color: c }); setColorPickerId(null); }}
                            style={{ width: 30, height: 30, borderRadius: 7, background: c, cursor: "pointer", border: c === t.color ? "3px solid #1e293b" : "2px solid transparent", transition: "transform 0.12s" }}
                            onMouseEnter={(e) => (e.currentTarget.style.transform = "scale(1.15)")} onMouseLeave={(e) => (e.currentTarget.style.transform = "scale(1)")} />
                        ))}
                      </div>
                      <div style={{ marginTop: 10, display: "flex", alignItems: "center", gap: 8, borderTop: "1px solid #f1f5f9", paddingTop: 10 }}>
                        <span style={{ fontSize: 10, color: "#94a3b8", fontWeight: 500 }}>Custom</span>
                        <input type="color" value={t.color} onChange={(e) => updateType(t.id, { color: e.target.value })}
                          style={{ width: 32, height: 24, border: "none", padding: 0, cursor: "pointer", background: "transparent" }} />
                      </div>
                    </div>
                  )}
                </div>

                {editTypeId === t.id ? (
                  <input data-editable autoFocus value={t.name} onChange={(e) => updateType(t.id, { name: e.target.value })}
                    onBlur={() => setEditTypeId(null)} onKeyDown={(e) => { if (e.key === "Enter") setEditTypeId(null); }}
                    style={{ flex: 1, fontSize: 13, fontWeight: 500, border: "none", borderBottom: "2px solid #3b82f6", outline: "none", background: "transparent", padding: "1px 0 2px", minWidth: 0, color: "#1e293b" }} />
                ) : (
                  <span onDoubleClick={() => setEditTypeId(t.id)}
                    style={{ flex: 1, fontSize: 13, fontWeight: 500, color: "#334155", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0, lineHeight: "1.3" }}>
                    {t.name}
                  </span>
                )}

                <div data-popup style={{ position: "relative", flexShrink: 0 }}>
                  <button onClick={(e) => { e.stopPropagation(); setDurPickerId(durPickerId === t.id ? null : t.id); setColorPickerId(null); }}
                    style={{ fontSize: 11, fontWeight: 600, color: "#475569", background: "#e8ecf1", border: "none", borderRadius: 20, padding: "3px 10px", cursor: "pointer", whiteSpace: "nowrap" }}>
                    {fmtDur(t.dur)}
                  </button>
                  {durPickerId === t.id && (
                    <div data-popup style={{ position: "absolute", top: 30, right: 0, zIndex: 99, background: "#fff", borderRadius: 12, boxShadow: "0 10px 40px rgba(0,0,0,0.15)", border: "1px solid #e2e8f0", padding: 12, width: 200 }}
                      onClick={(e) => e.stopPropagation()}>
                      <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginBottom: 8, letterSpacing: "0.04em" }}>Duration preset</div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                        {DURATIONS.map((d) => (
                          <button key={d} onClick={(e) => { e.stopPropagation(); updateType(t.id, { dur: d }); setDurPickerId(null); }}
                            style={{ padding: "6px 12px", fontSize: 12, fontWeight: 600, borderRadius: 7, border: "none", cursor: "pointer",
                              background: t.dur === d ? "#3b82f6" : "#f1f5f9", color: t.dur === d ? "#fff" : "#475569",
                              boxShadow: t.dur === d ? "0 1px 4px rgba(59,130,246,0.25)" : "none" }}>
                            {fmtDur(d)}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                <div className="type-actions" style={{ display: "flex", gap: 1, opacity: 0, transition: "opacity 0.15s", flexShrink: 0 }}>
                  <button onClick={(e) => { e.stopPropagation(); setEditTypeId(t.id); }} style={{ background: "none", border: "none", cursor: "pointer", padding: "2px 3px", fontSize: 14, color: "#94a3b8", lineHeight: 1 }}
                    onMouseEnter={(e) => (e.target.style.color = "#3b82f6")} onMouseLeave={(e) => (e.target.style.color = "#94a3b8")}>✎</button>
                  <button onClick={(e) => { e.stopPropagation(); deleteType(t.id); }} style={{ background: "none", border: "none", cursor: "pointer", padding: "2px 3px", fontSize: 16, color: "#94a3b8", lineHeight: 1 }}
                    onMouseEnter={(e) => (e.target.style.color = "#ef4444")} onMouseLeave={(e) => (e.target.style.color = "#94a3b8")}>×</button>
                </div>
              </div>
            ))}
            {types.length === 0 && (
              <div style={{ textAlign: "center", padding: "30px 10px", color: "#94a3b8", fontSize: 13 }}>No tasks yet. Click <strong>+ Add</strong>.</div>
            )}
          </div>

          <div style={{ padding: "10px 16px", borderTop: "1px solid #f1f5f9", fontSize: 11, color: "#b0b8c4" }}>
            Drag onto board · Double-click to rename
          </div>
        </div>

        {/* ─── BOARD ─────────────────────────────────────────────────── */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>

          <div style={{ display: "flex", alignItems: "center", gap: 5, padding: "6px 16px", background: "#fff", borderBottom: "1px solid #e2e8f0", flexShrink: 0 }}>
            <span style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", marginRight: 6, letterSpacing: "0.05em" }}>Jump</span>
            {JUMPS.map((h) => (
              <button key={h} className="jump-btn" onClick={() => jumpTo(h)}
                style={{ padding: "4px 10px", fontSize: 11, fontWeight: 600, borderRadius: 6, border: "1px solid #e2e8f0", background: "#fff", color: "#475569", cursor: "pointer", transition: "all 0.15s" }}>
                {fmtTime(h * 60)}
              </button>
            ))}
          </div>

          <div ref={scrollRef} style={{ flex: 1, overflow: "auto", background: "#f8fafc" }}>
            <div style={{ width: TOTAL_W + ROLE_W + 40, minHeight: "100%" }}>

              <div style={{ position: "sticky", top: 0, zIndex: 25, display: "flex", height: HEADER_H, background: "#fff", borderBottom: "2px solid #e2e8f0" }}>
                <div style={{ width: ROLE_W, flexShrink: 0, position: "sticky", left: 0, zIndex: 30, background: "#fff", borderRight: "1px solid #e2e8f0", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <span style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: "0.06em" }}>Roles</span>
                </div>
                <div style={{ position: "relative", width: TOTAL_W, height: HEADER_H }}>
                  {ticks.map(({ m, isHour }) => (
                    <div key={m} style={{ position: "absolute", left: minToPx(m), bottom: 0, display: "flex", flexDirection: "column", alignItems: "flex-start" }}>
                      <span style={{ fontSize: isHour ? 12 : 10, fontWeight: isHour ? 600 : 400, color: isHour ? "#475569" : "#b0b8c4", paddingLeft: 4, paddingBottom: 5, whiteSpace: "nowrap", userSelect: "none" }}>
                        {fmtTime(m)}
                      </span>
                      <div style={{ width: 1, height: isHour ? 8 : 4, background: isHour ? "#94a3b8" : "#cbd5e1" }} />
                    </div>
                  ))}
                </div>
              </div>

              <div ref={boardRef}>
                {roles.map((role) => {
                  const rowH = rowHeights[role.id];
                  const rowTasks = tasks.filter((t) => t.roleId === role.id);
                  const lanes = calcLanes(rowTasks);

                  return (
                    <div key={role.id} className="role-row"
                      style={{ display: "flex", height: rowH, minHeight: MIN_ROW_H, borderBottom: "1px solid #e2e8f0", transition: "background 0.15s",
                        background: hoverRoleId === role.id ? "#eff6ff" : "transparent" }}>

                      <div style={{ width: ROLE_W, flexShrink: 0, position: "sticky", left: 0, zIndex: 15, background: hoverRoleId === role.id ? "#eff6ff" : "#fff", borderRight: "1px solid #e2e8f0", display: "flex", alignItems: "center", padding: "0 10px 0 14px", gap: 6, transition: "background 0.15s" }}>
                        {editRoleId === role.id ? (
                          <input autoFocus value={editRoleName} onChange={(e) => setEditRoleName(e.target.value)}
                            onBlur={confirmRenameRole} onKeyDown={(e) => { if (e.key === "Enter") confirmRenameRole(); if (e.key === "Escape") setEditRoleId(null); }}
                            style={{ flex: 1, fontSize: 13, fontWeight: 500, border: "none", borderBottom: "2px solid #3b82f6", outline: "none", background: "transparent", padding: "2px 0", minWidth: 0, color: "#1e293b" }} />
                        ) : (
                          <span onDoubleClick={() => { setEditRoleId(role.id); setEditRoleName(role.name); }}
                            style={{ flex: 1, fontSize: 13, fontWeight: 500, color: "#334155", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", cursor: "default", lineHeight: "1.3" }}>
                            {role.name}
                          </span>
                        )}
                        <button className="role-del" onClick={() => deleteRole(role.id)}
                          style={{ opacity: 0, transition: "opacity 0.15s, color 0.12s", background: "none", border: "none", cursor: "pointer", fontSize: 18, color: "#b0b8c4", lineHeight: 1, padding: "0 2px", flexShrink: 0 }}
                          onMouseEnter={(e) => (e.currentTarget.style.color = "#ef4444")} onMouseLeave={(e) => (e.currentTarget.style.color = "#b0b8c4")}>×</button>
                      </div>

                      <div style={{ position: "relative", width: TOTAL_W, height: rowH }}>
                        {ticks.map(({ m, isHour }) => (
                          <div key={m} style={{ position: "absolute", left: minToPx(m), top: 0, bottom: 0, width: 1, background: isHour ? "#e2e8f0" : "#f1f5f9", pointerEvents: "none" }} />
                        ))}

                        {rowTasks.map((task) => {
                          const lane = lanes[task.id] || 0;
                          const left = minToPx(task.start);
                          const w = minToPx(task.dur);
                          const top = ROW_PAD + lane * LANE_H;

                          return (
                            <div key={task.id} className="task-block"
                              onMouseDown={(e) => {
                                if (e.target.closest("[data-resize]")) return;
                                onTaskMouseDown(e, task);
                              }}
                              style={{ position: "absolute", left, width: Math.max(w, 16), top, height: LANE_H - 6, borderRadius: 7, background: task.color, cursor: "grab", display: "flex", alignItems: "center", overflow: "visible", zIndex: 10, transition: "box-shadow 0.12s",
                                boxShadow: `0 1px 4px ${task.color}44, 0 0 0 1px ${task.color}22` }}
                              onMouseEnter={(e) => { e.currentTarget.style.boxShadow = `0 3px 12px ${task.color}55, 0 0 0 2px ${task.color}44`; e.currentTarget.style.zIndex = 20; }}
                              onMouseLeave={(e) => { e.currentTarget.style.boxShadow = `0 1px 4px ${task.color}44, 0 0 0 1px ${task.color}22`; e.currentTarget.style.zIndex = 10; }}>

                              <div data-resize onMouseDown={(e) => onResizeLeft(e, task)}
                                style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: 8, cursor: "ew-resize", borderRadius: "7px 0 0 7px", zIndex: 5 }}
                                onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(255,255,255,0.25)")} onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")} />

                              <div style={{ flex: 1, padding: "0 12px", overflow: "hidden", pointerEvents: "none", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 4 }}>
                                <span style={{ fontSize: 12, fontWeight: 600, color: "#fff", textOverflow: "ellipsis", overflow: "hidden", whiteSpace: "nowrap", textShadow: "0 1px 3px rgba(0,0,0,0.2)" }}>
                                  {task.name}
                                </span>
                                {w > 55 && (
                                  <span style={{ fontSize: 10, fontWeight: 500, color: "rgba(255,255,255,0.75)", flexShrink: 0, textShadow: "0 1px 2px rgba(0,0,0,0.15)" }}>
                                    {fmtDur(task.dur)}
                                  </span>
                                )}
                              </div>

                              {/* DELETE — positioned at top:0 right:0 inside visible overflow */}
                              <button className="del-btn"
                                onClick={(e) => { e.stopPropagation(); e.preventDefault(); deleteTask(task.id); }}
                                onMouseDown={(e) => e.stopPropagation()}
                                style={{ opacity: 0, transition: "opacity 0.12s", position: "absolute", top: 0, right: 0, width: 18, height: 18, borderRadius: "0 7px 0 6px", background: "rgba(0,0,0,0.35)", border: "none", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", fontSize: 12, color: "rgba(255,255,255,0.85)", lineHeight: 1, zIndex: 30, padding: 0 }}
                                onMouseEnter={(e) => { e.currentTarget.style.background = "#ef4444"; e.currentTarget.style.color = "#fff"; }}
                                onMouseLeave={(e) => { e.currentTarget.style.background = "rgba(0,0,0,0.35)"; e.currentTarget.style.color = "rgba(255,255,255,0.85)"; }}>
                                ×
                              </button>

                              <div data-resize onMouseDown={(e) => onResizeRight(e, task)}
                                style={{ position: "absolute", right: 0, top: 0, bottom: 0, width: 8, cursor: "ew-resize", borderRadius: "0 7px 7px 0", zIndex: 5 }}
                                onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(255,255,255,0.25)")} onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")} />
                            </div>
                          );
                        })}

                        {rowTasks.length === 0 && (
                          <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
                            <span style={{ fontSize: 11, color: "#d0d5dd", fontStyle: "italic" }}>Drop tasks here</span>
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}

                <div style={{ display: "flex", height: MIN_ROW_H, borderBottom: "1px solid #e2e8f0" }}>
                  <div style={{ width: ROLE_W, flexShrink: 0, position: "sticky", left: 0, zIndex: 15, background: "#fff", borderRight: "1px solid #e2e8f0", display: "flex", alignItems: "center", padding: "0 14px" }}>
                    {addingRole ? (
                      <input autoFocus value={newRoleName} onChange={(e) => setNewRoleName(e.target.value)} placeholder="Role name..."
                        onKeyDown={(e) => { if (e.key === "Enter") addRole(); if (e.key === "Escape") { setAddingRole(false); setNewRoleName(""); } }}
                        onBlur={() => { if (newRoleName.trim()) addRole(); else { setAddingRole(false); setNewRoleName(""); } }}
                        style={{ flex: 1, fontSize: 13, fontWeight: 500, border: "none", borderBottom: "2px solid #3b82f6", outline: "none", background: "transparent", padding: "2px 0", minWidth: 0, color: "#1e293b" }} />
                    ) : (
                      <button onClick={() => setAddingRole(true)}
                        style={{ fontSize: 12, fontWeight: 500, color: "#94a3b8", background: "none", border: "1.5px dashed #d1d5db", borderRadius: 7, padding: "6px 0", cursor: "pointer", transition: "all 0.15s", width: "100%" }}
                        onMouseEnter={(e) => { e.currentTarget.style.color = "#3b82f6"; e.currentTarget.style.borderColor = "#93c5fd"; e.currentTarget.style.background = "#eff6ff"; }}
                        onMouseLeave={(e) => { e.currentTarget.style.color = "#94a3b8"; e.currentTarget.style.borderColor = "#d1d5db"; e.currentTarget.style.background = "none"; }}>
                        + Add Role
                      </button>
                    )}
                  </div>
                  <div style={{ flex: 1 }} />
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* ═══ STATUS BAR ═══════════════════════════════════════════════ */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "5px 20px", background: "linear-gradient(135deg, #1e293b, #0f172a)", color: "#64748b", fontSize: 11, flexShrink: 0, fontWeight: 500 }}>
        <div style={{ display: "flex", gap: 18 }}>
          <span>{tasks.length} task{tasks.length !== 1 ? "s" : ""}</span>
          <span>{roles.length} role{roles.length !== 1 ? "s" : ""}</span>
          <span>{types.length} types</span>
        </div>
        <div style={{ display: "flex", gap: 18 }}>
          <span>Grid: {scale === 60 ? "1 hour" : `${scale} min`}</span>
          <span>Resize: 15m snap</span>
          <span>Overlap: on</span>
          <span>💾 Auto-saved</span>
        </div>
      </div>

      {/* ═══ DRAG GHOST ═══════════════════════════════════════════════ */}
      {ghost && (
        <div style={{ position: "fixed", left: ghost.x - 50, top: ghost.y - 18, pointerEvents: "none", zIndex: 9999, background: ghost.color, color: "#fff", padding: "7px 16px", borderRadius: 9, fontSize: 13, fontWeight: 600, boxShadow: `0 10px 30px ${ghost.color}55, 0 4px 12px rgba(0,0,0,0.15)`, opacity: 0.95, whiteSpace: "nowrap", textShadow: "0 1px 3px rgba(0,0,0,0.2)" }}>
          {ghost.name} · {fmtDur(ghost.dur)}
        </div>
      )}
    </div>
  );
}
