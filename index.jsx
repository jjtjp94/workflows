import React, { useState, useRef, useCallback, useEffect, useMemo } from "react";
import { 
  Plus, Trash2, GripVertical, X, Settings, 
  ChevronRight, ChevronLeft, Clock, Search, 
  ZoomIn, ZoomOut, Calendar, User, Briefcase 
} from "lucide-react";

// ─── Constants & Config ─────────────────────────────────────────────────────
const MINUTES_IN_DAY = 1440; // 24 hours
const SNAP_MINUTES = 15;     // Snap to grid (15 min)
const MIN_DURATION = 15;     // Minimum task size

const PRESET_DURATIONS = [
  { label: '15m', val: 15 },
  { label: '30m', val: 30 },
  { label: '45m', val: 45 },
  { label: '1h', val: 60 },
  { label: '1.5h', val: 90 },
  { label: '2h', val: 120 },
  { label: '4h', val: 240 },
  { label: '8h', val: 480 },
];

const PRESET_COLORS = [
  { bg: "#3b82f6", border: "#2563eb", text: "#fff" }, // Blue
  { bg: "#8b5cf6", border: "#7c3aed", text: "#fff" }, // Purple
  { bg: "#10b981", border: "#059669", text: "#fff" }, // Emerald
  { bg: "#f59e0b", border: "#d97706", text: "#fff" }, // Amber
  { bg: "#ef4444", border: "#dc2626", text: "#fff" }, // Red
  { bg: "#ec4899", border: "#db2777", text: "#fff" }, // Pink
  { bg: "#6b7280", border: "#4b5563", text: "#fff" }, // Gray
  { bg: "#14b8a6", border: "#0d9488", text: "#fff" }, // Teal
];

// Initial Data
const INITIAL_ROLES = [
  { id: "r1", name: "RPh: Verification" },
  { id: "r2", name: "RPh: Clinical" },
  { id: "r3", name: "Tech: IV Room" },
  { id: "r4", name: "Tech: Triage" },
  { id: "r5", name: "Tech: Delivery" },
];

const INITIAL_TASK_TYPES = [
  { id: "t1", name: "Morning Cart Fill", colorIdx: 0, defaultDuration: 60 },
  { id: "t2", name: "IV Batch", colorIdx: 2, defaultDuration: 120 },
  { id: "t3", name: "Stat Order", colorIdx: 4, defaultDuration: 30 },
  { id: "t4", name: "Lunch", colorIdx: 6, defaultDuration: 30 },
  { id: "t5", name: "Meeting", colorIdx: 1, defaultDuration: 60 },
];

// ─── Helpers ────────────────────────────────────────────────────────────────
const generateId = () => `id_${Math.random().toString(36).substr(2, 9)}`;

const minutesToTime = (m) => {
  const h = Math.floor(m / 60);
  const mm = Math.floor(m % 60);
  const ampm = h >= 12 && h < 24 ? "PM" : "AM";
  const h12 = h === 0 || h === 24 ? 12 : h > 12 ? h - 12 : h;
  return `${h12}:${String(mm).padStart(2, "0")} ${ampm}`;
};

// Logic to stack overlapping tasks into "lanes"
const calculateLanes = (tasks) => {
  if (!tasks.length) return { lanes: {}, maxLane: 0 };
  
  // Sort by start time
  const sorted = [...tasks].sort((a, b) => a.start - b.start);
  const lanes = {};
  const laneEnds = []; // array of end times for each lane index

  for (const task of sorted) {
    let assignedLane = -1;
    
    // Try to find a lane where this task fits
    for (let i = 0; i < laneEnds.length; i++) {
      if (task.start >= laneEnds[i]) {
        assignedLane = i;
        laneEnds[i] = task.start + task.duration;
        break;
      }
    }

    // If no lane fits, create a new one
    if (assignedLane === -1) {
      laneEnds.push(task.start + task.duration);
      assignedLane = laneEnds.length - 1;
    }

    lanes[task.id] = assignedLane;
  }

  return { lanes, maxLane: laneEnds.length };
};

// ─── Components ─────────────────────────────────────────────────────────────

const TaskLibraryItem = ({ type, onUpdate, onDelete, onDragStart }) => {
  const color = PRESET_COLORS[type.colorIdx];
  
  return (
    <div 
      className="group flex flex-col p-3 bg-white border border-gray-200 rounded-lg shadow-sm hover:shadow-md transition-all cursor-grab active:cursor-grabbing mb-2"
      draggable
      onDragStart={(e) => onDragStart(e, type)}
    >
      <div className="flex items-center justify-between mb-2">
        <div className="flex items-center gap-2 overflow-hidden">
          <div className="w-3 h-3 rounded-full flex-shrink-0" style={{ backgroundColor: color.bg }}></div>
          <input 
            className="text-sm font-medium text-gray-700 bg-transparent border-none focus:ring-0 p-0 w-full truncate"
            value={type.name}
            onChange={(e) => onUpdate({...type, name: e.target.value})}
          />
        </div>
        <button onClick={() => onDelete(type.id)} className="text-gray-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity">
          <X size={14} />
        </button>
      </div>

      <div className="flex items-center justify-between gap-2">
        <div className="flex gap-1">
          {PRESET_COLORS.map((c, idx) => (
            idx < 5 && ( // Show first 5 colors or valid ones
              <button 
                key={idx}
                className={`w-4 h-4 rounded-full transition-transform hover:scale-110 ${type.colorIdx === idx ? 'ring-2 ring-offset-1 ring-gray-400' : ''}`}
                style={{ backgroundColor: c.bg }}
                onClick={() => onUpdate({...type, colorIdx: idx})}
              />
            )
          ))}
        </div>
        
        <select 
          className="text-xs border-gray-200 bg-gray-50 rounded py-1 px-1 text-gray-600 focus:ring-blue-500 focus:border-blue-500"
          value={type.defaultDuration}
          onChange={(e) => onUpdate({...type, defaultDuration: parseInt(e.target.value)})}
        >
          {PRESET_DURATIONS.map(d => (
            <option key={d.label} value={d.val}>{d.label}</option>
          ))}
        </select>
      </div>
    </div>
  );
};

const TaskBlock = ({ task, style, onDragStart, onResizeStart, onDelete }) => {
  const color = PRESET_COLORS[task.colorIdx] || PRESET_COLORS[0];
  
  return (
    <div
      className="absolute rounded shadow-sm group select-none overflow-hidden flex flex-col justify-center px-2"
      style={{
        ...style,
        backgroundColor: color.bg,
        borderLeft: `3px solid ${color.border}`,
        opacity: style.isDragging ? 0.5 : 1,
        zIndex: style.isDragging ? 50 : 10,
        cursor: 'grab'
      }}
      draggable
      onDragStart={(e) => onDragStart(e, task)}
    >
      <div className="flex justify-between items-center pointer-events-none">
        <span className="text-xs font-semibold truncate drop-shadow-sm" style={{ color: color.text }}>
          {task.name}
        </span>
        {/* Only show delete on hover */}
        <button 
          className="pointer-events-auto opacity-0 group-hover:opacity-100 text-white hover:text-red-100"
          onClick={(e) => { e.stopPropagation(); onDelete(task.id); }}
        >
          <X size={12} />
        </button>
      </div>
      
      {task.duration >= 45 && (
        <span className="text-[10px] opacity-90 truncate pointer-events-none" style={{ color: color.text }}>
          {minutesToTime(task.start)} - {minutesToTime(task.start + task.duration)}
        </span>
      )}

      {/* Resize Handle */}
      <div 
        className="absolute top-0 bottom-0 right-0 w-3 cursor-ew-resize hover:bg-black/10 z-20 flex items-center justify-center"
        onMouseDown={(e) => onResizeStart(e, task)}
        onClick={(e) => e.stopPropagation()} 
      >
        <div className="w-0.5 h-3 bg-white/30 rounded-full" />
      </div>
    </div>
  );
};

// ─── Main Application ───────────────────────────────────────────────────────

export default function PharmacyScheduler() {
  const [tasks, setTasks] = useState([]);
  const [roles, setRoles] = useState(INITIAL_ROLES);
  const [taskTypes, setTaskTypes] = useState(INITIAL_TASK_TYPES);
  
  // View State
  const [zoom, setZoom] = useState(2.5); // Pixels per minute
  const [currentTime, setCurrentTime] = useState(new Date());

  // Refs
  const scrollContainerRef = useRef(null);
  const timelineRef = useRef(null);

  // ─── Calculated Values ────────────────────────────────────────────────────
  const hourWidth = 60 * zoom;
  const totalWidth = MINUTES_IN_DAY * zoom;

  // Group tasks by role for rendering
  const tasksByRole = useMemo(() => {
    const grouped = {};
    roles.forEach(r => grouped[r.id] = []);
    tasks.forEach(t => {
      if (grouped[t.roleId]) grouped[t.roleId].push(t);
    });
    return grouped;
  }, [tasks, roles]);

  // Update current time line
  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 60000);
    return () => clearInterval(timer);
  }, []);

  const currentMinutes = currentTime.getHours() * 60 + currentTime.getMinutes();

  // ─── Event Handlers ───────────────────────────────────────────────────────

  const handleDragStartNew = (e, type) => {
    e.dataTransfer.setData("application/type", "NEW");
    e.dataTransfer.setData("application/json", JSON.stringify(type));
    e.dataTransfer.effectAllowed = "copy";
  };

  const handleDragStartMove = (e, task) => {
    e.dataTransfer.setData("application/type", "MOVE");
    e.dataTransfer.setData("application/id", task.id);
    
    // Calculate offset from the start of the task block
    const rect = e.target.getBoundingClientRect();
    const offsetX = e.clientX - rect.left;
    e.dataTransfer.setData("application/offset", offsetX);
    e.dataTransfer.effectAllowed = "move";
  };

  const handleDrop = (e, roleId) => {
    e.preventDefault();
    const type = e.dataTransfer.getData("application/type");
    const rect = e.currentTarget.getBoundingClientRect();
    
    // Calculate time based on drop position + scroll
    // Note: e.currentTarget is the Row. e.clientX is global.
    // We need the X relative to the SCROLL CONTAINER content.
    
    // Let's rely on the relative click within the row:
    const relativeX = e.clientX - rect.left + scrollContainerRef.current.scrollLeft;
    // However, the Row component might be scrolled out. 
    // Better calculation:
    // The relative X inside the row container = (e.clientX - rect.left).
    // But since the row container scrolls, this DOM calculation is tricky if the row handles overflow.
    // Actually, the 'row' div is inside the 'scrollContainer'.
    // The 'row' div is wide (24h * zoom). It doesn't scroll itself, the parent does.
    // So e.clientX - rect.left IS the pixel position in the timeline IF the row has relative positioning.
    
    const clickX = e.clientX - rect.left;
    
    if (type === "NEW") {
      const taskType = JSON.parse(e.dataTransfer.getData("application/json"));
      let start = Math.floor(clickX / zoom);
      
      // Snap
      start = Math.round(start / SNAP_MINUTES) * SNAP_MINUTES;
      
      // Bounds check
      if (start < 0) start = 0;
      if (start + taskType.defaultDuration > MINUTES_IN_DAY) start = MINUTES_IN_DAY - taskType.defaultDuration;

      const newTask = {
        id: generateId(),
        roleId,
        name: taskType.name,
        colorIdx: taskType.colorIdx,
        start,
        duration: taskType.defaultDuration,
      };
      setTasks([...tasks, newTask]);

    } else if (type === "MOVE") {
      const taskId = e.dataTransfer.getData("application/id");
      const offsetPx = parseFloat(e.dataTransfer.getData("application/offset"));
      
      // Adjust start time by the offset where the user grabbed the block
      let start = Math.floor((clickX - offsetPx) / zoom);
      
      // Snap
      start = Math.round(start / SNAP_MINUTES) * SNAP_MINUTES;

      // Find task to ensure duration bounds
      const task = tasks.find(t => t.id === taskId);
      if (!task) return;

      if (start < 0) start = 0;
      if (start + task.duration > MINUTES_IN_DAY) start = MINUTES_IN_DAY - task.duration;

      setTasks(tasks.map(t => t.id === taskId ? { ...t, start, roleId } : t));
    }
  };

  const handleResizeStart = useCallback((e, task) => {
    e.preventDefault();
    e.stopPropagation();
    
    const startX = e.clientX;
    const startDuration = task.duration;

    const onMouseMove = (moveEvent) => {
      const deltaPx = moveEvent.clientX - startX;
      const deltaMin = deltaPx / zoom;
      
      let newDuration = startDuration + deltaMin;
      
      // Snap
      newDuration = Math.round(newDuration / SNAP_MINUTES) * SNAP_MINUTES;
      if (newDuration < MIN_DURATION) newDuration = MIN_DURATION;
      
      // Max bound
      if (task.start + newDuration > MINUTES_IN_DAY) newDuration = MINUTES_IN_DAY - task.start;

      setTasks(prev => prev.map(t => t.id === task.id ? { ...t, duration: newDuration } : t));
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }, [zoom]);

  // ─── Render Helpers ───────────────────────────────────────────────────────
  
  const scrollToNow = () => {
    if (scrollContainerRef.current) {
      const nowPx = currentMinutes * zoom;
      scrollContainerRef.current.scrollTo({ left: nowPx - 300, behavior: 'smooth' });
    }
  };

  const addRole = () => {
    const name = prompt("Enter Role Name:");
    if (name) setRoles([...roles, { id: generateId(), name }]);
  };

  const deleteRole = (id) => {
    if (confirm("Delete this role? Tasks in this lane will be removed.")) {
      setRoles(roles.filter(r => r.id !== id));
      setTasks(tasks.filter(t => t.roleId !== id));
    }
  };

  const addTaskType = () => {
    setTaskTypes([...taskTypes, { 
      id: generateId(), 
      name: "New Task", 
      colorIdx: Math.floor(Math.random() * 5), 
      defaultDuration: 60 
    }]);
  };

  // ─── JSX ──────────────────────────────────────────────────────────────────

  return (
    <div className="flex flex-col h-screen bg-gray-50 text-gray-800 font-sans">
      
      {/* HEADER */}
      <header className="flex items-center justify-between px-6 py-3 bg-white border-b border-gray-200 shadow-sm z-20">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg text-white">
            <Calendar size={20} />
          </div>
          <div>
            <h1 className="text-lg font-bold text-gray-900 leading-tight">Pharmacy Scheduler</h1>
            <p className="text-xs text-gray-500">
              {currentTime.toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric' })}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-4">
          {/* Zoom Control */}
          <div className="flex items-center gap-2 bg-gray-100 rounded-lg p-1.5 px-3">
            <ZoomOut size={16} className="text-gray-400" />
            <input 
              type="range" min="1" max="6" step="0.5" 
              value={zoom} onChange={(e) => setZoom(parseFloat(e.target.value))}
              className="w-24 h-1 bg-gray-300 rounded-lg appearance-none cursor-pointer accent-blue-600"
            />
            <ZoomIn size={16} className="text-gray-400" />
          </div>

          <button onClick={scrollToNow} className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-blue-600 bg-blue-50 rounded-lg hover:bg-blue-100 transition-colors">
            <Clock size={16} /> Jump to Now
          </button>
        </div>
      </header>

      <div className="flex flex-1 overflow-hidden">
        
        {/* SIDEBAR: LIBRARY */}
        <aside className="w-80 bg-white border-r border-gray-200 flex flex-col z-20 shadow-[4px_0_24px_rgba(0,0,0,0.02)]">
          <div className="p-4 border-b border-gray-100 flex items-center justify-between">
            <h2 className="font-semibold text-sm text-gray-700 uppercase tracking-wide flex items-center gap-2">
              <Briefcase size={14} /> Task Library
            </h2>
            <button onClick={addTaskType} className="p-1 hover:bg-gray-100 rounded text-blue-600">
              <Plus size={18} />
            </button>
          </div>
          
          <div className="flex-1 overflow-y-auto p-4">
            {taskTypes.map(type => (
              <TaskLibraryItem 
                key={type.id} 
                type={type} 
                onUpdate={(updated) => setTaskTypes(taskTypes.map(t => t.id === updated.id ? updated : t))}
                onDelete={(id) => setTaskTypes(taskTypes.filter(t => t.id !== id))}
                onDragStart={handleDragStartNew}
              />
            ))}
            <div className="mt-4 p-4 border-2 border-dashed border-gray-100 rounded-lg text-center">
              <p className="text-xs text-gray-400">Drag tasks from here onto the timeline.</p>
            </div>
          </div>
        </aside>

        {/* MAIN BOARD */}
        <div className="flex-1 flex flex-col relative bg-white">
          
          {/* SCROLLABLE AREA */}
          <div 
            ref={scrollContainerRef}
            className="flex-1 overflow-auto relative custom-scrollbar"
            style={{ scrollBehavior: 'smooth' }}
          >
            <div style={{ width: totalWidth, minWidth: '100%', position: 'relative' }}>
              
              {/* TIMELINE HEADER (Sticky) */}
              <div className="sticky top-0 z-30 bg-gray-50 border-b border-gray-200 h-10 flex text-xs text-gray-500 font-medium select-none shadow-sm">
                 {/* Empty corner for role column */}
                 <div className="sticky left-0 w-48 bg-gray-50 border-r border-gray-200 z-40 flex-shrink-0 flex items-center justify-between px-3">
                    <span>Roles</span>
                    <button onClick={addRole} className="p-1 hover:bg-gray-200 rounded text-gray-600" title="Add Role">
                      <Plus size={14} />
                    </button>
                 </div>

                 {/* Hours */}
                 <div className="relative flex-1">
                   {Array.from({ length: 24 }).map((_, h) => (
                     <div key={h} className="absolute top-0 bottom-0 border-l border-gray-300 pl-1 pt-2"
                        style={{ left: h * 60 * zoom }}>
                        {h === 0 ? "12 AM" : h === 12 ? "12 PM" : h > 12 ? `${h-12} PM` : `${h} AM`}
                     </div>
                   ))}
                   {/* Minor Grid Lines (every 30m if zoomed in) */}
                    {zoom > 2 && Array.from({ length: 48 }).map((_, h) => (
                      h % 2 !== 0 && (
                        <div key={h} className="absolute top-6 bottom-0 border-l border-gray-200 h-2"
                          style={{ left: h * 30 * zoom }} />
                      )
                   ))}
                 </div>
              </div>

              {/* CURRENT TIME INDICATOR LINE */}
              <div 
                className="absolute top-0 bottom-0 border-l-2 border-red-500 z-10 pointer-events-none"
                style={{ left: currentMinutes * zoom }}
              >
                <div className="bg-red-500 text-white text-[10px] px-1 rounded absolute -left-[18px] top-8 font-bold">
                  {minutesToTime(currentMinutes)}
                </div>
              </div>

              {/* ROLES / LANES */}
              {roles.map(role => {
                const roleTasks = tasksByRole[role.id];
                const { lanes, maxLane } = calculateLanes(roleTasks);
                const height = Math.max(60, (maxLane + 1) * 45); // Dynamic height based on overlapping tasks

                return (
                  <div 
                    key={role.id} 
                    className="flex border-b border-gray-100 relative group"
                    style={{ height }}
                    onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = "move"; }}
                    onDrop={(e) => handleDrop(e, role.id)}
                  >
                    {/* Sticky Role Label */}
                    <div className="sticky left-0 w-48 bg-white border-r border-gray-200 z-20 flex-shrink-0 flex items-center justify-between px-3 group-hover:bg-gray-50 transition-colors">
                      <div className="flex items-center gap-2 overflow-hidden">
                        <User size={14} className="text-gray-400" />
                        <span className="text-sm font-medium text-gray-700 truncate">{role.name}</span>
                      </div>
                      <button onClick={() => deleteRole(role.id)} className="opacity-0 group-hover:opacity-100 text-gray-400 hover:text-red-500 p-1">
                        <Trash2 size={14} />
                      </button>
                    </div>

                    {/* Timeline Grid Background */}
                    <div className="absolute inset-0 z-0 pointer-events-none">
                      {Array.from({ length: 24 }).map((_, h) => (
                        <div key={h} className="absolute top-0 bottom-0 border-l border-dashed border-gray-100"
                           style={{ left: h * 60 * zoom }} />
                      ))}
                    </div>

                    {/* Tasks */}
                    <div className="relative flex-1">
                      {roleTasks.map(task => {
                        const laneIndex = lanes[task.id];
                        const top = laneIndex * 40 + 5; // 40px height per lane, 5px padding
                        const left = task.start * zoom;
                        const width = task.duration * zoom;

                        return (
                          <TaskBlock 
                            key={task.id}
                            task={task}
                            style={{ top, left, width, height: 36 }} // Fixed height for blocks
                            onDragStart={handleDragStartMove}
                            onResizeStart={handleResizeStart}
                            onDelete={(id) => setTasks(tasks.filter(t => t.id !== id))}
                          />
                        );
                      })}
                    </div>
                  </div>
                );
              })}

              {/* Bottom padding for scroll */}
              <div className="h-12"></div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
