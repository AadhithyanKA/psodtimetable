import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Timetable Maker — Weekly course & faculty planner" },
      {
        name: "description",
        content:
          "Build modular week-by-week timetables with course allocation, faculty conflict detection, breaks and blocked slots. Export as Excel or CSV.",
      },
      { property: "og:title", content: "Timetable Maker" },
      {
        property: "og:description",
        content:
          "Plan any date range. Days as rows, time as columns. Click and drag to allocate courses across multiple classes.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

type Cell =
  | { kind: "empty" }
  | { kind: "break"; label: string }
  | { kind: "blocked"; label: string }
  | { kind: "course"; courseId: string };

type Course = {
  id: string;
  name: string;
  faculty: string;
  color: string;
  durationSlots: number;
  // Target number of sessions per week (used by auto-fill). 0 = don't auto-fill.
  weeklyPeriods?: number;
  // 0=Sun..6=Sat. undefined or empty = allowed on all days.
  allowedWeekdays?: number[];
  // Slot indices. undefined or empty = allowed in all periods.
  allowedSlots?: number[];
};
type ClassData = { id: string; name: string; grid: Record<string, Cell>; courses: Course[] };
type Slot = { start: string; end: string; isBreak?: boolean }; // 24h "HH:MM"
type State = {
  fromDate: string; // YYYY-MM-DD
  toDate: string;
  slots: Slot[];
  classes: ClassData[];
};

const COLORS = [
  "#fdba74", "#fcd34d", "#86efac", "#67e8f9",
  "#93c5fd", "#c4b5fd", "#f9a8d4", "#a7f3d0",
];
const STORAGE_KEY = "timetable-maker-v5";

const isoToday = () => new Date().toISOString().slice(0, 10);
const addDays = (iso: string, n: number) => {
  const d = new Date(iso + "T00:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
};
const daysBetween = (from: string, to: string): string[] => {
  const out: string[] = [];
  if (!from || !to) return out;
  const start = new Date(from + "T00:00:00");
  const end = new Date(to + "T00:00:00");
  if (end < start) return out;
  const diff = Math.round((end.getTime() - start.getTime()) / 86400000);
  for (let i = 0; i <= diff; i++) out.push(addDays(from, i));
  return out;
};
const dayLabel = (iso: string) => {
  const d = new Date(iso + "T00:00:00");
  return {
    weekday: d.toLocaleDateString(undefined, { weekday: "short" }),
    date: d.toLocaleDateString(undefined, { month: "short", day: "numeric" }),
  };
};
const weekdayOf = (iso: string): number =>
  new Date(iso + "T00:00:00").getDay();
const WEEKDAY_LABELS = ["S", "M", "T", "W", "T", "F", "S"];
const WEEKDAY_FULL = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const courseAllowedOn = (course: Course, iso: string): boolean => {
  const rule = course.allowedWeekdays;
  if (!rule || rule.length === 0) return true;
  return rule.includes(weekdayOf(iso));
};
const courseAllowedSlot = (course: Course, slotIdx: number): boolean => {
  const rule = course.allowedSlots;
  if (!rule || rule.length === 0) return true;
  return rule.includes(slotIdx);
};

const parseHM = (s: string): number => {
  const [h, m] = s.split(":").map((x) => parseInt(x, 10));
  if (Number.isNaN(h) || Number.isNaN(m)) return 0;
  return h * 60 + m;
};
const to12h = (mins: number): string => {
  const total = ((mins % (24 * 60)) + 24 * 60) % (24 * 60);
  const h24 = Math.floor(total / 60);
  const m = total % 60;
  const period = h24 >= 12 ? "PM" : "AM";
  const h = ((h24 + 11) % 12) + 1;
  return `${h}:${String(m).padStart(2, "0")} ${period}`;
};
const slotLabel = (s: Slot) => `${to12h(parseHM(s.start))} – ${to12h(parseHM(s.end))}`;
const slotMinutes = (s: Slot) => Math.max(0, parseHM(s.end) - parseHM(s.start));

const DEFAULT_SLOTS: Slot[] = [
  { start: "08:50", end: "09:45" },
  { start: "09:45", end: "10:40" },
  { start: "10:40", end: "10:50", isBreak: true },
  { start: "10:50", end: "11:45" },
  { start: "11:45", end: "12:35" },
  { start: "12:35", end: "13:25", isBreak: true },
  { start: "13:25", end: "14:20" },
  { start: "14:20", end: "14:30", isBreak: true },
  { start: "14:30", end: "15:25" },
  { start: "15:25", end: "16:15" },
];

function defaultState(): State {
  const from = isoToday();
  const to = addDays(from, 4);
  const seedCourses: Course[] = [
    { id: "c1", name: "Mathematics", faculty: "Dr. Smith", color: COLORS[0], durationSlots: 1, weeklyPeriods: 4 },
    { id: "c2", name: "Physics", faculty: "Dr. Jones", color: COLORS[2], durationSlots: 1, weeklyPeriods: 3 },
    { id: "c3", name: "Chemistry", faculty: "Dr. Patel", color: COLORS[4], durationSlots: 1, weeklyPeriods: 3 },
  ];
  const mkGrid = (): Record<string, Cell> => ({});
  const cloneCourses = (): Course[] => seedCourses.map((c) => ({ ...c }));
  return {
    fromDate: from,
    toDate: to,
    slots: DEFAULT_SLOTS,
    classes: [
      { id: "k1", name: "Class A", grid: mkGrid(), courses: cloneCourses() },
      { id: "k2", name: "Class B", grid: mkGrid(), courses: cloneCourses() },
    ],
  };
}

type Tool =
  | { kind: "course"; courseId: string }
  | { kind: "break" }
  | { kind: "blocked" }
  | { kind: "erase" };

function Index() {
  const [state, setState] = useState<State>(defaultState);
  const [activeClassId, setActiveClassId] = useState("k1");
  const [hydrated, setHydrated] = useState(false);
  const [picker, setPicker] = useState<{ date: string; slotIdx: number } | null>(null);
  const [armedTool, setArmedTool] = useState<Tool | null>(null);
  const [isPainting, setIsPainting] = useState(false);
  const [rulesFor, setRulesFor] = useState<string | null>(null); // courseId in activeClass
  const [bulkWeekdays, setBulkWeekdays] = useState<number[]>([]);
  const [bulkSlots, setBulkSlots] = useState<number[]>([]);
  const [bulkAllClasses, setBulkAllClasses] = useState(false);
  const gridRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setHydrated(true);
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as Partial<State>;
        const base = defaultState();
        // Migrate v4 → v5: top-level `courses` moved into each class
        const legacyCourses: Course[] | undefined = (parsed as { courses?: Course[] }).courses;
        const migratedClasses: ClassData[] = (parsed.classes ?? base.classes).map(
          (cls) => ({
            ...cls,
            courses:
              (cls as ClassData).courses ??
              (legacyCourses ? legacyCourses.map((c) => ({ ...c })) : []),
          }),
        );
        const merged: State = {
          ...base,
          ...parsed,
          slots: parsed.slots ?? base.slots,
          classes: migratedClasses,
        };
        setState(merged);
        setActiveClassId(merged.classes[0]?.id ?? "");
      }
    } catch {}
  }, []);
  useEffect(() => {
    if (hydrated) localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, [state, hydrated]);

  // Slot editors
  const clipGrids = (classes: ClassData[], maxIdx: number): ClassData[] =>
    classes.map((cls) => {
      const grid: Record<string, Cell> = {};
      Object.entries(cls.grid).forEach(([k, v]) => {
        const m = k.match(/^(.+)-(\d+)$/);
        if (!m) return;
        if (parseInt(m[2], 10) < maxIdx) grid[k] = v;
      });
      return { ...cls, grid };
    });
  const updateSlot = (i: number, patch: Partial<Slot>) =>
    setState((s) => ({
      ...s,
      slots: s.slots.map((sl, idx) => (idx === i ? { ...sl, ...patch } : sl)),
    }));
  const addSlot = () =>
    setState((s) => {
      const last = s.slots[s.slots.length - 1];
      const start = last ? last.end : "09:00";
      const startM = parseHM(start);
      const endM = Math.min(24 * 60 - 1, startM + 55);
      const end = `${String(Math.floor(endM / 60)).padStart(2, "0")}:${String(endM % 60).padStart(2, "0")}`;
      return { ...s, slots: [...s.slots, { start, end }] };
    });
  const removeSlot = (i: number) =>
    setState((s) => ({
      ...s,
      slots: s.slots.filter((_, idx) => idx !== i),
      classes: clipGrids(s.classes, s.slots.length - 1),
    }));
  const toggleSlotBreak = (i: number) =>
    setState((s) => ({
      ...s,
      slots: s.slots.map((sl, idx) => (idx === i ? { ...sl, isBreak: !sl.isBreak } : sl)),
    }));

  useEffect(() => {
    const up = () => setIsPainting(false);
    window.addEventListener("mouseup", up);
    return () => window.removeEventListener("mouseup", up);
  }, []);

  const dates = useMemo(() => daysBetween(state.fromDate, state.toDate), [state.fromDate, state.toDate]);
  const activeClass = state.classes.find((c) => c.id === activeClassId) ?? state.classes[0];

  // Conflict detection across classes
  const conflicts = useMemo(() => {
    const set = new Set<string>();
    dates.forEach((date) => {
      state.slots.forEach((_, i) => {
        const key = `${date}-${i}`;
        const facultyToClass: Record<string, string[]> = {};
        state.classes.forEach((cls) => {
          const cell = cls.grid[key];
          if (cell?.kind === "course") {
            const course = cls.courses.find((c) => c.id === cell.courseId);
            if (!course) return;
            (facultyToClass[course.faculty] ??= []).push(cls.id);
            // Rule violation: course placed on a weekday or period it isn't allowed
            if (!courseAllowedOn(course, date) || !courseAllowedSlot(course, i)) {
              set.add(`${cls.id}:${key}`);
            }
          }
        });
        Object.values(facultyToClass).forEach((clsIds) => {
          if (clsIds.length > 1) clsIds.forEach((id) => set.add(`${id}:${key}`));
        });
      });
    });
    return set;
  }, [state, dates]);

  const applyTool = (date: string, slotIdx: number, tool: Tool) => {
    // Enforce course rules — silently skip disallowed dates
    if (tool.kind === "course") {
      const active = state.classes.find((c) => c.id === activeClassId);
      const course = active?.courses.find((c) => c.id === tool.courseId);
      if (course && !courseAllowedOn(course, date)) return;
    }
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        if (cls.id !== activeClassId) return cls;
        const grid = { ...cls.grid };
        // How many slots does this tool span?
        let span = 1;
        if (tool.kind === "course") {
          const course = cls.courses.find((c) => c.id === tool.courseId);
          span = Math.max(1, course?.durationSlots ?? 1);
        }
        for (let k = 0; k < span; k++) {
          const idx = slotIdx + k;
          if (idx >= s.slots.length) break;
          if (s.slots[idx].isBreak) continue; // never write into break slots
          if (tool.kind === "course") {
            const course = cls.courses.find((c) => c.id === tool.courseId);
            if (course && !courseAllowedSlot(course, idx)) continue;
          }
          const key = `${date}-${idx}`;
          if (tool.kind === "erase") delete grid[key];
          else if (tool.kind === "break") grid[key] = { kind: "break", label: "Break" };
          else if (tool.kind === "blocked") grid[key] = { kind: "blocked", label: "Blocked" };
          else grid[key] = { kind: "course", courseId: tool.courseId };
        }
        return { ...cls, grid };
      }),
    }));
  };

  // Bulk block/break by weekday + slot indices
  const applyBulk = (
    weekdays: number[], // 0..6 (Sun..Sat)
    slotIdxs: number[],
    kind: "blocked" | "break" | "erase",
    allClasses: boolean,
  ) => {
    setState((s) => {
      const wdSet = new Set(weekdays);
      const sSet = new Set(slotIdxs);
      const targetDates = daysBetween(s.fromDate, s.toDate).filter((iso) =>
        wdSet.has(new Date(iso + "T00:00:00").getDay()),
      );
      return {
        ...s,
        classes: s.classes.map((cls) => {
          if (!allClasses && cls.id !== activeClassId) return cls;
          const grid = { ...cls.grid };
          targetDates.forEach((date) => {
            s.slots.forEach((_, i) => {
              if (!sSet.has(i)) return;
              if (s.slots[i].isBreak) return;
              const key = `${date}-${i}`;
              if (kind === "erase") delete grid[key];
              else if (kind === "break") grid[key] = { kind: "break", label: "Break" };
              else grid[key] = { kind: "blocked", label: "Blocked" };
            });
          });
          return { ...cls, grid };
        }),
      };
    });
  };

  const onCellMouseDown = (date: string, slotIdx: number, e: React.MouseEvent) => {
    if (state.slots[slotIdx]?.isBreak) { e.preventDefault(); return; }
    // Alt + right-click erases immediately
    if (e.button === 2 && e.altKey) {
      e.preventDefault();
      applyTool(date, slotIdx, { kind: "erase" });
      return;
    }
    // Right-click (no alt) opens picker to change tool
    if (e.button === 2) {
      e.preventDefault();
      setPicker({ date, slotIdx });
      return;
    }
    if (e.button !== 0) return;
    if (armedTool) {
      e.preventDefault();
      applyTool(date, slotIdx, armedTool);
      setIsPainting(true);
    } else {
      setPicker({ date, slotIdx });
    }
  };
  const onCellEnter = (date: string, slotIdx: number, e: React.MouseEvent) => {
    // Only paint while a mouse button is actually held down
    if (isPainting && armedTool && (e.buttons & 1) === 1) {
      applyTool(date, slotIdx, armedTool);
    }
  };

  const pickTool = (tool: Tool) => {
    if (!picker) return;
    applyTool(picker.date, picker.slotIdx, tool);
    setArmedTool(tool); // arm for subsequent click / click-drag
    setPicker(null);
  };

  // Class + course + slot editors
  const addClass = () =>
    setState((s) => {
      const id = `k${Date.now()}`;
      return {
        ...s,
        classes: [
          ...s.classes,
          { id, name: `Class ${String.fromCharCode(65 + s.classes.length)}`, grid: {}, courses: [] },
        ],
      };
    });
  const removeClass = (id: string) => {
    setState((s) =>
      s.classes.length > 1 ? { ...s, classes: s.classes.filter((c) => c.id !== id) } : s,
    );
    if (activeClassId === id) {
      const next = state.classes.find((c) => c.id !== id);
      if (next) setActiveClassId(next.id);
    }
  };
  const renameClass = (id: string, name: string) =>
    setState((s) => ({ ...s, classes: s.classes.map((c) => (c.id === id ? { ...c, name } : c)) }));

  // Courses are per-class — all edits scope to the active class
  const addCourse = () =>
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        if (cls.id !== activeClassId) return cls;
        return {
          ...cls,
          courses: [
            ...cls.courses,
            {
              id: `c${Date.now()}`,
              name: "New Course",
              faculty: "Faculty",
              color: COLORS[cls.courses.length % COLORS.length],
              durationSlots: 1,
              weeklyPeriods: 3,
            },
          ],
        };
      }),
    }));
  const updateCourse = (id: string, patch: Partial<Course>) =>
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        if (cls.id !== activeClassId) return cls;
        return {
          ...cls,
          courses: cls.courses.map((c) => (c.id === id ? { ...c, ...patch } : c)),
        };
      }),
    }));
  const removeCourse = (id: string) =>
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        if (cls.id !== activeClassId) return cls;
        const grid = { ...cls.grid };
        Object.keys(grid).forEach((k) => {
          const cell = grid[k];
          if (cell.kind === "course" && cell.courseId === id) delete grid[k];
        });
        return { ...cls, grid, courses: cls.courses.filter((c) => c.id !== id) };
      }),
    }));

  // Export
  const buildSheet = (cls: ClassData) => {
    const rows: string[][] = [];
    rows.push(["Day / Date", ...state.slots.map(slotLabel)]);
    dates.forEach((date) => {
      const { weekday, date: dstr } = dayLabel(date);
      const row = [`${weekday} ${dstr}`];
      state.slots.forEach((sl, i) => {
        if (sl.isBreak) { row.push("Break"); return; }
        const cell = cls.grid[`${date}-${i}`];
        if (!cell) row.push("");
        else if (cell.kind === "break") row.push(`Break: ${cell.label}`);
        else if (cell.kind === "blocked") row.push(`Blocked: ${cell.label}`);
        else if (cell.kind === "course") {
          const c = cls.courses.find((x) => x.id === cell.courseId);
          row.push(c ? `${c.name} (${c.faculty})` : "");
        } else row.push("");
      });
      rows.push(row);
    });
    return rows;
  };
  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    state.classes.forEach((cls) => {
      const ws = XLSX.utils.aoa_to_sheet(buildSheet(cls));
      XLSX.utils.book_append_sheet(wb, ws, cls.name.slice(0, 31) || "Class");
    });
    XLSX.writeFile(wb, `timetable_${state.fromDate}_to_${state.toDate}.xlsx`);
  };
  const exportCSV = () => {
    state.classes.forEach((cls) => {
      const csv = buildSheet(cls)
        .map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","))
        .join("\n");
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `timetable_${cls.name}_${state.fromDate}_to_${state.toDate}.csv`;
      a.click();
      URL.revokeObjectURL(url);
    });
  };

  const cellDisplay = (cell: Cell | undefined, courses: Course[]) => {
    if (!cell || cell.kind === "empty") return { text: "", bg: "#fff", fg: "#94a3b8" };
    if (cell.kind === "break") return { text: cell.label, bg: "#fef3c7", fg: "#92400e" };
    if (cell.kind === "blocked") return { text: cell.label, bg: "#e5e7eb", fg: "#374151" };
    const course = courses.find((c) => c.id === cell.courseId);
    return {
      text: course ? `${course.name}\n${course.faculty}` : "?",
      bg: course?.color ?? "#ddd",
      fg: "#1f2937",
    };
  };

  const hasConflicts = conflicts.size > 0;

  return (
    <div
      className="min-h-screen w-full bg-[#f5f3ee] text-[#2d2d2d]"
      style={{ fontFamily: "'Manrope', system-ui, sans-serif" }}
      onMouseLeave={() => setIsPainting(false)}
    >
      <div className="mx-auto flex min-h-screen w-full max-w-[1500px] flex-col md:flex-row md:border-x-2 md:border-[#0d0d0d]">
        {/* Sidebar */}
        <aside className="w-full shrink-0 border-b-2 border-[#0d0d0d] bg-[#e8e4dd] md:w-[320px] md:border-b-0 md:border-r-2">
          <div className="border-b border-[#0d0d0d]/10 bg-[#0d0d0d] px-6 py-5 text-[#f5f3ee]">
            <h1
              className="text-xl font-bold tracking-tight"
              style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
            >
              Timetable Maker
            </h1>
            <p className="mt-1 text-[11px] uppercase tracking-[0.2em] text-[#f5f3ee]/60">
              Modular · Weekly · Editorial
            </p>
          </div>

          <div className="space-y-8 p-5">
            {/* Classes */}
            <section>
              <div className="mb-3 flex items-center justify-between">
                <h3
                  className="text-[11px] font-bold uppercase tracking-widest text-[#0d0d0d]"
                  style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                >
                  Classes
                </h3>
                <button
                  onClick={addClass}
                  className="border border-[#0d0d0d] bg-[#f5f3ee] px-2 py-0.5 text-[10px] font-bold uppercase tracking-wider hover:bg-[#0d0d0d] hover:text-[#f5f3ee]"
                >
                  + Add
                </button>
              </div>
              <div className="space-y-2">
                {state.classes.map((cls) => {
                  const active = cls.id === activeClassId;
                  return (
                    <div key={cls.id} className="flex items-stretch gap-1">
                      <button
                        onClick={() => setActiveClassId(cls.id)}
                        className={`flex-1 border px-3 py-2 text-left text-sm ${
                          active
                            ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee]"
                            : "border-[#0d0d0d]/20 bg-white hover:border-[#0d0d0d]"
                        }`}
                      >
                        <input
                          value={cls.name}
                          onClick={(e) => e.stopPropagation()}
                          onChange={(e) => renameClass(cls.id, e.target.value)}
                          className="w-full bg-transparent font-semibold outline-none"
                        />
                      </button>
                      {state.classes.length > 1 && (
                        <button
                          onClick={() => removeClass(cls.id)}
                          className="border border-[#0d0d0d]/20 bg-white px-2 text-xs text-[#2d2d2d]/50 hover:border-red-500 hover:text-red-600"
                          aria-label="Remove class"
                        >
                          ×
                        </button>
                      )}
                    </div>
                  );
                })}
              </div>
            </section>

            {/* Courses */}
            <section>
              <div className="mb-3 flex items-center justify-between">
                <h3
                  className="text-[11px] font-bold uppercase tracking-widest text-[#0d0d0d]"
                  style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                >
                  Courses · {activeClass?.name ?? ""}
                </h3>
                <button
                  onClick={addCourse}
                  className="border border-[#0d0d0d] bg-[#f5f3ee] px-2 py-0.5 text-[10px] font-bold uppercase tracking-wider hover:bg-[#0d0d0d] hover:text-[#f5f3ee]"
                >
                  + Add
                </button>
              </div>
              <div className="space-y-2">
                {(activeClass?.courses ?? []).map((c) => (
                  <div key={c.id} className="border border-[#0d0d0d]/30 bg-white">
                    <div className="flex items-center gap-2 border-b border-[#0d0d0d]/10 px-3 py-2">
                      <span
                        className="h-3 w-3 shrink-0"
                        style={{ backgroundColor: c.color }}
                      />
                      <input
                        value={c.name}
                        onChange={(e) => updateCourse(c.id, { name: e.target.value })}
                        placeholder="Course"
                        className="min-w-0 flex-1 bg-transparent text-sm font-bold outline-none"
                      />
                      <button
                        onClick={() => removeCourse(c.id)}
                        className="text-xs text-[#2d2d2d]/40 hover:text-red-600"
                        aria-label="Remove course"
                      >
                        ×
                      </button>
                    </div>
                    <div className="grid grid-cols-[minmax(0,1fr)_auto] gap-2 px-3 py-2">
                      <input
                        value={c.faculty}
                        onChange={(e) => updateCourse(c.id, { faculty: e.target.value })}
                        placeholder="Faculty"
                        className="min-w-0 border-b border-dashed border-[#0d0d0d]/20 bg-transparent text-xs text-[#2d2d2d]/70 outline-none focus:border-[#0d0d0d]"
                      />
                      <label
                        className="flex items-center gap-1 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        <input
                          type="number"
                          min={1}
                          value={c.durationSlots}
                          onChange={(e) =>
                            updateCourse(c.id, {
                              durationSlots: Math.max(1, parseInt(e.target.value || "1", 10)),
                            })
                          }
                          className="w-10 border border-[#0d0d0d]/20 bg-white px-1 py-0.5 text-center text-xs"
                        />
                        <span>slot</span>
                      </label>
                    </div>
                    <button
                      onClick={() => setRulesFor(c.id)}
                      className="flex w-full items-center justify-between border-t border-dashed border-[#0d0d0d]/15 px-3 py-2 text-left hover:bg-[#f5f3ee]"
                    >
                      <span
                        className="text-[10px] font-bold uppercase tracking-wider text-[#2d2d2d]/60"
                        style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                      >
                        Available days
                      </span>
                      <span
                        className="truncate text-[10px] text-[#2d2d2d]/70"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        {c.allowedWeekdays && c.allowedWeekdays.length > 0
                          ? c.allowedWeekdays.map((w) => WEEKDAY_FULL[w]).join(" ")
                          : "All days"}
                        {" · "}
                        {c.allowedSlots && c.allowedSlots.length > 0
                          ? `P${c.allowedSlots.map((i) => i + 1).join(" P")}`
                          : "All periods"}
                      </span>
                    </button>
                  </div>
                ))}
              </div>
            </section>

            {/* Bulk block/break */}
            <section className="border border-dashed border-[#0d0d0d]/50 bg-[#f5f3ee] p-4">
              <h3
                className="mb-2 text-[11px] font-bold uppercase tracking-widest"
                style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
              >
                Bulk Block / Break
              </h3>
              <p className="mb-3 text-[11px] text-[#2d2d2d]/70">
                Pick weekdays and slots to apply an action across the whole date range.
              </p>

              <div className="mb-3">
                <div className="mb-1 text-[10px] font-bold uppercase tracking-wider text-[#2d2d2d]/60">
                  Weekdays
                </div>
                <div className="flex flex-wrap gap-1">
                  {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((w, i) => {
                    const on = bulkWeekdays.includes(i);
                    return (
                      <button
                        key={i}
                        onClick={() =>
                          setBulkWeekdays((prev) =>
                            prev.includes(i) ? prev.filter((x) => x !== i) : [...prev, i],
                          )
                        }
                        className={`border px-2 py-1 text-[10px] font-bold uppercase tracking-wider ${
                          on
                            ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee]"
                            : "border-[#0d0d0d]/30 bg-white text-[#2d2d2d]"
                        }`}
                      >
                        {w}
                      </button>
                    );
                  })}
                </div>
              </div>

              <div className="mb-3">
                <div className="mb-1 text-[10px] font-bold uppercase tracking-wider text-[#2d2d2d]/60">
                  Time slots
                </div>
                <div
                  className="max-h-36 space-y-1 overflow-y-auto border border-[#0d0d0d]/20 bg-white p-2 text-[11px]"
                  style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                >
                  {state.slots.map((slot, i) => {
                    const on = bulkSlots.includes(i);
                    return (
                      <label key={i} className="flex items-center gap-2">
                        <input
                          type="checkbox"
                          checked={on}
                          onChange={() =>
                            setBulkSlots((prev) =>
                              prev.includes(i) ? prev.filter((x) => x !== i) : [...prev, i],
                            )
                          }
                          className="accent-[#0d0d0d]"
                        />
                        <span>
                          {slotLabel(slot)}
                          {slot.isBreak && (
                            <span className="ml-1 text-[#b45309]">·break</span>
                          )}
                        </span>
                      </label>
                    );
                  })}
                </div>
              </div>

              <label className="mb-3 flex items-center gap-2 text-[11px]">
                <input
                  type="checkbox"
                  checked={bulkAllClasses}
                  onChange={(e) => setBulkAllClasses(e.target.checked)}
                  className="accent-[#0d0d0d]"
                />
                <span className="font-medium">Apply to all classes</span>
              </label>

              <div className="grid grid-cols-3 gap-1">
                <button
                  onClick={() => applyBulk(bulkWeekdays, bulkSlots, "blocked", bulkAllClasses)}
                  disabled={bulkWeekdays.length === 0 || bulkSlots.length === 0}
                  className="border border-[#0d0d0d] bg-white px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-[#e8e4dd] disabled:opacity-40"
                >
                  Block
                </button>
                <button
                  onClick={() => applyBulk(bulkWeekdays, bulkSlots, "break", bulkAllClasses)}
                  disabled={bulkWeekdays.length === 0 || bulkSlots.length === 0}
                  className="border border-[#b45309] bg-[#d97706] px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider text-white hover:bg-[#b45309] disabled:opacity-40"
                >
                  Break
                </button>
                <button
                  onClick={() => applyBulk(bulkWeekdays, bulkSlots, "erase", bulkAllClasses)}
                  disabled={bulkWeekdays.length === 0 || bulkSlots.length === 0}
                  className="border border-[#0d0d0d]/40 bg-white px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:border-[#0d0d0d] disabled:opacity-40"
                >
                  Clear
                </button>
              </div>
            </section>

            {/* Time slots editor */}
            <section>
              <div className="mb-3 flex items-center justify-between">
                <h3
                  className="text-[11px] font-bold uppercase tracking-widest text-[#0d0d0d]"
                  style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                >
                  Time Slots
                </h3>
                <button
                  onClick={addSlot}
                  className="border border-[#0d0d0d] bg-[#f5f3ee] px-2 py-0.5 text-[10px] font-bold uppercase tracking-wider hover:bg-[#0d0d0d] hover:text-[#f5f3ee]"
                >
                  + Add
                </button>
              </div>
              <div
                className="space-y-1 text-[11px]"
                style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
              >
                {state.slots.map((sl, i) => (
                  <div
                    key={i}
                    className={`flex items-center gap-1 border px-1.5 py-1 ${
                      sl.isBreak
                        ? "border-[#d97706]/40 bg-[#fef3c7]"
                        : "border-[#0d0d0d]/20 bg-white"
                    }`}
                  >
                    <span className="w-4 text-center text-[#2d2d2d]/40">{i + 1}</span>
                    <input
                      type="time"
                      value={sl.start}
                      onChange={(e) => updateSlot(i, { start: e.target.value })}
                      className="w-20 border-b border-dashed border-[#0d0d0d]/20 bg-transparent px-0.5 py-0.5 outline-none focus:border-[#0d0d0d]"
                    />
                    <span className="text-[#2d2d2d]/40">—</span>
                    <input
                      type="time"
                      value={sl.end}
                      onChange={(e) => updateSlot(i, { end: e.target.value })}
                      className="w-20 border-b border-dashed border-[#0d0d0d]/20 bg-transparent px-0.5 py-0.5 outline-none focus:border-[#0d0d0d]"
                    />
                    <button
                      onClick={() => toggleSlotBreak(i)}
                      title="Toggle break"
                      className={`ml-auto border px-1.5 py-0.5 text-[9px] font-bold uppercase tracking-wider ${
                        sl.isBreak
                          ? "border-[#b45309] bg-[#d97706] text-white"
                          : "border-[#0d0d0d]/30 bg-[#f5f3ee] text-[#2d2d2d]"
                      }`}
                    >
                      Brk
                    </button>
                    <button
                      onClick={() => removeSlot(i)}
                      className="px-1 text-[#2d2d2d]/40 hover:text-red-600"
                    >
                      ×
                    </button>
                  </div>
                ))}
              </div>
            </section>
          </div>
        </aside>

        {/* Main */}
        <main className="flex min-w-0 flex-1 flex-col">
          {/* Toolbar */}
          <header className="grid grid-cols-[minmax(0,1fr)_auto] items-center gap-4 border-b-2 border-[#0d0d0d] bg-white/60 px-4 py-3 sm:flex sm:flex-wrap sm:justify-between sm:px-8 sm:py-4">
            <div className="flex min-w-0 flex-wrap items-end gap-4">
              <div>
                <div className="text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/50">
                  From
                </div>
                <input
                  type="date"
                  value={state.fromDate}
                  onChange={(e) => setState((s) => ({ ...s, fromDate: e.target.value }))}
                  className="border-b border-[#0d0d0d] bg-transparent py-0.5 text-sm font-semibold outline-none"
                  style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                />
              </div>
              <span className="pb-1 text-lg text-[#2d2d2d]/30">/</span>
              <div>
                <div className="text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/50">
                  To
                </div>
                <input
                  type="date"
                  value={state.toDate}
                  onChange={(e) => setState((s) => ({ ...s, toDate: e.target.value }))}
                  className="border-b border-[#0d0d0d] bg-transparent py-0.5 text-sm font-semibold outline-none"
                  style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                />
              </div>
            </div>
            <div className="flex shrink-0 items-center gap-2">
              <button
                onClick={exportCSV}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5"
              >
                CSV
              </button>
              <button
                onClick={exportExcel}
                className="border-2 border-[#0d0d0d] bg-[#0d0d0d] px-3 py-2 text-[11px] font-bold uppercase tracking-wider text-[#f5f3ee] transition-transform hover:opacity-90 active:translate-y-0.5"
              >
                Export Excel
              </button>
            </div>
          </header>

          {/* Status strip */}
          <div className="flex flex-wrap items-center gap-3 border-b border-[#0d0d0d]/10 bg-[#f5f3ee] px-4 py-2 text-xs sm:px-8">
            <div className="flex min-w-0 items-center gap-2">
              <span className="text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/50">
                Class
              </span>
              <span
                className="truncate font-bold"
                style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
              >
                {activeClass?.name}
              </span>
            </div>
            <span className="text-[#2d2d2d]/20">·</span>
            <span
              className="text-[11px] text-[#2d2d2d]/60"
              style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
            >
              {dates.length} day{dates.length === 1 ? "" : "s"} · {state.slots.length} slots
            </span>
            {armedTool && (
              <div className="ml-auto flex items-center gap-2 border-2 border-[#0d0d0d] bg-white px-2 py-1">
                <span className="text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/50">
                  Tool
                </span>
                {armedTool.kind === "course" ? (
                  <span className="flex items-center gap-1">
                    <span
                      className="inline-block h-3 w-3"
                      style={{
                        backgroundColor:
                          (activeClass?.courses ?? []).find((c) => c.id === armedTool.courseId)?.color ?? "#ddd",
                      }}
                    />
                    <span className="font-bold">
                      {(activeClass?.courses ?? []).find((c) => c.id === armedTool.courseId)?.name ?? "?"}
                    </span>
                  </span>
                ) : (
                  <span className="font-bold uppercase tracking-wider">{armedTool.kind}</span>
                )}
                <button
                  onClick={() => setArmedTool(null)}
                  className="text-[#2d2d2d]/40 hover:text-red-600"
                  aria-label="Clear tool"
                >
                  ×
                </button>
              </div>
            )}
          </div>

          {/* Alerts */}
          <div className="space-y-2 px-4 pt-4 sm:px-8">
            {hasConflicts && (
              <div className="border-2 border-red-600 bg-red-50 px-3 py-2 text-xs font-semibold text-red-700">
                Faculty conflict — the same teacher is scheduled in two classes at the same time (red cells).
              </div>
            )}
            {dates.length === 0 && (
              <div className="border-2 border-[#d97706] bg-[#fef3c7] px-3 py-2 text-xs font-semibold text-[#b45309]">
                Pick a valid date range.
              </div>
            )}
            <p className="text-[11px] text-[#2d2d2d]/60">
              Click any cell to pick a course. Click-drag to paint. Right-click to change tool. Alt+right-click to erase.
            </p>
          </div>

          {/* Grid */}
          <div className="flex-1 overflow-auto p-4 sm:p-8">
            <div className="inline-block min-w-full" ref={gridRef}>
              <table className="w-full border-collapse border-2 border-[#0d0d0d] text-left">
                <thead>
                  <tr className="bg-[#0d0d0d] text-[#f5f3ee]">
                    <th
                      className="sticky left-0 z-10 w-32 border border-[#f5f3ee]/20 bg-[#0d0d0d] p-3 text-[10px] font-bold uppercase tracking-widest"
                      style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                    >
                      Day
                    </th>
                    {state.slots.map((slot, i) => (
                      <th
                        key={i}
                        className="border border-[#f5f3ee]/20 p-2 text-center align-middle"
                        style={{ minWidth: 110 }}
                      >
                        <div
                          className="whitespace-nowrap text-[10px] font-bold uppercase tracking-wider"
                          style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                        >
                          {slotLabel(slot)}
                        </div>
                        {slot.isBreak && (
                          <div className="mt-0.5 text-[9px] font-bold uppercase tracking-widest text-[#fbbf24]">
                            Break
                          </div>
                        )}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="bg-white">
                  {dates.map((date) => {
                    const { weekday, date: dstr } = dayLabel(date);
                    return (
                      <tr key={date}>
                        <th
                          className="sticky left-0 z-10 border-2 border-[#0d0d0d] bg-[#e8e4dd] p-3 text-left align-middle"
                          style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                        >
                          <div className="text-xs font-bold uppercase tracking-wider">
                            {weekday}
                          </div>
                          <div className="text-[10px] font-medium text-[#2d2d2d]/60">
                            {dstr}
                          </div>
                        </th>
                        {state.slots.map((sl, i) => {
                          if (!activeClass) return null;
                          const key = `${date}-${i}`;
                          const cell = activeClass.grid[key];
                          const isConflict = conflicts.has(`${activeClass.id}:${key}`);

                          if (sl.isBreak) {
                            return (
                              <td
                                key={i}
                                onContextMenu={(e) => e.preventDefault()}
                                className="border border-[#0d0d0d]/10 bg-[#fef3c7] p-1 text-center align-middle"
                                style={{ minWidth: 110, height: 64 }}
                              >
                                <div className="text-[10px] font-bold uppercase tracking-widest text-[#b45309]">
                                  Break
                                </div>
                              </td>
                            );
                          }

                          const course =
                            cell?.kind === "course"
                              ? (activeClass?.courses ?? []).find((c) => c.id === cell.courseId)
                              : undefined;

                          return (
                            <td
                              key={i}
                              onMouseDown={(e) => onCellMouseDown(date, i, e)}
                              onMouseEnter={(e) => onCellEnter(date, i, e)}
                              onContextMenu={(e) => e.preventDefault()}
                              className={`p-1 align-middle ${
                                isConflict
                                  ? "border-2 border-red-600"
                                  : "border border-[#0d0d0d]/10"
                              } cursor-pointer`}
                              style={{ minWidth: 110, height: 64 }}
                            >
                              {!cell || cell.kind === "empty" ? (
                                <div className="flex h-full min-h-[52px] items-center justify-center border-2 border-dashed border-[#0d0d0d]/10 text-[#2d2d2d]/25 transition-colors hover:border-[#0d0d0d]/60 hover:text-[#0d0d0d]">
                                  +
                                </div>
                              ) : cell.kind === "break" ? (
                                <div className="flex h-full min-h-[52px] items-center justify-center bg-[#fef3c7] text-[10px] font-bold uppercase tracking-widest text-[#b45309]">
                                  {cell.label}
                                </div>
                              ) : cell.kind === "blocked" ? (
                                <div className="flex h-full min-h-[52px] items-center justify-center bg-[repeating-linear-gradient(45deg,#e8e4dd_0_6px,#d9d5ce_6px_12px)] text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/60">
                                  {cell.label}
                                </div>
                              ) : (
                                <div
                                  className={`flex h-full min-h-[52px] flex-col justify-between p-1.5 ${
                                    isConflict ? "bg-red-100" : ""
                                  }`}
                                  style={
                                    isConflict
                                      ? undefined
                                      : {
                                          backgroundColor: `${course?.color ?? "#ddd"}55`,
                                          borderLeft: `4px solid ${course?.color ?? "#ddd"}`,
                                        }
                                  }
                                >
                                  <div
                                    className={`text-[11px] font-bold leading-tight ${
                                      isConflict ? "text-red-700" : "text-[#0d0d0d]"
                                    }`}
                                    style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                                  >
                                    {course?.name ?? "?"}
                                  </div>
                                  <div
                                    className={`text-[9px] uppercase tracking-wider ${
                                      isConflict ? "font-bold text-red-700" : "text-[#2d2d2d]/60"
                                    }`}
                                  >
                                    {isConflict ? "Conflict · " : ""}
                                    {course?.faculty ?? ""}
                                  </div>
                                </div>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </main>
      </div>

      {/* Picker modal */}
      {picker && (
        <div
          className="fixed inset-0 z-40 bg-[#0d0d0d]/40 p-4 backdrop-blur-sm"
          onClick={() => setPicker(null)}
        >
          <div
            className="absolute left-1/2 top-1/2 w-[min(400px,calc(100%-2rem))] -translate-x-1/2 -translate-y-1/2 border-2 border-[#0d0d0d] bg-[#f5f3ee] shadow-[8px_8px_0px_0px_#0d0d0d]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="border-b-2 border-[#0d0d0d] bg-[#0d0d0d] px-4 py-3 text-[#f5f3ee]">
              <div className="text-[10px] font-bold uppercase tracking-widest text-[#f5f3ee]/60">
                Assign Slot
              </div>
              <div
                className="mt-0.5 text-sm font-bold"
                style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
              >
                {dayLabel(picker.date).weekday} {dayLabel(picker.date).date}
                {" · "}
                <span style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}>
                  {state.slots[picker.slotIdx] ? slotLabel(state.slots[picker.slotIdx]) : ""}
                </span>
              </div>
            </div>

            <div className="max-h-[60vh] overflow-y-auto p-4">
              <div className="mb-2 text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/60">
                Courses
              </div>
              <div className="mb-4 space-y-1">
                {(activeClass?.courses ?? []).length === 0 && (
                  <div className="text-xs text-[#2d2d2d]/60">
                    No courses yet. Add one from the sidebar.
                  </div>
                )}
                {(activeClass?.courses ?? []).map((c) => {
                  const allowed =
                    courseAllowedOn(c, picker.date) &&
                    courseAllowedSlot(c, picker.slotIdx);
                  const ruleLabel =
                    (c.allowedWeekdays && c.allowedWeekdays.length > 0) ||
                    (c.allowedSlots && c.allowedSlots.length > 0)
                      ? [
                          c.allowedWeekdays && c.allowedWeekdays.length > 0
                            ? c.allowedWeekdays.map((w) => WEEKDAY_FULL[w]).join(",")
                            : "any day",
                          c.allowedSlots && c.allowedSlots.length > 0
                            ? "P" + c.allowedSlots.map((i) => i + 1).join(",")
                            : "any period",
                        ].join(" · ")
                      : null;
                  return (
                    <button
                      key={c.id}
                      disabled={!allowed}
                      onClick={() => pickTool({ kind: "course", courseId: c.id })}
                      className={
                        "flex w-full items-center gap-3 border px-3 py-2 text-left transition-colors " +
                        (allowed
                          ? "border-[#0d0d0d]/30 bg-white hover:border-[#0d0d0d]"
                          : "cursor-not-allowed border-[#0d0d0d]/10 bg-[#f5f3ee] opacity-50")
                      }
                    >
                      <span className="h-4 w-4 shrink-0" style={{ backgroundColor: c.color }} />
                      <span className="min-w-0 flex-1">
                        <span
                          className="block truncate text-sm font-bold"
                          style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                        >
                          {c.name}
                        </span>
                        <span className="block truncate text-[11px] text-[#2d2d2d]/60">
                          {c.faculty}
                          {ruleLabel && ` · ${ruleLabel} only`}
                        </span>
                      </span>
                      {!allowed && (
                        <span className="shrink-0 text-[9px] font-bold uppercase tracking-wider text-red-600">
                          Unavailable
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>

              <div className="mb-2 text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/60">
                Actions
              </div>
              <div className="grid grid-cols-3 gap-2">
                <button
                  onClick={() => pickTool({ kind: "break" })}
                  className="border-2 border-[#b45309] bg-[#fef3c7] px-2 py-2 text-[10px] font-bold uppercase tracking-wider text-[#b45309] hover:bg-[#fde68a]"
                >
                  Break
                </button>
                <button
                  onClick={() => pickTool({ kind: "blocked" })}
                  className="border-2 border-[#0d0d0d] bg-[#e8e4dd] px-2 py-2 text-[10px] font-bold uppercase tracking-wider text-[#2d2d2d] hover:bg-[#d9d5ce]"
                >
                  Block
                </button>
                <button
                  onClick={() => pickTool({ kind: "erase" })}
                  className="border-2 border-[#0d0d0d]/40 bg-white px-2 py-2 text-[10px] font-bold uppercase tracking-wider hover:border-[#0d0d0d]"
                >
                  Clear
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Rules editor modal */}
      {rulesFor && activeClass && (() => {
        const course = activeClass.courses.find((c) => c.id === rulesFor);
        if (!course) return null;
        const wdRule = course.allowedWeekdays ?? [];
        const wdAll = wdRule.length === 0;
        const slotRule = course.allowedSlots ?? [];
        const slotAll = slotRule.length === 0;
        const nonBreakIdxs = state.slots
          .map((sl, i) => ({ sl, i }))
          .filter((x) => !x.sl.isBreak);
        return (
          <div
            className="fixed inset-0 z-50 flex items-center justify-center bg-[#0d0d0d]/60 p-4"
            onClick={() => setRulesFor(null)}
          >
            <div
              className="w-full max-w-md border-2 border-[#0d0d0d] bg-[#f5f3ee]"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-start justify-between border-b-2 border-[#0d0d0d] bg-[#0d0d0d] px-4 py-3 text-[#f5f3ee]">
                <div className="min-w-0">
                  <div className="text-[10px] uppercase tracking-widest text-[#f5f3ee]/60">
                    Course rules · {activeClass.name}
                  </div>
                  <div
                    className="truncate text-base font-bold"
                    style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                  >
                    {course.name}{" "}
                    <span className="text-xs font-normal text-[#f5f3ee]/60">
                      · {course.faculty}
                    </span>
                  </div>
                </div>
                <button
                  onClick={() => setRulesFor(null)}
                  className="text-lg leading-none text-[#f5f3ee]/60 hover:text-[#f5f3ee]"
                  aria-label="Close"
                >
                  ×
                </button>
              </div>
              <div className="max-h-[70vh] overflow-y-auto p-4 space-y-4">
                <div>
                  <div className="mb-2 flex items-center justify-between">
                    <span className="text-[11px] font-bold uppercase tracking-widest text-[#2d2d2d]/70">
                      Available weekdays
                    </span>
                    <button
                      onClick={() => updateCourse(course.id, { allowedWeekdays: [] })}
                      className="text-[10px] uppercase tracking-wider text-[#2d2d2d]/50 hover:text-[#0d0d0d]"
                    >
                      All days
                    </button>
                  </div>
                  <div className="flex flex-wrap gap-1">
                    {WEEKDAY_LABELS.map((lbl, wd) => {
                      const active = wdAll || wdRule.includes(wd);
                      return (
                        <button
                          key={wd}
                          title={WEEKDAY_FULL[wd]}
                          onClick={() => {
                            const base = wdAll ? [0, 1, 2, 3, 4, 5, 6] : [...wdRule];
                            const next = base.includes(wd)
                              ? base.filter((x) => x !== wd)
                              : [...base, wd].sort();
                            updateCourse(course.id, {
                              allowedWeekdays: next.length === 7 ? [] : next,
                            });
                          }}
                          className={
                            "flex h-9 w-9 items-center justify-center border text-xs font-bold " +
                            (active
                              ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee]"
                              : "border-[#0d0d0d]/20 bg-white text-[#2d2d2d]/40 hover:border-[#0d0d0d]/50")
                          }
                          style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                        >
                          {lbl}
                        </button>
                      );
                    })}
                  </div>
                  <p className="mt-1 text-[10px] text-[#2d2d2d]/50">
                    Pick the weekdays this faculty is available. Deselect all to
                    treat every day as allowed.
                  </p>
                </div>

                <div>
                  <div className="mb-2 flex items-center justify-between">
                    <span className="text-[11px] font-bold uppercase tracking-widest text-[#2d2d2d]/70">
                      Available periods
                    </span>
                    <button
                      onClick={() => updateCourse(course.id, { allowedSlots: [] })}
                      className="text-[10px] uppercase tracking-wider text-[#2d2d2d]/50 hover:text-[#0d0d0d]"
                    >
                      All periods
                    </button>
                  </div>
                  <div className="grid grid-cols-1 gap-1 sm:grid-cols-2">
                    {nonBreakIdxs.map(({ sl, i }) => {
                      const active = slotAll || slotRule.includes(i);
                      const periodNum =
                        state.slots.slice(0, i + 1).filter((x) => !x.isBreak).length;
                      return (
                        <button
                          key={i}
                          onClick={() => {
                            const allIdxs = nonBreakIdxs.map((x) => x.i);
                            const base = slotAll ? [...allIdxs] : [...slotRule];
                            const next = base.includes(i)
                              ? base.filter((x) => x !== i)
                              : [...base, i].sort((a, b) => a - b);
                            updateCourse(course.id, {
                              allowedSlots:
                                next.length === allIdxs.length ? [] : next,
                            });
                          }}
                          className={
                            "flex items-center justify-between gap-2 border px-2 py-1.5 text-left text-xs " +
                            (active
                              ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee]"
                              : "border-[#0d0d0d]/20 bg-white text-[#2d2d2d]/60 hover:border-[#0d0d0d]/50")
                          }
                        >
                          <span
                            className="font-bold"
                            style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                          >
                            P{periodNum}
                          </span>
                          <span
                            className="text-[10px] opacity-80"
                            style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                          >
                            {slotLabel(sl)}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                  <p className="mt-1 text-[10px] text-[#2d2d2d]/50">
                    Pick the periods this course can be scheduled in. Break
                    slots are excluded.
                  </p>
                </div>
              </div>
              <div className="flex justify-end border-t-2 border-[#0d0d0d] bg-[#e8e4dd] px-4 py-2">
                <button
                  onClick={() => setRulesFor(null)}
                  className="border-2 border-[#0d0d0d] bg-[#f5f3ee] px-4 py-1.5 text-[11px] font-bold uppercase tracking-wider hover:bg-[#0d0d0d] hover:text-[#f5f3ee]"
                >
                  Done
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}
