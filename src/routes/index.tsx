import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import XLSXStyle from "xlsx-js-style";

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
  // Total number of sessions to place across the whole date range.
  // When set (>0) this overrides the per-week weeklyPeriods target during auto-fill
  // and drives the planned/remaining counters.
  totalSessions?: number;
  // 0=Sun..6=Sat. undefined or empty = allowed on all days.
  allowedWeekdays?: number[];
  // Slot indices. undefined or empty = allowed in all periods.
  allowedSlots?: number[];
  // Per-weekday period override. If a weekday key is present (even as [])
  // it fully replaces `allowedSlots` for that weekday. Missing key = fall
  // back to `allowedSlots`.
  allowedSlotsByWeekday?: Record<number, number[]>;
  // Optional active date range for this course. Undefined = entire timetable range.
  fromDate?: string; // YYYY-MM-DD
  toDate?: string; // YYYY-MM-DD
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
const INITIAL_FROM_DATE = "2026-07-25";

const isoToday = () => new Date().toISOString().slice(0, 10);
const parseIsoParts = (iso: string): [number, number, number] | null => {
  const match = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  return [parseInt(match[1], 10), parseInt(match[2], 10), parseInt(match[3], 10)];
};
const utcDateFromIso = (iso: string) => {
  const parts = parseIsoParts(iso);
  if (!parts) return null;
  return new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
};
const addDays = (iso: string, n: number) => {
  const d = utcDateFromIso(iso) ?? new Date(Date.UTC(2026, 6, 25));
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
};
const daysBetween = (from: string, to: string): string[] => {
  const out: string[] = [];
  if (!from || !to) return out;
  const start = utcDateFromIso(from);
  const end = utcDateFromIso(to);
  if (!start || !end) return out;
  if (end < start) return out;
  const diff = Math.round((end.getTime() - start.getTime()) / 86400000);
  for (let i = 0; i <= diff; i++) out.push(addDays(from, i));
  return out;
};
const MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const WEEKDAY_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const dayLabel = (iso: string) => {
  const d = utcDateFromIso(iso) ?? new Date(Date.UTC(2026, 6, 25));
  return {
    weekday: WEEKDAY_SHORT[d.getUTCDay()] ?? "Sun",
    date: `${d.getUTCDate()} ${MONTH_SHORT[d.getUTCMonth()] ?? "Jan"}`,
  };
};
const weekdayOf = (iso: string): number =>
  utcDateFromIso(iso)?.getUTCDay() ?? 0;
const WEEKDAY_LABELS = ["S", "M", "T", "W", "T", "F", "S"];
const WEEKDAY_FULL = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const courseAllowedDate = (course: Course, iso: string): boolean => {
  if (course.fromDate && iso < course.fromDate) return false;
  if (course.toDate && iso > course.toDate) return false;
  return true;
};
const courseAllowedOn = (course: Course, iso: string): boolean => {
  if (!courseAllowedDate(course, iso)) return false;
  const rule = course.allowedWeekdays;
  if (!rule || rule.length === 0) return true;
  return rule.includes(weekdayOf(iso));
};
const courseAllowedSlot = (course: Course, slotIdx: number): boolean => {
  const rule = course.allowedSlots;
  if (!rule || rule.length === 0) return true;
  return rule.includes(slotIdx);
};
// Weekday-aware slot check. Per-weekday override wins over allowedSlots.
const courseAllowedSlotOn = (
  course: Course,
  slotIdx: number,
  iso: string,
): boolean => {
  const perDay = course.allowedSlotsByWeekday?.[weekdayOf(iso)];
  if (perDay !== undefined) return perDay.includes(slotIdx);
  return courseAllowedSlot(course, slotIdx);
};
// Effective allowed slot indices for a course on a given date (weekday-aware).
// Returns null when "all periods" are allowed (no restriction).
const effectiveAllowedSlots = (
  course: Course,
  iso: string,
): number[] | null => {
  const perDay = course.allowedSlotsByWeekday?.[weekdayOf(iso)];
  if (perDay !== undefined) return [...perDay].sort((a, b) => a - b);
  const base = course.allowedSlots;
  if (!base || base.length === 0) return null;
  return [...base].sort((a, b) => a - b);
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

function defaultState(from = isoToday()): State {
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

type LegacyCourse = Course & { allowedPeriods?: number[] };
type SavedClassData = Partial<ClassData> & { courses?: LegacyCourse[] };
type SavedState = Partial<Omit<State, "classes">> & {
  classes?: SavedClassData[];
  courses?: LegacyCourse[];
};

const nonBreakCount = (slots: Slot[]) => slots.filter((slot) => !slot.isBreak).length;
const cleanDurationSlots = (value: unknown, slots: Slot[]): number => {
  const parsed = typeof value === "number" ? value : parseInt(String(value ?? "1"), 10);
  const whole = Number.isFinite(parsed) ? Math.floor(parsed) : 1;
  const max = Math.max(1, nonBreakCount(slots));
  return whole >= 1 && whole <= max ? whole : 1;
};
const isValidIso = (s: unknown): s is string =>
  typeof s === "string" && /^\d{4}-\d{2}-\d{2}$/.test(s);
const cleanCourse = (course: LegacyCourse, slots: Slot[]): Course => {
  const { allowedPeriods, ...rest } = course;
  const rawAllowedSlots = rest.allowedSlots ?? allowedPeriods ?? [];
  const validSlotIdx = (idx: unknown) =>
    typeof idx === "number" &&
    idx >= 0 &&
    idx < slots.length &&
    !slots[idx].isBreak;
  let allowedByWd: Record<number, number[]> | undefined;
  const rawByWd = (rest as { allowedSlotsByWeekday?: unknown }).allowedSlotsByWeekday;
  if (rawByWd && typeof rawByWd === "object") {
    const out: Record<number, number[]> = {};
    for (const [k, v] of Object.entries(rawByWd as Record<string, unknown>)) {
      const wd = parseInt(k, 10);
      if (wd < 0 || wd > 6 || Number.isNaN(wd)) continue;
      if (!Array.isArray(v)) continue;
      out[wd] = (v as unknown[]).filter(validSlotIdx) as number[];
    }
    if (Object.keys(out).length > 0) allowedByWd = out;
  }
  return {
    ...rest,
    id: rest.id || `c${Date.now()}`,
    name: rest.name || "New Course",
    faculty: rest.faculty || "Faculty",
    color: rest.color || COLORS[0],
    durationSlots: cleanDurationSlots(rest.durationSlots, slots),
    weeklyPeriods: Math.max(0, Math.floor(rest.weeklyPeriods ?? 0)),
    totalSessions:
      rest.totalSessions === undefined || rest.totalSessions === null
        ? undefined
        : Math.max(0, Math.floor(rest.totalSessions)) || undefined,
    allowedWeekdays: (rest.allowedWeekdays ?? []).filter((day) => day >= 0 && day <= 6),
    allowedSlots: rawAllowedSlots.filter((idx) => idx >= 0 && idx < slots.length && !slots[idx].isBreak),
    allowedSlotsByWeekday: allowedByWd,
    fromDate: isValidIso(rest.fromDate) ? rest.fromDate : undefined,
    toDate: isValidIso(rest.toDate) ? rest.toDate : undefined,
  };
};
const normalizeStateSnapshot = (snapshot: SavedState): State => {
  const base = defaultState();
  const slots = Array.isArray(snapshot.slots) && snapshot.slots.length > 0 ? snapshot.slots : base.slots;
  const legacyCourses = snapshot.courses?.map((course) => cleanCourse(course, slots));
  const sourceClasses = Array.isArray(snapshot.classes) && snapshot.classes.length > 0 ? snapshot.classes : base.classes;
  const classes: ClassData[] = sourceClasses.map((cls, index) => {
    const savedCourses = cls.courses && cls.courses.length > 0
      ? cls.courses.map((course) => cleanCourse(course, slots))
      : legacyCourses
        ? legacyCourses.map((course) => ({ ...course }))
        : [];
    return {
      id: cls.id || `k${index + 1}`,
      name: cls.name || `Class ${String.fromCharCode(65 + index)}`,
      grid: cls.grid ?? {},
      courses: savedCourses,
    };
  });
  return {
    fromDate: snapshot.fromDate || base.fromDate,
    toDate: snapshot.toDate || base.toDate,
    slots,
    classes,
  };
};

type Tool =
  | { kind: "course"; courseId: string }
  | { kind: "break" }
  | { kind: "blocked" }
  | { kind: "erase" };

function Index() {
  const [state, setState] = useState<State>(() => defaultState(INITIAL_FROM_DATE));
  const [activeClassId, setActiveClassId] = useState("k1");
  const [hydrated, setHydrated] = useState(false);
  const [picker, setPicker] = useState<{ date: string; slotIdx: number } | null>(null);
  const [armedTool, setArmedTool] = useState<Tool | null>(null);
  const [isPainting, setIsPainting] = useState(false);
  const [rulesFor, setRulesFor] = useState<string | null>(null); // courseId in activeClass
  const [bulkWeekdays, setBulkWeekdays] = useState<number[]>([]);
  const [bulkSlots, setBulkSlots] = useState<number[]>([]);
  const [bulkAllClasses, setBulkAllClasses] = useState(false);
  const [autoFillReport, setAutoFillReport] = useState<string>("");
  const gridRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Map slot index → 1-based period number, skipping breaks.
  const periodNumberFor = (slotIdx: number): number =>
    state.slots.slice(0, slotIdx + 1).filter((x) => !x.isBreak).length;
  const periodLabelFor = (slotIdx: number): string =>
    state.slots[slotIdx]?.isBreak ? "Br" : `P${periodNumberFor(slotIdx)}`;

  useEffect(() => {
    setHydrated(true);
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const merged = normalizeStateSnapshot(JSON.parse(raw) as SavedState);
        setState(merged);
        setActiveClassId(merged.classes[0]?.id ?? "");
      } else {
        setState(defaultState());
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
            if (!courseAllowedOn(course, date) || !courseAllowedSlotOn(course, i, date)) {
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
        span = cleanDurationSlots(course?.durationSlots, s.slots);
        }
        for (let k = 0; k < span; k++) {
          const idx = slotIdx + k;
          if (idx >= s.slots.length) break;
          if (s.slots[idx].isBreak) continue; // never write into break slots
          if (tool.kind === "course") {
            const course = cls.courses.find((c) => c.id === tool.courseId);
            if (course && !courseAllowedSlotOn(course, idx, date)) continue;
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
        wdSet.has(weekdayOf(iso)),
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

  const clearTimetable = () => {
    if (!confirm("Clear every course assignment from all classes? Breaks and blocked slots will stay.")) return;
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        const grid: Record<string, Cell> = {};
        Object.entries(cls.grid).forEach(([key, cell]) => {
          if (cell.kind !== "course") grid[key] = cell;
        });
        return { ...cls, grid };
      }),
    }));
    setAutoFillReport("Timetable cleared — course assignments removed.");
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

  // ------------ Auto-populate ------------
  const autoPopulate = (opts: { overwrite: boolean; strictRules?: boolean }) => {
    // ISO year+week key for grouping weekly targets.
    const weekKey = (iso: string) => {
      const d = utcDateFromIso(iso);
      if (!d) return "invalid";
      const day = (d.getUTCDay() + 6) % 7; // Mon=0
      const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() - day + 3));
      const first = new Date(Date.UTC(t.getUTCFullYear(), 0, 4));
      const wk = 1 + Math.round(((t.getTime() - first.getTime()) / 86400000 - 3 + ((first.getUTCDay() + 6) % 7)) / 7);
      return `${t.getUTCFullYear()}-W${wk}`;
    };

    setState((s) => {
      const classes: ClassData[] = s.classes.map((cls) => ({
        ...cls,
        grid: { ...cls.grid },
      }));
      const workingDates = daysBetween(s.fromDate, s.toDate);

      let totalTarget = 0;
      let placedCount = 0;
      const unmet: string[] = [];

      if (opts.overwrite) {
        classes.forEach((cls) => {
          Object.keys(cls.grid).forEach((k) => {
            const cell = cls.grid[k];
            if (cell && cell.kind === "course") delete cls.grid[k];
          });
        });
      }

      const facultyBusy: Record<string, Set<string>> = {};
      classes.forEach((cls) => {
        Object.entries(cls.grid).forEach(([key, cell]) => {
          if (cell.kind !== "course") return;
          const course = cls.courses.find((c) => c.id === cell.courseId);
          if (course) (facultyBusy[key] ??= new Set()).add(course.faculty);
        });
      });

      const allNonBreakStarts = s.slots
        .map((slot, idx) => ({ slot, idx }))
        .filter(({ slot }) => !slot.isBreak)
        .map(({ idx }) => idx);

      const startSlotsFor = (course: Course, date: string): number[] => {
        const eff = effectiveAllowedSlots(course, date);
        const candidates = eff ?? allNonBreakStarts;
        return candidates
          .filter((idx) => idx >= 0 && idx < s.slots.length && !s.slots[idx].isBreak)
          .filter((idx) => opts.strictRules || courseAllowedSlotOn(course, idx, date))
          .sort((a, b) => a - b);
      };

      const spanFitsCourse = (course: Course, start: number, date: string): boolean => {
        for (let i = 0; i < cleanDurationSlots(course.durationSlots, s.slots); i++) {
          const idx = start + i;
          if (idx >= s.slots.length) return false;
          if (s.slots[idx].isBreak) return false;
          if (!opts.strictRules && !courseAllowedSlotOn(course, idx, date)) return false;
        }
        return true;
      };

      const canPlace = (cls: ClassData, course: Course, date: string, start: number): boolean => {
        if (!courseAllowedOn(course, date)) return false;
        if (!startSlotsFor(course, date).includes(start)) return false;
        if (!spanFitsCourse(course, start, date)) return false;
        for (let i = 0; i < cleanDurationSlots(course.durationSlots, s.slots); i++) {
          const idx = start + i;
          const key = `${date}-${idx}`;
          const existing = cls.grid[key];
          if (existing && existing.kind !== "empty") return false;
          if (facultyBusy[key]?.has(course.faculty)) return false;
        }
        return true;
      };

      const countPlacedStarts = (cls: ClassData, course: Course, dateList: string[]) => {
        let placed = 0;
        const perDay: Record<string, number> = {};
        const perSlot: Record<number, number> = {};
        dateList.forEach((d) => {
          perDay[d] = 0;
          s.slots.forEach((_, i) => {
            const cell = cls.grid[`${d}-${i}`];
            if (cell?.kind !== "course" || cell.courseId !== course.id) return;
            const prev = cls.grid[`${d}-${i - 1}`];
            if (!prev || prev.kind !== "course" || prev.courseId !== course.id) {
              placed++;
              perDay[d]++;
              perSlot[i] = (perSlot[i] ?? 0) + 1;
            }
          });
        });
        return { placed, perDay, perSlot };
      };

      type AutoTask = {
        cls: ClassData;
        course: Course;
        remaining: number;
        perDay: Record<string, number>;
        perSlot: Record<number, number>;
        startsByDate: Record<string, number[]>;
        label: string;
      };

      const tasks: AutoTask[] = [];
      const addTask = (cls: ClassData, course: Course, dateList: string[], desiredSessions: number, label: string) => {
        const startsByDate: Record<string, number[]> = {};
        let possibleStarts = 0;
        dateList.forEach((d) => {
          if (!courseAllowedOn(course, d)) return;
          const starts = startSlotsFor(course, d).filter((start) => spanFitsCourse(course, start, d));
          startsByDate[d] = starts;
          possibleStarts += starts.length;
        });
        if (desiredSessions <= 0 && opts.strictRules && possibleStarts > 0) {
          desiredSessions = possibleStarts;
        }
        if (desiredSessions <= 0) return;
        const { placed, perDay, perSlot } = countPlacedStarts(cls, course, dateList);
        const remaining = Math.max(0, desiredSessions - placed);
        totalTarget += remaining;
        if (remaining > 0) {
          tasks.push({ cls, course, remaining, perDay, perSlot, startsByDate, label });
        }
      };

      const weeks = new Map<string, string[]>();
      workingDates.forEach((d) => {
        const key = weekKey(d);
        const week = weeks.get(key);
        if (week) week.push(d);
        else weeks.set(key, [d]);
      });

      classes.forEach((cls) => {
        cls.courses.forEach((course) => {
          if (course.totalSessions && course.totalSessions > 0) {
            addTask(cls, course, workingDates, course.totalSessions, "total");
            return;
          }
          weeks.forEach((weekDates, key) => {
            const desired = course.weeklyPeriods && course.weeklyPeriods > 0 ? course.weeklyPeriods : 0;
            addTask(cls, course, weekDates, desired, key);
          });
        });
      });

      const availableCount = (task: AutoTask) => {
        let count = 0;
        Object.entries(task.startsByDate).forEach(([date, starts]) => {
          starts.forEach((start) => {
            if (canPlace(task.cls, task.course, date, start)) count++;
          });
        });
        return count;
      };

      const classDayLoad = (cls: ClassData, date: string) =>
        s.slots.reduce((sum, _, i) => {
          const cell = cls.grid[`${date}-${i}`];
          return sum + (cell?.kind === "course" ? 1 : 0);
        }, 0);

      let progressed = true;
      while (progressed) {
        progressed = false;
        const availability = new Map<AutoTask, number>();
        tasks.forEach((task) => availability.set(task, availableCount(task)));
        tasks.sort((a, b) => {
          const aAvail = availability.get(a) ?? 0;
          const bAvail = availability.get(b) ?? 0;
          if (aAvail === 0 && bAvail > 0) return 1;
          if (bAvail === 0 && aAvail > 0) return -1;
          const aSlack = aAvail - a.remaining;
          const bSlack = bAvail - b.remaining;
          if (aSlack !== bSlack) return aSlack - bSlack;
          if (b.remaining !== a.remaining) return b.remaining - a.remaining;
          const aRules = (a.course.allowedWeekdays?.length || 7) + Object.values(a.course.allowedSlotsByWeekday ?? {}).flat().length + (a.course.allowedSlots?.length || s.slots.length);
          const bRules = (b.course.allowedWeekdays?.length || 7) + Object.values(b.course.allowedSlotsByWeekday ?? {}).flat().length + (b.course.allowedSlots?.length || s.slots.length);
          return aRules - bRules;
        });

        for (const task of tasks) {
          if (task.remaining <= 0) continue;
          let best: { date: string; slot: number; score: number } | null = null;
          for (const [date, starts] of Object.entries(task.startsByDate)) {
            for (const sIdx of starts) {
              if (!canPlace(task.cls, task.course, date, sIdx)) return;
              const score =
                (task.perDay[date] ?? 0) * 1000000 +
                classDayLoad(task.cls, date) * 10000 +
                (task.perSlot[sIdx] ?? 0) * 100 +
                sIdx;
              if (!best || score < best.score) best = { date, slot: sIdx, score };
            }
          }
          if (!best) continue;

          const span = cleanDurationSlots(task.course.durationSlots, s.slots);
          for (let i = 0; i < span; i++) {
            const key = `${best.date}-${best.slot + i}`;
            task.cls.grid[key] = { kind: "course", courseId: task.course.id };
            (facultyBusy[key] ??= new Set()).add(task.course.faculty);
          }
          task.perDay[best.date] = (task.perDay[best.date] ?? 0) + 1;
          task.perSlot[best.slot] = (task.perSlot[best.slot] ?? 0) + 1;
          task.remaining--;
          placedCount++;
          progressed = true;
        }
      }

      tasks.forEach((task) => {
        if (task.remaining <= 0) return;
        const openNow = availableCount(task);
        unmet.push(
          `${task.cls.name} · ${task.course.name}: ${task.remaining} left${openNow === 0 ? " (no open rule slots)" : ` (${openNow} open rule slots)`}`,
        );
      });

      const visibleCourseCounts = classes.map((cls) => {
        const count = workingDates.reduce((sum, date) => {
          return sum + s.slots.reduce((slotSum, _, slotIdx) => {
            const cell = cls.grid[`${date}-${slotIdx}`];
            return slotSum + (cell?.kind === "course" ? 1 : 0);
          }, 0);
        }, 0);
        return { id: cls.id, name: cls.name, count };
      });
      const activeVisibleCount =
        visibleCourseCounts.find((item) => item.id === activeClassId)?.count ?? 0;
      const firstVisibleClass = visibleCourseCounts.find((item) => item.count > 0);

      // Diagnostic feedback stays on screen instead of blocking with popups.
      queueMicrotask(() => {
        const mode = opts.strictRules ? "Fill by Rules" : opts.overwrite ? "Regenerate" : "Fill Empty";
        const focusFirstFilledRow = () => {
          window.requestAnimationFrame(() => {
            window.requestAnimationFrame(() => {
              const row = gridRef.current?.querySelector<HTMLElement>('[data-filled-row="true"]');
              row?.scrollIntoView({ block: "center", inline: "nearest" });
            });
          });
        };
        if (firstVisibleClass && activeVisibleCount === 0) {
          setActiveClassId(firstVisibleClass.id);
          focusFirstFilledRow();
        } else if (firstVisibleClass) {
          focusFirstFilledRow();
        }
        if (totalTarget === 0) {
          const validDates = workingDates.length;
          const courseCount = classes.reduce((sum, cls) => sum + cls.courses.length, 0);
          const nonBreakCount = s.slots.filter((slot) => !slot.isBreak).length;
          const reason = [
            validDates === 0 ? "date range" : "",
            courseCount === 0 ? "courses" : "",
            nonBreakCount === 0 ? "non-break periods" : "",
          ].filter(Boolean).join(", ");
          setAutoFillReport(`${mode}: nothing to place${reason ? ` — check ${reason}.` : "."}`);
        } else if (placedCount === 0) {
          if (firstVisibleClass) {
            const showingCount = activeVisibleCount > 0 ? activeVisibleCount : firstVisibleClass.count;
            const showingClass = activeVisibleCount > 0
              ? visibleCourseCounts.find((item) => item.id === activeClassId)?.name
              : firstVisibleClass.name;
            setAutoFillReport(
              `${mode}: already filled — showing ${showingCount} course slot${showingCount === 1 ? "" : "s"}${showingClass ? ` in ${showingClass}` : ""}.`,
            );
          } else {
            setAutoFillReport(
              `${mode}: no visible course slots were placed. Check that the selected dates match the course rules and that rule slots are not blocked.`,
            );
          }
        } else if (unmet.length > 0) {
          setAutoFillReport(
            `${mode}: placed ${placedCount} of ${totalTarget}. Remaining: ${unmet.slice(0, 3).join("; ")}`,
          );
        } else {
          setAutoFillReport(`${mode}: placed ${placedCount} of ${totalTarget} planned sessions.`);
        }
      });

      return { ...s, classes };
    });
  };

  // ------------ CSV block upload ------------
  const downloadBlockTemplate = () => {
    const nonBreakCount = state.slots.filter((sl) => !sl.isBreak).length || 8;
    const d0 = state.fromDate || isoToday();
    const sample = [
      "# Auto-block template. Save as .csv and upload via 'Upload blocker CSV'.",
      "# date   = YYYY-MM-DD",
      `# periods = 'all' | comma/range list of period numbers (1..${nonBreakCount}), e.g. '1,2' or '5-${nonBreakCount}'`,
      "# label  = optional text shown in the blocked cell (default: Block)",
      "# scope  = optional 'all' (default) or exact class name; case-insensitive",
      "date,periods,label,scope",
      `${d0},all,Holiday,all`,
      `${addDays(d0, 1)},7-${nonBreakCount},Sports,all`,
      `${addDays(d0, 2)},"1,2",Assembly,${state.classes[0]?.name ?? "Class A"}`,
    ].join("\n");
    const blob = new Blob([sample], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "block-template.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  const parseCsvRow = (line: string): string[] => {
    const out: string[] = [];
    let cur = "";
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQ) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') inQ = false;
        else cur += ch;
      } else {
        if (ch === '"') inQ = true;
        else if (ch === ",") { out.push(cur); cur = ""; }
        else cur += ch;
      }
    }
    out.push(cur);
    return out.map((v) => v.trim());
  };

  const importBlockCsv = async (file: File) => {
    const text = await file.text();
    const rawLines = text.split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !l.startsWith("#"));
    if (rawLines.length === 0) { alert("CSV is empty."); return; }
    const header = parseCsvRow(rawLines[0]).map((c) => c.toLowerCase());
    const dateIdx = header.indexOf("date");
    const periodsIdx = header.indexOf("periods");
    const labelIdx = header.indexOf("label");
    const scopeIdx = header.indexOf("scope");
    if (dateIdx < 0) { alert("CSV missing required 'date' column."); return; }

    const nonBreakIdxs = state.slots.map((sl, i) => ({ sl, i })).filter((x) => !x.sl.isBreak).map((x) => x.i);
    const parsePeriods = (str: string): number[] => {
      const s = (str || "all").trim().toLowerCase();
      if (!s || s === "all" || s === "*") return nonBreakIdxs;
      const nums = new Set<number>();
      s.split(/[;,]/).forEach((p) => {
        const m = p.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
        if (!m) return;
        const a = parseInt(m[1], 10);
        const b = m[2] ? parseInt(m[2], 10) : a;
        for (let n = Math.min(a, b); n <= Math.max(a, b); n++) nums.add(n);
      });
      return [...nums].map((n) => nonBreakIdxs[n - 1]).filter((x): x is number => x !== undefined);
    };

    let applied = 0, skipped = 0, outOfRange = 0, noPeriods = 0;
    const errors: string[] = [];
    setState((s) => {
      const inRange = new Set(daysBetween(s.fromDate, s.toDate));
      const classes: ClassData[] = s.classes.map((cls) => ({ ...cls, grid: { ...cls.grid } }));
      rawLines.slice(1).forEach((line, i) => {
        const cells = parseCsvRow(line);
        const date = cells[dateIdx];
        if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) { errors.push(`Row ${i + 2}: bad date "${date}"`); skipped++; return; }
        const periods = parsePeriods(periodsIdx >= 0 ? cells[periodsIdx] : "all");
        if (periods.length === 0) { errors.push(`Row ${i + 2}: no valid periods parsed from "${cells[periodsIdx] ?? ""}"`); noPeriods++; skipped++; return; }
        if (!inRange.has(date)) { errors.push(`Row ${i + 2}: date ${date} is outside timetable range ${s.fromDate}..${s.toDate}`); outOfRange++; skipped++; return; }
        const label = (labelIdx >= 0 ? cells[labelIdx] : "") || "Block";
        const scope = ((scopeIdx >= 0 ? cells[scopeIdx] : "") || "all").toLowerCase();
        const targets = scope === "all" || scope === "*" || scope === ""
          ? classes
          : classes.filter((c) => c.name.toLowerCase() === scope);
        if (targets.length === 0) { errors.push(`Row ${i + 2}: unknown scope "${cells[scopeIdx]}"`); skipped++; return; }
        targets.forEach((cls) => {
          periods.forEach((slotIdx) => {
            cls.grid[`${date}-${slotIdx}`] = { kind: "blocked", label };
            applied++;
          });
        });
      });
      return { ...s, classes };
    });
    const extras: string[] = [];
    if (outOfRange) extras.push(`${outOfRange} row(s) outside timetable date range (extend From/To to include them).`);
    if (noPeriods) extras.push(`${noPeriods} row(s) had no valid periods.`);
    const msg =
      `Applied ${applied} blocked cells.` +
      (skipped ? ` Skipped ${skipped} row(s).` : "") +
      (extras.length ? `\n\n${extras.join("\n")}` : "") +
      (errors.length ? `\n\n${errors.slice(0, 8).join("\n")}` : "");
    alert(msg);
  };

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
    const wb = XLSXStyle.utils.book_new();
    const hexClean = (h: string) => (h || "").replace("#", "").padStart(6, "0").slice(-6).toUpperCase();
    const textColorFor = (hex: string) => {
      const h = hexClean(hex);
      const r = parseInt(h.slice(0, 2), 16);
      const g = parseInt(h.slice(2, 4), 16);
      const b = parseInt(h.slice(4, 6), 16);
      const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      return lum > 0.6 ? "111111" : "FFFFFF";
    };
    const border = { style: "thin", color: { rgb: "CCCCCC" } } as const;
    const baseBorders = { top: border, bottom: border, left: border, right: border };
    state.classes.forEach((cls) => {
      const data = buildSheet(cls);
      const ws = XLSXStyle.utils.aoa_to_sheet(data);
      const numCols = data[0].length;
      ws["!cols"] = Array.from({ length: numCols }, (_, i) => ({ wch: i === 0 ? 18 : 20 }));
      ws["!rows"] = data.map((_, i) => ({ hpt: i === 0 ? 24 : 32 }));
      for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < numCols; c++) {
          const addr = XLSXStyle.utils.encode_cell({ r, c });
          if (!ws[addr]) ws[addr] = { t: "s", v: "" };
          const cellStyle: Record<string, unknown> = {
            alignment: { horizontal: "center", vertical: "center", wrapText: true },
            border: baseBorders,
            font: { name: "Calibri", sz: 11 },
          };
          if (r === 0 || c === 0) {
            cellStyle.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } };
            cellStyle.fill = { patternType: "solid", fgColor: { rgb: "0D0D0D" } };
          } else {
            const date = dates[r - 1];
            const slotIdx = c - 1;
            const key = `${date}-${slotIdx}`;
            const isConflict = conflicts.has(`${cls.id}:${key}`);
            const sl = state.slots[slotIdx];
            const cell = cls.grid[key];
            if (isConflict) {
              cellStyle.fill = { patternType: "solid", fgColor: { rgb: "FECACA" } };
              cellStyle.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "991B1B" } };
            } else if (sl?.isBreak || cell?.kind === "break") {
              cellStyle.fill = { patternType: "solid", fgColor: { rgb: "FDE68A" } };
              cellStyle.font = { name: "Calibri", sz: 11, italic: true, color: { rgb: "78350F" } };
            } else if (cell?.kind === "blocked") {
              cellStyle.fill = { patternType: "solid", fgColor: { rgb: "374151" } };
              cellStyle.font = { name: "Calibri", sz: 11, color: { rgb: "FFFFFF" } };
            } else if (cell?.kind === "course") {
              const course = cls.courses.find((x) => x.id === cell.courseId);
              const bg = hexClean(course?.color ?? "#DDDDDD");
              cellStyle.fill = { patternType: "solid", fgColor: { rgb: bg } };
              cellStyle.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: textColorFor(bg) } };
            }
          }
          (ws[addr] as { s?: unknown }).s = cellStyle;
        }
      }
      ws["!freeze"] = { xSplit: 1, ySplit: 1 };
      XLSXStyle.utils.book_append_sheet(wb, ws, cls.name.slice(0, 31) || "Class");
    });
    XLSXStyle.writeFile(wb, `timetable_${state.fromDate}_to_${state.toDate}.xlsx`);
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
  const exportASC = () => {
    const pad = (s: string, n: number) => {
      const t = s.length > n ? s.slice(0, n - 1) + "…" : s;
      return t + " ".repeat(Math.max(0, n - t.length));
    };
    const parts: string[] = [];
    parts.push(`Timetable  ${state.fromDate}  to  ${state.toDate}`);
    parts.push("=".repeat(72));
    state.classes.forEach((cls) => {
      const rows = buildSheet(cls);
      const colWidths = rows[0].map((_, ci) =>
        Math.min(22, Math.max(...rows.map((r) => String(r[ci] ?? "").length))),
      );
      const sep = "+" + colWidths.map((w) => "-".repeat(w + 2)).join("+") + "+";
      const fmt = (r: string[]) =>
        "| " + r.map((v, ci) => pad(String(v ?? ""), colWidths[ci])).join(" | ") + " |";
      parts.push("");
      parts.push(`Class: ${cls.name}`);
      parts.push(sep);
      parts.push(fmt(rows[0]));
      parts.push(sep);
      rows.slice(1).forEach((r) => parts.push(fmt(r)));
      parts.push(sep);
    });
    const blob = new Blob([parts.join("\n")], { type: "text/plain;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `timetable_${state.fromDate}_to_${state.toDate}.asc`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Save / Load .aadhi file (full app state snapshot)
  const saveAadhi = () => {
    const payload = {
      app: "timetable-maker",
      version: 5,
      savedAt: new Date().toISOString(),
      state,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const stamp = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `timetable_${stamp}.aadhi`;
    a.click();
    URL.revokeObjectURL(url);
  };
  const loadAadhi = async (file: File) => {
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      const loaded: SavedState | undefined = parsed?.state ?? parsed;
      if (!loaded || !Array.isArray(loaded.classes) || !Array.isArray(loaded.slots)) {
        alert("This doesn't look like a valid .aadhi file.");
        return;
      }
      if (!confirm("Load this file? Your current timetable will be replaced.")) return;
      const normalized = normalizeStateSnapshot(loaded);
      setState(normalized);
      setActiveClassId(normalized.classes[0]?.id ?? "");
      setAutoFillReport("Loaded .aadhi file. Course span values above the available periods were corrected to 1.");
    } catch {
      alert("Could not read this .aadhi file.");
    }
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
  const activeVisibleCourseSlots = useMemo(() => {
    if (!activeClass) return 0;
    return dates.reduce((sum, date) => {
      return sum + state.slots.reduce((slotSum, _, slotIdx) => {
        const cell = activeClass.grid[`${date}-${slotIdx}`];
        return slotSum + (cell?.kind === "course" ? 1 : 0);
      }, 0);
    }, 0);
  }, [activeClass, dates, state.slots]);

  // Planned vs placed sessions for the current date range. A "session" is one
  // course start (multi-slot durations count as one). Planned = sum of each
  // course's weeklyPeriods × number of ISO weeks covered by the range.
  const sessionStats = useMemo(() => {
    const isoWeekKey = (iso: string) => {
      const d = utcDateFromIso(iso);
      if (!d) return "invalid";
      const day = (d.getUTCDay() + 6) % 7;
      const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() - day + 3));
      const first = new Date(Date.UTC(t.getUTCFullYear(), 0, 4));
      const wk = 1 + Math.round(((t.getTime() - first.getTime()) / 86400000 - 3 + ((first.getUTCDay() + 6) % 7)) / 7);
      return `${t.getUTCFullYear()}-W${wk}`;
    };
    const weekCount = new Set(dates.map(isoWeekKey)).size;
    const countPlaced = (cls: ClassData) => {
      let n = 0;
      dates.forEach((d) => {
        state.slots.forEach((_, i) => {
          const cell = cls.grid[`${d}-${i}`];
          if (cell?.kind !== "course") return;
          const prev = cls.grid[`${d}-${i - 1}`];
          if (!prev || prev.kind !== "course" || prev.courseId !== cell.courseId) n++;
        });
      });
      return n;
    };
    const planFor = (cls: ClassData) =>
      cls.courses.reduce((sum, c) => {
        if (c.totalSessions && c.totalSessions > 0) return sum + c.totalSessions;
        return sum + Math.max(0, c.weeklyPeriods ?? 0) * weekCount;
      }, 0);
    const activePlanned = activeClass ? planFor(activeClass) : 0;
    const activePlaced = activeClass ? countPlaced(activeClass) : 0;
    const totalPlanned = state.classes.reduce((s, c) => s + planFor(c), 0);
    const totalPlaced = state.classes.reduce((s, c) => s + countPlaced(c), 0);
    return {
      activePlanned,
      activePlaced,
      activeRemaining: Math.max(0, activePlanned - activePlaced),
      totalPlanned,
      totalPlaced,
      totalRemaining: Math.max(0, totalPlanned - totalPlaced),
    };
  }, [activeClass, dates, state.slots, state.classes]);

  // Per-course assigned session counts for the active class (session starts only).
  const coursePlacementCounts = useMemo(() => {
    const map = new Map<string, number>();
    if (!activeClass) return map;
    dates.forEach((d) => {
      state.slots.forEach((_, i) => {
        const cell = activeClass.grid[`${d}-${i}`];
        if (cell?.kind !== "course") return;
        const prev = activeClass.grid[`${d}-${i - 1}`];
        if (!prev || prev.kind !== "course" || prev.courseId !== cell.courseId) {
          map.set(cell.courseId, (map.get(cell.courseId) ?? 0) + 1);
        }
      });
    });
    return map;
  }, [activeClass, dates, state.slots]);

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
                    <div className="grid grid-cols-[minmax(0,1fr)_auto_auto] gap-2 px-3 py-2">
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
                          max={Math.max(1, nonBreakCount(state.slots))}
                          value={c.durationSlots}
                          onChange={(e) =>
                            updateCourse(c.id, {
                              durationSlots: cleanDurationSlots(e.target.value, state.slots),
                            })
                          }
                          className="w-10 border border-[#0d0d0d]/20 bg-white px-1 py-0.5 text-center text-xs"
                        />
                        <span title="Consecutive periods per session">span</span>
                      </label>
                      <label
                        title="Sessions per week (auto-fill target)"
                        className="flex items-center gap-1 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        <input
                          type="number"
                          min={0}
                          value={c.weeklyPeriods ?? 0}
                          onChange={(e) =>
                            updateCourse(c.id, {
                              weeklyPeriods: Math.max(0, parseInt(e.target.value || "0", 10)),
                            })
                          }
                          className="w-10 border border-[#0d0d0d]/20 bg-white px-1 py-0.5 text-center text-xs"
                        />
                        <span>/wk</span>
                      </label>
                      <label
                        title="Total sessions across the whole date range. Overrides /wk when set."
                        className="flex items-center gap-1 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        <input
                          type="number"
                          min={0}
                          value={c.totalSessions ?? 0}
                          onChange={(e) => {
                            const n = Math.max(0, parseInt(e.target.value || "0", 10));
                            updateCourse(c.id, { totalSessions: n > 0 ? n : undefined });
                          }}
                          className="w-12 border border-[#0d0d0d]/20 bg-white px-1 py-0.5 text-center text-xs"
                        />
                        <span>
                          {(c.totalSessions ?? 0) > 0
                            ? `${coursePlacementCounts.get(c.id) ?? 0} / ${c.totalSessions}`
                            : `${coursePlacementCounts.get(c.id) ?? 0} placed`}
                          {" "}· total
                        </span>
                      </label>
                    </div>
                    <div className="grid grid-cols-2 gap-2 px-3 pb-2">
                      <label
                        className="flex flex-col gap-0.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        <span>From</span>
                        <input
                          type="date"
                          value={c.fromDate ?? ""}
                          onChange={(e) =>
                            updateCourse(c.id, {
                              fromDate: e.target.value || undefined,
                            })
                          }
                          className="w-full border border-[#0d0d0d]/20 bg-white px-1 py-0.5 text-[10px] outline-none focus:border-[#0d0d0d]"
                        />
                      </label>
                      <label
                        className="flex flex-col gap-0.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        <span>To</span>
                        <input
                          type="date"
                          value={c.toDate ?? ""}
                          onChange={(e) =>
                            updateCourse(c.id, {
                              toDate: e.target.value || undefined,
                            })
                          }
                          className="w-full border border-[#0d0d0d]/20 bg-white px-1 py-0.5 text-[10px] outline-none focus:border-[#0d0d0d]"
                        />
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
                        {c.fromDate || c.toDate
                          ? `${c.fromDate ?? "start"} → ${c.toDate ?? "end"}`
                          : c.allowedWeekdays && c.allowedWeekdays.length > 0
                            ? c.allowedWeekdays.map((w) => WEEKDAY_FULL[w]).join(" ")
                            : "All days"}
                        {" · "}
                        {c.allowedSlots && c.allowedSlots.length > 0
                          ? c.allowedSlots.map((i) => `P${periodNumberFor(i)}`).join(" ")
                          : "All periods"}
                        {c.allowedSlotsByWeekday && Object.keys(c.allowedSlotsByWeekday).length > 0
                          ? " · per-day"
                          : ""}
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

            {/* Auto-fill + CSV blocker */}
            <section className="border border-dashed border-[#0d0d0d]/50 bg-[#f5f3ee] p-4">
              <h3
                className="mb-2 text-[11px] font-bold uppercase tracking-widest"
                style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
              >
                Auto-fill & Blocker
              </h3>
              <p className="mb-3 text-[11px] text-[#2d2d2d]/70">
                Auto-fill packs each course into the week using its rules. If <b>/wk</b> is 0,
                Fill by Rules uses every allowed weekday and period opportunity.
              </p>
              <div className="mb-3 grid grid-cols-2 gap-1">
                <button
                  onClick={() => autoPopulate({ overwrite: false })}
                  className="border-2 border-[#0d0d0d] bg-white px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
                >
                  Fill Empty
                </button>
                <button
                  onClick={() => {
                    if (confirm("Clear all courses and re-generate?")) autoPopulate({ overwrite: true });
                  }}
                  className="border-2 border-[#0d0d0d] bg-[#0d0d0d] px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider text-[#f5f3ee] hover:opacity-90"
                >
                  Regenerate
                </button>
              </div>
              <div className="mb-3 grid grid-cols-2 gap-1">
                <button
                  onClick={() => autoPopulate({ overwrite: false, strictRules: true })}
                  className="border-2 border-[#0d0d0d] bg-amber-200 px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-amber-300"
                  title="Places courses in their rule periods. If /wk is 0, uses every allowed period opportunity."
                >
                  Fill by Rules
                </button>
                <button
                  onClick={() => {
                    if (confirm("Clear all courses and fill only by course rules?")) {
                      autoPopulate({ overwrite: true, strictRules: true });
                    }
                  }}
                  className="border-2 border-[#0d0d0d] bg-amber-300 px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-amber-400"
                  title="Clears current course entries first, then fills only in the selected rule periods."
                >
                  Regen Rules
                </button>
              </div>
              {autoFillReport && (
                <div className="mb-3 border border-[#0d0d0d]/30 bg-white px-3 py-2 text-[11px] font-bold text-[#2d2d2d]">
                  {autoFillReport}
                </div>
              )}

              <div className="border-t border-dashed border-[#0d0d0d]/30 pt-3">
                <div className="mb-1 text-[10px] font-bold uppercase tracking-wider text-[#2d2d2d]/60">
                  Block dates via CSV
                </div>
                <div className="flex flex-wrap gap-1">
                  <button
                    onClick={downloadBlockTemplate}
                    className="border border-[#0d0d0d]/60 bg-white px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
                  >
                    Template
                  </button>
                  <label className="cursor-pointer border border-[#0d0d0d] bg-[#0d0d0d] px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider text-[#f5f3ee] hover:opacity-90">
                    Upload CSV
                    <input
                      type="file"
                      accept=".csv,text/csv"
                      className="hidden"
                      onChange={(e) => {
                        const f = e.target.files?.[0];
                        if (f) importBlockCsv(f);
                        e.target.value = "";
                      }}
                    />
                  </label>
                </div>
                <p className="mt-2 text-[10px] text-[#2d2d2d]/60">
                  Columns: <code>date, periods, label, scope</code>. Periods are 1-based
                  over non-break slots (e.g. <code>1,2</code> or <code>5-8</code> or <code>all</code>).
                </p>
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
              <input
                ref={fileInputRef}
                type="file"
                accept=".aadhi,application/json"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) loadAadhi(f);
                  e.target.value = "";
                }}
              />
              <button
                onClick={clearTimetable}
                className="border-2 border-red-700 bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider text-red-700 transition-transform hover:bg-red-50 active:translate-y-0.5"
                title="Clear all course assignments"
              >
                Clear
              </button>
              <button
                onClick={() => fileInputRef.current?.click()}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5"
                title="Load a .aadhi file"
              >
                Load
              </button>
              <button
                onClick={saveAadhi}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5"
                title="Save as .aadhi file"
              >
                Save
              </button>
              <button
                onClick={exportCSV}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5"
              >
                CSV
              </button>
              <button
                onClick={exportASC}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5"
                title="Export as ASCII text (.asc)"
              >
                ASC
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
              {dates.length} day{dates.length === 1 ? "" : "s"} · {state.slots.length} slots · {activeVisibleCourseSlots} filled
            </span>
            <span className="text-[#2d2d2d]/20">·</span>
            <span
              className="flex flex-wrap items-center gap-2 text-[11px]"
              style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
            >
              <span
                className={
                  "border-2 px-2 py-0.5 font-bold " +
                  (sessionStats.activeRemaining === 0
                    ? "border-emerald-700 bg-emerald-50 text-emerald-800"
                    : "border-[#0d0d0d] bg-white text-[#0d0d0d]")
                }
                title="This class · placed / planned (remaining)"
              >
                {activeClass?.name ?? "Class"}: {sessionStats.activePlaced}/{sessionStats.activePlanned}
                <span className="ml-1 text-[#2d2d2d]/60">
                  · {sessionStats.activeRemaining} left
                </span>
              </span>
              <span
                className={
                  "border-2 px-2 py-0.5 font-bold " +
                  (sessionStats.totalRemaining === 0
                    ? "border-emerald-700 bg-emerald-50 text-emerald-800"
                    : "border-[#0d0d0d]/60 bg-[#f5f3ee] text-[#0d0d0d]")
                }
                title="All classes · placed / planned (remaining)"
              >
                All: {sessionStats.totalPlaced}/{sessionStats.totalPlanned}
                <span className="ml-1 text-[#2d2d2d]/60">
                  · {sessionStats.totalRemaining} left
                </span>
              </span>
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
                    const rowHasCourse = activeClass
                      ? state.slots.some((_, slotIdx) => activeClass.grid[`${date}-${slotIdx}`]?.kind === "course")
                      : false;
                    return (
                      <tr key={date} data-filled-row={rowHasCourse ? "true" : undefined}>
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
                    courseAllowedSlotOn(c, picker.slotIdx, picker.date);
                  // Count placed sessions of this course in the current date range
                  let placedForCourse = 0;
                  if (activeClass) {
                    dates.forEach((d) => {
                      state.slots.forEach((_, i) => {
                        const cell = activeClass.grid[`${d}-${i}`];
                        if (cell?.kind === "course" && cell.courseId === c.id) {
                          const prev = activeClass.grid[`${d}-${i - 1}`];
                          if (!prev || prev.kind !== "course" || prev.courseId !== c.id)
                            placedForCourse++;
                        }
                      });
                    });
                  }
                  const dateRange = c.fromDate || c.toDate ? `${c.fromDate ?? "start"} → ${c.toDate ?? "end"}` : null;
                  const ruleLabel =
                    dateRange ||
                    (c.allowedWeekdays && c.allowedWeekdays.length > 0) ||
                    (c.allowedSlots && c.allowedSlots.length > 0)
                      ? [
                          dateRange ??
                            (c.allowedWeekdays && c.allowedWeekdays.length > 0
                              ? c.allowedWeekdays.map((w) => WEEKDAY_FULL[w]).join(",")
                              : "any day"),
                          c.allowedSlots && c.allowedSlots.length > 0
                            ? c.allowedSlots
                                .map((i) => `P${periodNumberFor(i)}`)
                                .join(",")
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
                        <span
                          className="mt-0.5 block text-[10px] font-bold uppercase tracking-wider text-[#2d2d2d]/70"
                          style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                        >
                          {c.totalSessions && c.totalSessions > 0
                            ? `${placedForCourse} / ${c.totalSessions} sessions`
                            : `${placedForCourse} placed · no total set`}
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
                      Active date range
                    </span>
                    <button
                      onClick={() => updateCourse(course.id, { fromDate: undefined, toDate: undefined })}
                      className="text-[10px] uppercase tracking-wider text-[#2d2d2d]/50 hover:text-[#0d0d0d]"
                    >
                      All dates
                    </button>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <label className="flex flex-col gap-0.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60">
                      <span>From</span>
                      <input
                        type="date"
                        value={course.fromDate ?? ""}
                        onChange={(e) => updateCourse(course.id, { fromDate: e.target.value || undefined })}
                        className="border border-[#0d0d0d]/20 bg-white px-2 py-1 text-xs outline-none focus:border-[#0d0d0d]"
                      />
                    </label>
                    <label className="flex flex-col gap-0.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60">
                      <span>To</span>
                      <input
                        type="date"
                        value={course.toDate ?? ""}
                        onChange={(e) => updateCourse(course.id, { toDate: e.target.value || undefined })}
                        className="border border-[#0d0d0d]/20 bg-white px-2 py-1 text-xs outline-none focus:border-[#0d0d0d]"
                      />
                    </label>
                  </div>
                  <p className="mt-1 text-[10px] text-[#2d2d2d]/50">
                    Restrict this course to a specific date window. Leave blank to use the full timetable range.
                  </p>
                </div>

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
                            const base = wdAll ? [] : [...wdRule];
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
                            const base = slotAll ? [] : [...slotRule];
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

                <div>
                  <div className="mb-2 flex items-center justify-between">
                    <span className="text-[11px] font-bold uppercase tracking-widest text-[#2d2d2d]/70">
                      Per-weekday periods
                    </span>
                    <button
                      onClick={() =>
                        updateCourse(course.id, { allowedSlotsByWeekday: undefined })
                      }
                      className="text-[10px] uppercase tracking-wider text-[#2d2d2d]/50 hover:text-[#0d0d0d]"
                    >
                      Reset all
                    </button>
                  </div>
                  <p className="mb-2 text-[10px] text-[#2d2d2d]/50">
                    Optional. Override the default periods above for specific
                    weekdays (e.g. Mon P1–P2, Thu P5–P6). Unset weekdays fall
                    back to the default.
                  </p>
                  <div className="space-y-1.5">
                    {WEEKDAY_LABELS.map((_lbl, wd) => {
                      if (!wdAll && !wdRule.includes(wd)) return null;
                      const byWd = course.allowedSlotsByWeekday ?? {};
                      const override = byWd[wd];
                      const isCustom = override !== undefined;
                      const activeSet = isCustom
                        ? new Set(override)
                        : new Set(slotAll ? nonBreakIdxs.map((x) => x.i) : slotRule);
                      return (
                        <div
                          key={wd}
                          className="border border-[#0d0d0d]/15 bg-white px-2 py-1.5"
                        >
                          <div className="mb-1 flex items-center justify-between">
                            <span
                              className="text-[10px] font-bold uppercase tracking-wider"
                              style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                            >
                              {WEEKDAY_FULL[wd]}
                              {!isCustom && (
                                <span className="ml-1 text-[9px] font-normal text-[#2d2d2d]/40">
                                  · default
                                </span>
                              )}
                            </span>
                            {isCustom && (
                              <button
                                onClick={() => {
                                  const next = { ...byWd };
                                  delete next[wd];
                                  updateCourse(course.id, {
                                    allowedSlotsByWeekday:
                                      Object.keys(next).length > 0 ? next : undefined,
                                  });
                                }}
                                className="text-[9px] uppercase tracking-wider text-[#2d2d2d]/50 hover:text-[#0d0d0d]"
                              >
                                Reset
                              </button>
                            )}
                          </div>
                          <div className="flex flex-wrap gap-1">
                            {nonBreakIdxs.map(({ i }) => {
                              const periodNum = state.slots
                                .slice(0, i + 1)
                                .filter((x) => !x.isBreak).length;
                              const on = activeSet.has(i);
                              return (
                                <button
                                  key={i}
                                  onClick={() => {
                                    const baseArr = isCustom
                                      ? [...override!]
                                      : slotAll
                                        ? nonBreakIdxs.map((x) => x.i)
                                        : [...slotRule];
                                    const nextArr = baseArr.includes(i)
                                      ? baseArr.filter((x) => x !== i)
                                      : [...baseArr, i].sort((a, b) => a - b);
                                    const nextByWd = { ...byWd, [wd]: nextArr };
                                    updateCourse(course.id, {
                                      allowedSlotsByWeekday: nextByWd,
                                    });
                                  }}
                                  className={
                                    "min-w-[2.25rem] border px-1.5 py-0.5 text-[10px] font-bold " +
                                    (on
                                      ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee]"
                                      : isCustom
                                        ? "border-[#0d0d0d]/20 bg-white text-[#2d2d2d]/40 hover:border-[#0d0d0d]/50"
                                        : "border-dashed border-[#0d0d0d]/20 bg-white text-[#2d2d2d]/30 hover:border-[#0d0d0d]/40")
                                  }
                                  style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                                >
                                  P{periodNum}
                                </button>
                              );
                            })}
                          </div>
                        </div>
                      );
                    })}
                  </div>
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
