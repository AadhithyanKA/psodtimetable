import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import XLSXStyle from "xlsx-js-style";
import JSZip from "jszip";

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
  | { kind: "course"; courseId: string; locked?: boolean };

type Course = {
  id: string;
  name: string;
  faculty: string;
  classroom?: string;
  color: string;
  durationSlots: number;
  minDurationSlots?: number;
  maxDurationSlots?: number;
  stretchRuleEnabled?: boolean;
  // Target number of sessions per week (used by auto-fill). 0 = don't auto-fill.
  weeklyPeriods?: number;
  // LTPC structure. When any of L/T/P > 0 the weekly session target is
  // derived as L + T + P and overrides `weeklyPeriods` in auto-fill and
  // planned counters. C is informational.
  lectureHours?: number;
  tutorialHours?: number;
  practicalHours?: number;
  credits?: number;
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
  // When true, the course is skipped by Fill by Rules / auto-populate.
  disabled?: boolean;
  // When true, the course is a common/combined subject across classes.
  common?: boolean;
};
type ClassData = {
  id: string;
  name: string;
  grid: Record<string, Cell>;
  courses: Course[];
  // Optional department assignment for grouping in reports.
  department?: string;
  // Optional per-class date range overrides. Undefined = fall back to state.fromDate/state.toDate.
  fromDate?: string;
  toDate?: string;
};

const classFromDate = (cls: Pick<ClassData, "fromDate">, s: { fromDate: string }) =>
  cls.fromDate && isValidIso(cls.fromDate) ? cls.fromDate : s.fromDate;
const classToDate = (cls: Pick<ClassData, "toDate">, s: { toDate: string }) =>
  cls.toDate && isValidIso(cls.toDate) ? cls.toDate : s.toDate;
const classDatesFor = (
  cls: Pick<ClassData, "fromDate" | "toDate">,
  s: { fromDate: string; toDate: string },
) => daysBetween(classFromDate(cls, s), classToDate(cls, s));
const stateForClass = <T extends { fromDate: string; toDate: string }>(
  cls: Pick<ClassData, "fromDate" | "toDate">,
  s: T,
): T => ({ ...s, fromDate: classFromDate(cls, s), toDate: classToDate(cls, s) });

const weekIndexForDate = (
  cls: Pick<ClassData, "fromDate">,
  s: { fromDate: string },
  date: string,
): number => {
  const start = utcDateFromIso(classFromDate(cls, s));
  const cur = utcDateFromIso(date);
  if (!start || !cur) return 1;

  const startDay = start.getUTCDay();
  const startOffset = startDay === 0 ? 6 : startDay - 1;
  const startMonday = new Date(start.getTime());
  startMonday.setUTCDate(startMonday.getUTCDate() - startOffset);
  startMonday.setUTCHours(0, 0, 0, 0);

  const curDay = cur.getUTCDay();
  const curOffset = curDay === 0 ? 6 : curDay - 1;
  const curMonday = new Date(cur.getTime());
  curMonday.setUTCDate(curMonday.getUTCDate() - curOffset);
  curMonday.setUTCHours(0, 0, 0, 0);

  return Math.floor((curMonday.getTime() - startMonday.getTime()) / (7 * 86400000)) + 1;
};
type Slot = { start: string; end: string; isBreak?: boolean }; // 24h "HH:MM"
type State = {
  fromDate: string; // YYYY-MM-DD
  toDate: string;
  slots: Slot[];
  classes: ClassData[];
  // When true, the timetable is locked: no edits, no auto-fill, and
  // conflict/override warnings are suppressed so overrides become permanent.
  frozen?: boolean;
  ignoredConflicts?: string[]; // keys of conflicts ignored by the user
};

const COLORS = [
  "#fdba74", // Peach
  "#fcd34d", // Soft Gold
  "#86efac", // Light Emerald
  "#67e8f9", // Pale Cyan
  "#93c5fd", // Sky Blue
  "#c4b5fd", // Soft Lavender
  "#f9a8d4", // Pink Rose
  "#a7f3d0", // Light Mint
  "#fca5a5", // Coral Red
  "#fed7aa", // Apricot
  "#fef08a", // Pale Yellow
  "#bbf7d0", // Soft Mint
  "#99f6e4", // Light Aquamarine
  "#bfdbfe", // Soft Sky
  "#e9d5ff", // Soft Lilac
  "#fbcfe8", // Pale Rose
  "#cbd5e1", // Pale Slate
  "#ddd6fe", // Lavender Mist
  "#fda4af", // Soft Salmon
  "#e2e8f0", // Soft Gray
];

const hslToHex = (hslStr: string): string => {
  const match = hslStr.match(/hsl\((\d+),\s*(\d+)%,\s*(\d+)%\)/);
  if (!match) return "#cccccc";
  const h = parseInt(match[1], 10) / 360;
  const s = parseInt(match[2], 10) / 100;
  const l = parseInt(match[3], 10) / 100;
  let r, g, b;
  if (s === 0) {
    r = g = b = l;
  } else {
    const hue2rgb = (p: number, q: number, t: number) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue2rgb(p, q, h + 1 / 3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1 / 3);
  }
  const toHex = (x: number) =>
    Math.round(x * 255)
      .toString(16)
      .padStart(2, "0");
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
};

const getUniqueCourseColor = (existingCourses: Course[]): string => {
  const used = new Set((existingCourses || []).map((c) => c.color.toLowerCase()));
  for (const color of COLORS) {
    if (!used.has(color.toLowerCase())) {
      return color;
    }
  }
  // Golden Angle (137.5 deg) distribution for distinct pastel HSL colors
  const hue = ((existingCourses || []).length * 137.5) % 360;
  return `hsl(${Math.round(hue)}, 75%, 85%)`;
};
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
const MONTH_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];
const WEEKDAY_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const dayLabel = (iso: string) => {
  const d = utcDateFromIso(iso) ?? new Date(Date.UTC(2026, 6, 25));
  return {
    weekday: WEEKDAY_SHORT[d.getUTCDay()] ?? "Sun",
    date: `${d.getUTCDate()} ${MONTH_SHORT[d.getUTCMonth()] ?? "Jan"}`,
  };
};
const weekdayOf = (iso: string): number => utcDateFromIso(iso)?.getUTCDay() ?? 0;
const WEEKDAY_LABELS = ["S", "M", "T", "W", "T", "F", "S"];
const WEEKDAY_FULL = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
// Faculty helpers. A course's `faculty` field may hold one name or several
// names separated by commas / semicolons / slashes to model co-taught
// sessions. All conflict checks compare the SET of faculty names — two
// courses conflict if they share ANY teacher.
const getFaculties = (course: Pick<Course, "faculty">): string[] => {
  const raw = (course.faculty || "").trim();
  if (!raw) return [];
  return raw
    .split(/[,;/]+/)
    .map((s) => s.trim())
    .filter(Boolean);
};
const facultyKey = (course: Pick<Course, "faculty">): string =>
  getFaculties(course).slice().sort().join("|");
const sharesFaculty = (a: Pick<Course, "faculty">, b: Pick<Course, "faculty">): boolean => {
  const set = new Set(getFaculties(a));
  return getFaculties(b).some((f) => set.has(f));
};
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
const courseAllowedSlotOn = (course: Course, slotIdx: number, iso: string): boolean => {
  const perDay = course.allowedSlotsByWeekday?.[weekdayOf(iso)];
  if (perDay !== undefined) return perDay.includes(slotIdx);
  return courseAllowedSlot(course, slotIdx);
};
// Effective allowed slot indices for a course on a given date (weekday-aware).
// Returns null when "all periods" are allowed (no restriction).
const effectiveAllowedSlots = (course: Course, iso: string): number[] | null => {
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
    {
      id: "c1",
      name: "Mathematics",
      faculty: "Dr. Smith",
      color: COLORS[0],
      durationSlots: 1,
      weeklyPeriods: 4,
    },
    {
      id: "c2",
      name: "Physics",
      faculty: "Dr. Jones",
      color: COLORS[2],
      durationSlots: 1,
      weeklyPeriods: 3,
    },
    {
      id: "c3",
      name: "Chemistry",
      faculty: "Dr. Patel",
      color: COLORS[4],
      durationSlots: 1,
      weeklyPeriods: 3,
    },
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
    ignoredConflicts: [],
  };
}

type LegacyCourse = Course & { allowedPeriods?: number[] };
type SavedClassData = Partial<ClassData> & { courses?: LegacyCourse[] };
type SavedState = Partial<Omit<State, "classes">> & {
  classes?: SavedClassData[];
  courses?: LegacyCourse[];
};

const nonBreakCount = (slots: Slot[]) => slots.filter((slot) => !slot.isBreak).length;
// LTPC-derived weekly session target. Each unit of L, T, or P contributes
// 1 scheduled period/week; multi-period practical blocks are counted by their
// occupied periods so export totals match the visible timetable.
// Falls back to weeklyPeriods when no LTPC values are set.
const courseWeeklyTarget = (course: Course): number => {
  const L = Math.max(0, course.lectureHours ?? 0);
  const T = Math.max(0, course.tutorialHours ?? 0);
  const P = Math.max(0, course.practicalHours ?? 0);
  if (L + T + P > 0) return L + T + P;
  return Math.max(0, course.weeklyPeriods ?? 0);
};
// Number of calendar weeks covered by [from, to] inclusive (ceil days/7).
const weeksInRange = (from?: string, to?: string): number => {
  if (!from || !to) return 0;
  const d1 = Date.parse(`${from}T00:00:00Z`);
  const d2 = Date.parse(`${to}T00:00:00Z`);
  if (!Number.isFinite(d1) || !Number.isFinite(d2) || d2 < d1) return 0;
  const days = Math.floor((d2 - d1) / 86400000) + 1;
  return Math.max(1, Math.ceil(days / 7));
};
const effectiveCourseRange = (
  course: Course,
  state: { fromDate: string; toDate: string },
): { from: string; to: string } | null => {
  const from = [state.fromDate, course.fromDate].filter(isValidIso).sort()[1] ?? state.fromDate;
  const to = [state.toDate, course.toDate].filter(isValidIso).sort()[0] ?? state.toDate;
  if (!isValidIso(from) || !isValidIso(to) || to < from) return null;
  return { from, to };
};
// Weeks a course is scheduled over: course from/to is intersected with the
// timetable range so totals are based on the actual timetable being designed.
const courseSemesterWeeks = (
  course: Course,
  state: { fromDate: string; toDate: string },
): number => {
  const range = effectiveCourseRange(course, state);
  if (!range) return 0;
  return weeksInRange(range.from, range.to);
};
const courseCycleForDate = (
  course: Course,
  state: { fromDate: string; toDate: string },
  iso: string,
): number | null => {
  const start = utcDateFromIso(state.fromDate);
  const current = utcDateFromIso(iso);
  if (!start || !current || current < start) return null;

  const startDay = start.getUTCDay();
  const startOffset = startDay === 0 ? 6 : startDay - 1;
  const startMonday = new Date(start.getTime());
  startMonday.setUTCDate(startMonday.getUTCDate() - startOffset);
  startMonday.setUTCHours(0, 0, 0, 0);

  const curDay = current.getUTCDay();
  const curOffset = curDay === 0 ? 6 : curDay - 1;
  const curMonday = new Date(current.getTime());
  curMonday.setUTCDate(curMonday.getUTCDate() - curOffset);
  curMonday.setUTCHours(0, 0, 0, 0);

  return Math.floor((curMonday.getTime() - startMonday.getTime()) / (7 * 86400000)) + 1;
};
// Total sessions across the course's active range. Explicit totalSessions
// overrides; otherwise LTPC → (L+T+P) × effective timetable weeks.
// Example over 15 weeks: L=1,T=0,P=4 → 5 × 15 = 75 (15 theory + 60 practical periods).
const courseTotalTarget = (course: Course, weeks: number): number => {
  if (course.totalSessions && course.totalSessions > 0) return course.totalSessions;
  const L = Math.max(0, course.lectureHours ?? 0);
  const T = Math.max(0, course.tutorialHours ?? 0);
  const P = Math.max(0, course.practicalHours ?? 0);
  if (L + T + P > 0) return (L + T + P) * Math.max(0, weeks);
  const weekly = Math.max(0, course.weeklyPeriods ?? 0);
  if (weekly > 0 && weeks > 0) return weekly * weeks;
  return 0;
};
const splitTotalAcrossWeeks = (total: number, weeks: number): number[] => {
  const cleanTotal = Math.max(0, Math.floor(total));
  const cleanWeeks = Math.max(0, Math.floor(weeks));
  if (cleanTotal <= 0 || cleanWeeks <= 0) return [];
  const base = Math.floor(cleanTotal / cleanWeeks);
  const remainder = cleanTotal % cleanWeeks;
  return Array.from({ length: cleanWeeks }, (_, index) => base + (index < remainder ? 1 : 0));
};
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
  const validSlotIdx = (idx: unknown) => {
    const num = parseInt(String(idx), 10);
    return !Number.isNaN(num) && num >= 0 && num < slots.length && !slots[num].isBreak;
  };
  let allowedByWd: Record<number, number[]> | undefined;
  const rawByWd = (rest as { allowedSlotsByWeekday?: unknown }).allowedSlotsByWeekday;
  if (rawByWd && typeof rawByWd === "object") {
    const out: Record<number, number[]> = {};
    for (const [k, v] of Object.entries(rawByWd as Record<string, unknown>)) {
      const wd = parseInt(k, 10);
      if (wd < 0 || wd > 6 || Number.isNaN(wd)) continue;
      if (!Array.isArray(v)) continue;
      out[wd] = (v as unknown[])
        .map((idx) => parseInt(String(idx), 10))
        .filter(validSlotIdx) as number[];
    }
    if (Object.keys(out).length > 0) allowedByWd = out;
  }
  return {
    ...rest,
    id: rest.id || `c${Date.now()}`,
    name: rest.name || "New Course",
    faculty: rest.faculty || "Faculty",
    classroom: typeof rest.classroom === "string" ? rest.classroom : "",
    disabled: rest.disabled === true ? true : undefined,
    common: rest.common === true ? true : undefined,
    color: rest.color || COLORS[0],
    durationSlots: cleanDurationSlots(rest.durationSlots, slots),
    minDurationSlots:
      (rest as { minDurationSlots?: unknown }).minDurationSlots !== undefined
        ? Math.max(
            1,
            parseInt(String((rest as { minDurationSlots?: unknown }).minDurationSlots), 10),
          )
        : undefined,
    maxDurationSlots:
      (rest as { maxDurationSlots?: unknown }).maxDurationSlots !== undefined
        ? Math.max(
            1,
            parseInt(String((rest as { maxDurationSlots?: unknown }).maxDurationSlots), 10),
          )
        : undefined,
    stretchRuleEnabled: (rest as { stretchRuleEnabled?: unknown }).stretchRuleEnabled !== false,
    weeklyPeriods: Math.max(0, Math.floor(rest.weeklyPeriods ?? 0)),
    lectureHours: Math.max(0, Math.floor(rest.lectureHours ?? 0)) || undefined,
    tutorialHours: Math.max(0, Math.floor(rest.tutorialHours ?? 0)) || undefined,
    practicalHours: Math.max(0, Math.floor(rest.practicalHours ?? 0)) || undefined,
    credits: Math.max(0, Math.floor(rest.credits ?? 0)) || undefined,
    // Explicit totalSessions always wins as an override. When it's blank we
    // fall back to the LTPC-derived total via courseTotalTarget().
    totalSessions: (() => {
      if (rest.totalSessions === undefined || rest.totalSessions === null) return undefined;
      const n = Math.max(0, Math.floor(rest.totalSessions));
      return n > 0 ? n : undefined;
    })(),
    allowedWeekdays: (rest.allowedWeekdays ?? [])
      .map((day) => parseInt(String(day), 10))
      .filter((day) => !Number.isNaN(day) && day >= 0 && day <= 6),
    allowedSlots: rawAllowedSlots.map((idx) => parseInt(String(idx), 10)).filter(validSlotIdx),
    allowedSlotsByWeekday: allowedByWd,
    fromDate: isValidIso(rest.fromDate) ? rest.fromDate : undefined,
    toDate: isValidIso(rest.toDate) ? rest.toDate : undefined,
  };
};
const normalizeStateSnapshot = (snapshot: SavedState): State => {
  const base = defaultState();
  const slots =
    Array.isArray(snapshot.slots) && snapshot.slots.length > 0 ? snapshot.slots : base.slots;
  const legacyCourses = snapshot.courses?.map((course) => cleanCourse(course, slots));
  const sourceClasses =
    Array.isArray(snapshot.classes) && snapshot.classes.length > 0
      ? snapshot.classes
      : base.classes;
  const classes: ClassData[] = sourceClasses.map((cls, index) => {
    const savedCourses =
      cls.courses && cls.courses.length > 0
        ? cls.courses.map((course) => cleanCourse(course, slots))
        : legacyCourses
          ? legacyCourses.map((course) => ({ ...course }))
          : [];
    return {
      id: cls.id || `k${index + 1}`,
      name: cls.name || `Class ${String.fromCharCode(65 + index)}`,
      grid: cls.grid ?? {},
      courses: savedCourses,
      department: typeof cls.department === "string" ? cls.department : undefined,
      fromDate: isValidIso(cls.fromDate) ? cls.fromDate : undefined,
      toDate: isValidIso(cls.toDate) ? cls.toDate : undefined,
    };
  });
  return {
    fromDate: snapshot.fromDate || base.fromDate,
    toDate: snapshot.toDate || base.toDate,
    slots,
    classes,
    frozen: Boolean(snapshot.frozen),
    ignoredConflicts: Array.isArray(snapshot.ignoredConflicts) ? snapshot.ignoredConflicts : [],
  };
};

const countCoursePeriodsInDates = (
  grid: Record<string, Cell>,
  course: Course,
  slots: Slot[],
  dateList: string[],
): number => {
  let count = 0;
  dateList.forEach((date) => {
    for (let i = 0; i < slots.length; i++) {
      const cell = grid[`${date}-${i}`];
      if (cell?.kind === "course" && cell.courseId === course.id) {
        count++;
      }
    }
  });
  return count;
};

const countCourseSessionsInDates = (
  grid: Record<string, Cell>,
  course: Course,
  slots: Slot[],
  dateList: string[],
  onStart?: (date: string, slotIdx: number) => void,
): number => {
  let count = 0;
  dateList.forEach((date) => {
    let slotIdx = 0;
    while (slotIdx < slots.length) {
      const cell = grid[`${date}-${slotIdx}`];
      if (cell?.kind !== "course" || cell.courseId !== course.id) {
        slotIdx++;
        continue;
      }
      // Found start of a contiguous block
      count++;
      onStart?.(date, slotIdx);

      // Skip the rest of this contiguous block
      while (slotIdx < slots.length) {
        const nextCell = grid[`${date}-${slotIdx}`];
        if (nextCell?.kind === "course" && nextCell.courseId === course.id) {
          slotIdx++;
        } else {
          break;
        }
      }
    }
  });
  return count;
};

const courseMinStretch = (course: Course): number => {
  return course.minDurationSlots ?? course.durationSlots ?? 1;
};

const courseMaxStretch = (course: Course): number => {
  return course.maxDurationSlots ?? course.durationSlots ?? 1;
};

const courseBlockFitsRules = (
  grid: Record<string, Cell>,
  course: Course,
  slots: Slot[],
  date: string,
  start: number,
  actualLength: number,
): boolean => {
  if (!courseAllowedOn(course, date)) return false;
  if (course.stretchRuleEnabled === false) {
    for (let i = 0; i < actualLength; i++) {
      const idx = start + i;
      if (idx >= slots.length || slots[idx]?.isBreak) return false;
      if (!courseAllowedSlotOn(course, idx, date)) return false;
    }
    return true;
  }
  const min = courseMinStretch(course);
  const max = courseMaxStretch(course);
  if (actualLength < min || actualLength > max) return false;
  for (let i = 0; i < actualLength; i++) {
    const idx = start + i;
    if (idx >= slots.length || slots[idx]?.isBreak) return false;
    if (!courseAllowedSlotOn(course, idx, date)) return false;
  }
  return true;
};

const courseSpanFitsRules = (
  course: Course,
  slots: Slot[],
  date: string,
  start: number,
): boolean => {
  if (!courseAllowedOn(course, date)) return false;
  if (course.stretchRuleEnabled === false) {
    const span = cleanDurationSlots(course.durationSlots, slots);
    for (let i = 0; i < span; i++) {
      const idx = start + i;
      if (idx >= slots.length || slots[idx]?.isBreak) return false;
      if (!courseAllowedSlotOn(course, idx, date)) return false;
    }
    return true;
  }
  const min = courseMinStretch(course);
  const max = courseMaxStretch(course);
  const span = cleanDurationSlots(course.durationSlots, slots);
  if (span < min || span > max) return false;
  for (let i = 0; i < span; i++) {
    const idx = start + i;
    if (idx >= slots.length || slots[idx]?.isBreak) return false;
    if (!courseAllowedSlotOn(course, idx, date)) return false;
  }
  return true;
};

const countCourseRuleCapacity = (course: Course, slots: Slot[], dateList: string[]): number => {
  const nonBreakStarts = slots
    .map((slot, idx) => ({ slot, idx }))
    .filter(({ slot }) => !slot.isBreak)
    .map(({ idx }) => idx);
  const span = cleanDurationSlots(course.durationSlots, slots);
  return dateList.reduce((sum, date) => {
    if (!courseAllowedOn(course, date)) return sum;
    const starts = effectiveAllowedSlots(course, date) ?? nonBreakStarts;
    const validStarts = starts
      .filter((start) => {
        return (
          start >= 0 && start < slots.length && courseSpanFitsRules(course, slots, date, start)
        );
      })
      .sort((a, b) => a - b);
    let count = 0;
    let nextFreeStart = 0;
    validStarts.forEach((start) => {
      if (start < nextFreeStart) return;
      count++;
      nextFreeStart = start + span;
    });
    return sum + count;
  }, 0);
};

type Tool =
  | { kind: "course"; courseId: string }
  | { kind: "break" }
  | { kind: "blocked" }
  | { kind: "erase" };

function Index() {
  const [state, rawSetState] = useState<State>(() => defaultState(INITIAL_FROM_DATE));
  const [past, setPast] = useState<State[]>([]);
  const [future, setFuture] = useState<State[]>([]);
  const dragStartStateRef = useRef<State | null>(null);

  const setState = useCallback(
    (updater: State | ((prev: State) => State), isIntermediate = false) => {
      rawSetState((current) => {
        const next = typeof updater === "function" ? updater(current) : updater;
        if (JSON.stringify(current) !== JSON.stringify(next)) {
          if (!isIntermediate) {
            setPast((p) => {
              const updated = [...p, current];
              if (updated.length > 5) {
                updated.shift();
              }
              return updated;
            });
            setFuture([]);
          }
        }
        return next;
      });
    },
    [],
  );
  const [activeClassId, setActiveClassId] = useState("k1");
  const [hydrated, setHydrated] = useState(false);
  const [picker, setPicker] = useState<{ date: string; slotIdx: number } | null>(null);
  const [armedTool, setArmedTool] = useState<Tool | null>(null);
  const [isPainting, setIsPainting] = useState(false);
  const [overrideMode, setOverrideMode] = useState(false);
  const [rulesFor, setRulesFor] = useState<string | null>(null); // courseId in activeClass
  const [bulkWeekdays, setBulkWeekdays] = useState<number[]>([]);
  const [bulkSlots, setBulkSlots] = useState<number[]>([]);
  const [bulkAllClasses, setBulkAllClasses] = useState(false);
  const [autoFillReport, setAutoFillReport] = useState<string>("");
  const [autoStatus, setAutoStatus] = useState<{ label: string; phase: string } | null>(null);
  const [pendingScrollClassId, setPendingScrollClassId] = useState<string | null>(null);
  const [pendingBlockCsv, setPendingBlockCsv] = useState<{ name: string; text: string } | null>(
    null,
  );
  const [blockReport, setBlockReport] = useState<string>("");
  const [cursorPos, setCursorPos] = useState<{ x: number; y: number } | null>(null);
  const [courseContextMenu, setCourseContextMenu] = useState<{
    courseId: string;
    x: number;
    y: number;
  } | null>(null);
  const [classContextMenu, setClassContextMenu] = useState<{
    classId: string;
    x: number;
    y: number;
  } | null>(null);
  const [facultyPanelOpen, setFacultyPanelOpen] = useState(false);
  const [controlPanelOpen, setControlPanelOpen] = useState(false);
  const [expandedControlClass, setExpandedControlClass] = useState<string | null>(null);
  const [selectedFaculty, setSelectedFaculty] = useState<string | null>(null);
  const [editingFaculty, setEditingFaculty] = useState<string | null>(null);
  const [facultyEditVal, setFacultyEditVal] = useState("");
  const gridRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const adminInputRef = useRef<HTMLInputElement>(null);
  const [adminOpen, setAdminOpen] = useState(false);
  const [adminRefs, setAdminRefs] = useState<Array<{ id: string; name: string; state: State }>>([]);
  const [warningsPanelOpen, setWarningsPanelOpen] = useState(false);
  const [pendingScroll, setPendingScroll] = useState<{
    classId: string;
    date: string;
    slotIdx: number;
  } | null>(null);

  // Map slot index → 1-based period number, skipping breaks.
  const periodNumberFor = (slotIdx: number): number =>
    state.slots.slice(0, slotIdx + 1).filter((x) => !x.isBreak).length;
  const periodLabelFor = (slotIdx: number): string =>
    state.slots[slotIdx]?.isBreak ? "Br" : `P${periodNumberFor(slotIdx)}`;

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const merged = normalizeStateSnapshot(JSON.parse(raw) as SavedState);
        setState(merged);
        setActiveClassId(merged.classes[0]?.id ?? "");
      } else {
        setState(defaultState());
      }
    } catch {
      // Ignore
    }
    setHydrated(true);
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

  const undo = useCallback(() => {
    setPast((p) => {
      if (p.length === 0) return p;
      const prev = p[p.length - 1];
      const newPast = p.slice(0, -1);
      rawSetState((current) => {
        setFuture((f) => {
          const updated = [current, ...f];
          if (updated.length > 5) {
            updated.pop();
          }
          return updated;
        });
        return prev;
      });
      return newPast;
    });
  }, []);

  const redo = useCallback(() => {
    setFuture((f) => {
      if (f.length === 0) return f;
      const next = f[0];
      const newFuture = f.slice(1);
      rawSetState((current) => {
        setPast((p) => {
          const updated = [...p, current];
          if (updated.length > 5) {
            updated.shift();
          }
          return updated;
        });
        return next;
      });
      return newFuture;
    });
  }, []);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const activeEl = document.activeElement;
      const isInput =
        activeEl &&
        (activeEl.tagName === "INPUT" ||
          activeEl.tagName === "TEXTAREA" ||
          activeEl.getAttribute("contenteditable") === "true");

      if (isInput) return;

      const isMac = navigator.platform.toUpperCase().indexOf("MAC") >= 0;
      const isCmdOrCtrl = isMac ? e.metaKey : e.ctrlKey;

      if (isCmdOrCtrl) {
        if (e.key.toLowerCase() === "z") {
          e.preventDefault();
          undo();
        } else if (e.key.toLowerCase() === "r" || e.key.toLowerCase() === "y") {
          e.preventDefault();
          redo();
        }
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [undo, redo]);

  useEffect(() => {
    const up = () => {
      setIsPainting(false);
      if (dragStartStateRef.current) {
        const initial = dragStartStateRef.current;
        dragStartStateRef.current = null;
        rawSetState((current) => {
          if (JSON.stringify(initial) !== JSON.stringify(current)) {
            setPast((p) => {
              const updated = [...p, initial];
              if (updated.length > 5) {
                updated.shift();
              }
              return updated;
            });
            setFuture([]);
          }
          return current;
        });
      }
    };
    window.addEventListener("mouseup", up);
    return () => window.removeEventListener("mouseup", up);
  }, []);

  useEffect(() => {
    const dismiss = () => {
      setCourseContextMenu(null);
      setClassContextMenu(null);
    };
    window.addEventListener("click", dismiss);
    window.addEventListener("contextmenu", dismiss);
    return () => {
      window.removeEventListener("click", dismiss);
      window.removeEventListener("contextmenu", dismiss);
    };
  }, []);

  const activeClass = state.classes.find((c) => c.id === activeClassId) ?? state.classes[0];
  const dates = useMemo(
    () =>
      activeClass
        ? daysBetween(classFromDate(activeClass, state), classToDate(activeClass, state))
        : daysBetween(state.fromDate, state.toDate),
    [activeClass?.id, activeClass?.fromDate, activeClass?.toDate, state.fromDate, state.toDate],
  );

  useEffect(() => {
    if (!pendingScrollClassId || pendingScrollClassId !== activeClass?.id) return;
    const frame = window.requestAnimationFrame(() => {
      const row = gridRef.current?.querySelector<HTMLElement>('[data-filled-row="true"]');
      row?.scrollIntoView({ block: "start", inline: "nearest", behavior: "smooth" });
      setPendingScrollClassId(null);
    });
    return () => window.cancelAnimationFrame(frame);
  }, [activeClass, pendingScrollClassId, state.classes]);

  // Effect to handle scroll and highlight when navigating to conflicts
  useEffect(() => {
    if (pendingScroll && activeClassId === pendingScroll.classId) {
      const timer = setTimeout(() => {
        const cellId = `cell-${pendingScroll.classId}-${pendingScroll.date}-${pendingScroll.slotIdx}`;
        const element = document.getElementById(cellId);
        if (element) {
          element.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
          element.classList.add("cell-highlight-flash");
          setTimeout(() => {
            element.classList.remove("cell-highlight-flash");
          }, 3000);
        }
        setPendingScroll(null);
      }, 150);
      return () => clearTimeout(timer);
    }
  }, [activeClassId, pendingScroll]);

  // Conflict detection across classes (global check)
  const conflicts = useMemo(() => {
    const set = new Set<string>();
    if (state.frozen) return set;

    const union = new Set<string>();
    state.classes.forEach((cls) => {
      const clsDates = daysBetween(classFromDate(cls, state), classToDate(cls, state));
      clsDates.forEach((d) => union.add(d));
    });
    const allDates = Array.from(union);

    allDates.forEach((date) => {
      state.slots.forEach((_, i) => {
        const key = `${date}-${i}`;
        const facultyToAssignments: Record<string, Array<{ classId: string; course: Course }>> = {};
        state.classes.forEach((cls) => {
          const cell = cls.grid[key];
          if (cell?.kind === "course") {
            const course = cls.courses.find((c) => c.id === cell.courseId);
            if (!course) return;
            getFaculties(course).forEach((f) => {
              (facultyToAssignments[f] ??= []).push({ classId: cls.id, course });
            });
            // Rule violation: course placed on a weekday or period it isn't allowed
            if (!courseAllowedOn(course, date) || !courseAllowedSlotOn(course, i, date)) {
              const keyStr = `${cls.id}:${key}`;
              if (!state.ignoredConflicts?.includes(keyStr)) {
                set.add(keyStr);
              }
            }
          }
        });

        Object.entries(facultyToAssignments).forEach(([_, assignments]) => {
          for (let aIdx = 0; aIdx < assignments.length; aIdx++) {
            for (let bIdx = aIdx + 1; bIdx < assignments.length; bIdx++) {
              const a = assignments[aIdx];
              const b = assignments[bIdx];
              if (a.classId !== b.classId) {
                const isSameCommonCourse =
                  a.course.common &&
                  b.course.common &&
                  a.course.name.toLowerCase().trim() === b.course.name.toLowerCase().trim();

                if (!isSameCommonCourse) {
                  const keyStrA = `${a.classId}:${key}`;
                  const keyStrB = `${b.classId}:${key}`;
                  if (!state.ignoredConflicts?.includes(keyStrA)) {
                    set.add(keyStrA);
                  }
                  if (!state.ignoredConflicts?.includes(keyStrB)) {
                    set.add(keyStrB);
                  }
                }
              }
            }
          }
        });
      });
    });
    return set;
  }, [state]);

  const conflictDetailsList = useMemo(() => {
    const list: Array<{
      id: string;
      classId: string;
      className: string;
      date: string;
      weekday: string;
      slotIdx: number;
      periodLabel: string;
      type: "faculty" | "rule";
      faculty?: string;
      courseName?: string;
      description: string;
    }> = [];

    if (state.frozen) return list;

    const union = new Set<string>();
    state.classes.forEach((cls) => {
      const clsDates = daysBetween(classFromDate(cls, state), classToDate(cls, state));
      clsDates.forEach((d) => union.add(d));
    });
    const allDates = Array.from(union).sort();

    allDates.forEach((date) => {
      const dayOfWeek = utcDateFromIso(date)?.getUTCDay() ?? 0;
      const dayName = WEEKDAY_FULL[dayOfWeek];
      state.slots.forEach((_, i) => {
        const key = `${date}-${i}`;
        const facultyToAssignments: Record<
          string,
          Array<{ classId: string; className: string; course: Course }>
        > = {};

        state.classes.forEach((cls) => {
          const cell = cls.grid[key];
          if (cell?.kind === "course") {
            const course = cls.courses.find((c) => c.id === cell.courseId);
            if (!course) return;

            getFaculties(course).forEach((f) => {
              (facultyToAssignments[f] ??= []).push({
                classId: cls.id,
                className: cls.name,
                course,
              });
            });

            // Rule violation
            if (!courseAllowedOn(course, date) || !courseAllowedSlotOn(course, i, date)) {
              const keyStr = `${cls.id}:${key}`;
              if (!state.ignoredConflicts?.includes(keyStr)) {
                list.push({
                  id: `${cls.id}:${key}:rule`,
                  classId: cls.id,
                  className: cls.name,
                  date,
                  weekday: dayName,
                  slotIdx: i,
                  periodLabel: periodLabelFor(i),
                  type: "rule",
                  courseName: course.name || course.id,
                  description: `Course "${course.name || course.id}" is placed on a slot/weekday not allowed by its rules.`,
                });
              }
            }
          }
        });

        // Faculty conflicts
        Object.entries(facultyToAssignments).forEach(([faculty, assignments]) => {
          for (let aIdx = 0; aIdx < assignments.length; aIdx++) {
            const a = assignments[aIdx];
            const matchingClashes = assignments.filter((b) => {
              if (a.classId === b.classId) return false;
              const isSameCommonCourse =
                a.course.common &&
                b.course.common &&
                a.course.name.toLowerCase().trim() === b.course.name.toLowerCase().trim();
              return !isSameCommonCourse;
            });

            if (matchingClashes.length > 0) {
              const keyStr = `${a.classId}:${key}`;
              if (!state.ignoredConflicts?.includes(keyStr)) {
                const others = matchingClashes
                  .map((b) => `"${b.course.name || b.course.id}" in ${b.className}`)
                  .join(" & ");

                list.push({
                  id: `${a.classId}:${key}:faculty:${faculty}`,
                  classId: a.classId,
                  className: a.className,
                  date,
                  weekday: dayName,
                  slotIdx: i,
                  periodLabel: periodLabelFor(i),
                  type: "faculty",
                  faculty,
                  courseName: a.course.name || a.course.id,
                  description: `Faculty "${faculty}" is scheduled to teach "${a.course.name || a.course.id}" here, but is also teaching ${others}.`,
                });
              }
            }
          }
        });
      });
    });

    list.sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      if (a.slotIdx !== b.slotIdx) return a.slotIdx - b.slotIdx;
      return a.className.localeCompare(b.className);
    });

    return list;
  }, [state, periodLabelFor]);

  const applyTool = (date: string, slotIdx: number, tool: Tool, isIntermediate = false) => {
    if (state.frozen) {
      setAutoFillReport("Timetable is frozen — unfreeze to make changes.");
      return;
    }
    if (tool.kind === "course") {
      const active = state.classes.find((c) => c.id === activeClassId);
      const course = active?.courses.find((c) => c.id === tool.courseId);
      if (!active || !course || (!overrideMode && !courseAllowedOn(course, date))) {
        setAutoFillReport("Cannot place course — this date is outside its rules.");
        return;
      }
      const span = cleanDurationSlots(course.durationSlots, state.slots);
      for (let k = 0; k < span; k++) {
        const idx = slotIdx + k;
        if (
          idx >= state.slots.length ||
          state.slots[idx]?.isBreak ||
          (!overrideMode && !courseAllowedSlotOn(course, idx, date))
        ) {
          setAutoFillReport(
            "Cannot place course — the full session must fit only inside selected rule periods.",
          );
          return;
        }
        const key = `${date}-${idx}`;
        const facultyBusy = state.classes.some((cls) => {
          if (cls.id === activeClassId) return false;
          const cell = cls.grid[key];
          if (cell?.kind !== "course") return false;
          const otherCourse = cls.courses.find((c) => c.id === cell.courseId);
          return otherCourse ? sharesFaculty(otherCourse, course) : false;
        });
        if (facultyBusy && !overrideMode) {
          setAutoFillReport(
            "Cannot place course — this faculty is already assigned in another class at that time.",
          );
          return;
        }
      }
    }
    setState(
      (s) => ({
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
              if (course && !overrideMode && !courseAllowedSlotOn(course, idx, date)) continue;
            }
            const key = `${date}-${idx}`;
            if (tool.kind === "erase") delete grid[key];
            else if (tool.kind === "break") grid[key] = { kind: "break", label: "Break" };
            else if (tool.kind === "blocked") grid[key] = { kind: "blocked", label: "Blocked" };
            else grid[key] = { kind: "course", courseId: tool.courseId, locked: true };
          }
          return { ...cls, grid };
        }),
      }),
      isIntermediate,
    );
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
      return {
        ...s,
        classes: s.classes.map((cls) => {
          if (!allClasses && cls.id !== activeClassId) return cls;
          const grid = { ...cls.grid };
          const targetDates = classDatesFor(cls, s).filter((iso) => wdSet.has(weekdayOf(iso)));
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
    if (state.frozen) {
      setAutoFillReport("Timetable is frozen — unfreeze to clear.");
      return;
    }
    if (
      !confirm(
        "Clear every course assignment from all classes? Breaks and blocked slots will stay.",
      )
    )
      return;
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

  const clearCurrentClass = () => {
    if (state.frozen) {
      setAutoFillReport("Timetable is frozen — unfreeze to clear.");
      return;
    }
    if (!activeClass) return;
    if (
      !confirm(
        `Clear all course assignments from "${activeClass.name}"? Breaks and blocked slots will stay.`,
      )
    )
      return;
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        if (cls.id !== activeClassId) return cls;
        const grid: Record<string, Cell> = {};
        Object.entries(cls.grid).forEach(([key, cell]) => {
          if (cell.kind !== "course") grid[key] = cell;
        });
        return { ...cls, grid };
      }),
    }));
    setAutoFillReport(`Cleared "${activeClass.name}" — course assignments removed.`);
  };

  const clearCoursePlacements = (courseId: string) => {
    if (state.frozen) {
      setAutoFillReport("Timetable is frozen — unfreeze to clear.");
      return;
    }
    const course = activeClass?.courses.find((c) => c.id === courseId);
    if (!course || !activeClass) return;
    if (!confirm(`Clear all scheduled placements for "${course.name}" from "${activeClass.name}"?`))
      return;
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        if (cls.id !== activeClassId) return cls;
        const grid = { ...cls.grid };
        Object.keys(grid).forEach((key) => {
          const cell = grid[key];
          if (cell?.kind === "course" && cell.courseId === courseId) {
            delete grid[key];
          }
        });
        return { ...cls, grid };
      }),
    }));
    setAutoFillReport(`Cleared all placements for course "${course.name}".`);
  };

  const onCellMouseDown = (date: string, slotIdx: number, e: React.MouseEvent) => {
    if (state.slots[slotIdx]?.isBreak) {
      e.preventDefault();
      return;
    }
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
      dragStartStateRef.current = state;
      applyTool(date, slotIdx, armedTool, true);
      setIsPainting(true);
    } else {
      setPicker({ date, slotIdx });
    }
  };
  const onCellEnter = (date: string, slotIdx: number, e: React.MouseEvent) => {
    // Only paint while a mouse button is actually held down
    if (isPainting && armedTool && (e.buttons & 1) === 1) {
      applyTool(date, slotIdx, armedTool, true);
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
          {
            id,
            name: `Class ${String.fromCharCode(65 + s.classes.length)}`,
            grid: {},
            courses: [],
            fromDate: s.fromDate,
            toDate: s.toDate,
          },
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

  const duplicateClass = (id: string) => {
    setState((s) => {
      const original = s.classes.find((c) => c.id === id);
      if (!original) return s;
      const newId = `k${Date.now()}`;
      const copy: ClassData = {
        id: newId,
        name: `${original.name} (Copy)`,
        grid: JSON.parse(JSON.stringify(original.grid)),
        courses: original.courses.map((c) => ({ ...c })),
        department: original.department,
        fromDate: original.fromDate,
        toDate: original.toDate,
      };
      return {
        ...s,
        classes: [...s.classes, copy],
      };
    });
  };

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
              classroom: "",
              color: getUniqueCourseColor(cls.courses),
              durationSlots: 1,
              weeklyPeriods: 3,
            },
          ],
        };
      }),
    }));
  const duplicateCourse = (courseId: string) =>
    setState((s) => {
      const activeClass = s.classes.find((c) => c.id === activeClassId);
      if (!activeClass) return s;
      const original = activeClass.courses.find((c) => c.id === courseId);
      if (!original) return s;
      const copy: Course = {
        ...original,
        id: `c${Date.now()}`,
        name: `${original.name} (Copy)`,
        color: getUniqueCourseColor(activeClass.courses),
      };
      return {
        ...s,
        classes: s.classes.map((cls) => {
          if (cls.id !== activeClassId) return cls;
          return {
            ...cls,
            courses: [...cls.courses, copy],
          };
        }),
      };
    });

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
      const wk =
        1 +
        Math.round(
          ((t.getTime() - first.getTime()) / 86400000 - 3 + ((first.getUTCDay() + 6) % 7)) / 7,
        );
      return `${t.getUTCFullYear()}-W${wk}`;
    };

    setState((s) => {
      const classes: ClassData[] = s.classes.map((cls) => ({
        ...cls,
        grid: { ...cls.grid },
      }));
      const workingDates = daysBetween(s.fromDate, s.toDate);
      const workingDatesByClass = new Map<string, string[]>(
        classes.map((cls) => [cls.id, classDatesFor(cls, s)]),
      );

      let totalTarget = 0;
      let placedCount = 0;
      const unmet: string[] = [];

      const mutableClasses = classes.filter((cls) => cls.id === activeClassId);

      if (opts.overwrite) {
        mutableClasses.forEach((cls) => {
          Object.keys(cls.grid).forEach((k) => {
            const cell = cls.grid[k];
            if (cell && cell.kind === "course" && !cell.locked) delete cls.grid[k];
          });
        });
      }

      const removeInvalidPlacements = (): number => {
        let removed = 0;
        const deleteCourseRun = (cls: ClassData, date: string, start: number, courseId: string) => {
          let idx = start;
          while (idx < s.slots.length) {
            const key = `${date}-${idx}`;
            const cell = cls.grid[key];
            if (cell?.kind !== "course" || cell.courseId !== courseId) break;
            // Never remove frozen/locked placements — they are permanent
            // overrides and must survive Fill by Rules.
            if (cell.locked) {
              idx++;
              continue;
            }
            delete cls.grid[key];
            removed++;
            idx++;
          }
        };

        const removeRuleBreakers = () => {
          mutableClasses.forEach((cls) => {
            workingDates.forEach((date) => {
              for (let slotIdx = 0; slotIdx < s.slots.length; slotIdx++) {
                const key = `${date}-${slotIdx}`;
                const cell = cls.grid[key];
                if (cell?.kind !== "course") continue;
                if (cell.locked) continue;
                const prev = slotIdx > 0 ? cls.grid[`${date}-${slotIdx - 1}`] : undefined;
                if (prev?.kind === "course" && prev.courseId === cell.courseId) continue;
                const course = cls.courses.find((c) => c.id === cell.courseId);
                if (!course) {
                  deleteCourseRun(cls, date, slotIdx, cell.courseId);
                  continue;
                }
                let L = 0;
                while (slotIdx + L < s.slots.length) {
                  const part = cls.grid[`${date}-${slotIdx + L}`];
                  if (part?.kind === "course" && part.courseId === cell.courseId) {
                    L++;
                  } else {
                    break;
                  }
                }
                if (!courseBlockFitsRules(cls.grid, course, s.slots, date, slotIdx, L)) {
                  deleteCourseRun(cls, date, slotIdx, cell.courseId);
                }
                slotIdx += L - 1;
              }
            });
          });
        };

        removeRuleBreakers();
        workingDates.forEach((date) => {
          s.slots.forEach((_, slotIdx) => {
            const key = `${date}-${slotIdx}`;
            const byFaculty = new Map<string, ClassData[]>();
            classes.forEach((cls) => {
              const cell = cls.grid[key];
              if (cell?.kind !== "course") return;
              const course = cls.courses.find((c) => c.id === cell.courseId);
              if (!course) return;
              getFaculties(course).forEach((f) => {
                const list = byFaculty.get(f) ?? [];
                if (!list.includes(cls)) list.push(cls);
                byFaculty.set(f, list);
              });
            });
            byFaculty.forEach((busyClasses) => {
              if (busyClasses.length < 2) return;
              busyClasses.forEach((cls) => {
                if (!mutableClasses.includes(cls)) return;
                const cell = cls.grid[key];
                if (cell?.kind !== "course") return;
                if (cell.locked) return;
                delete cls.grid[key];
                removed++;
              });
            });
          });
        });
        removeRuleBreakers();
        return removed;
      };

      const safetyRemoved = removeInvalidPlacements();

      const facultyBusy: Record<string, Set<string>> = {};
      classes.forEach((cls) => {
        Object.entries(cls.grid).forEach(([key, cell]) => {
          if (cell.kind !== "course") return;
          const course = cls.courses.find((c) => c.id === cell.courseId);
          if (course) {
            const bucket = (facultyBusy[key] ??= new Set());
            getFaculties(course).forEach((f) => bucket.add(f));
          }
        });
      });
      const globalDateOrder = new Map(workingDates.map((date, index) => [date, index]));

      const allNonBreakStarts = s.slots
        .map((slot, idx) => ({ slot, idx }))
        .filter(({ slot }) => !slot.isBreak)
        .map(({ idx }) => idx);

      const startSlotsFor = (course: Course, date: string): number[] => {
        // If not strictRules, we ignore course period rules and can start anywhere
        if (!opts.strictRules) {
          return allNonBreakStarts;
        }
        const eff = effectiveAllowedSlots(course, date);
        const candidates = eff ?? allNonBreakStarts;
        return candidates
          .filter((idx) => idx >= 0 && idx < s.slots.length && !s.slots[idx].isBreak)
          .sort((a, b) => a - b);
      };

      const spanFitsCourse = (course: Course, start: number, date: string): boolean => {
        for (let i = 0; i < cleanDurationSlots(course.durationSlots, s.slots); i++) {
          const idx = start + i;
          if (idx >= s.slots.length) return false;
          if (s.slots[idx].isBreak) return false;
          if (opts.strictRules && !courseAllowedSlotOn(course, idx, date)) return false;
        }
        return true;
      };

      const canPlace = (
        cls: ClassData,
        course: Course,
        date: string,
        start: number,
        span: number,
      ): boolean => {
        if (!courseAllowedOn(course, date)) return false;
        if (!startSlotsFor(course, date).includes(start)) return false;
        if (opts.strictRules) {
          if (!courseBlockFitsRules(cls.grid, course, s.slots, date, start, span)) return false;
        } else {
          for (let i = 0; i < span; i++) {
            const idx = start + i;
            if (idx >= s.slots.length || s.slots[idx].isBreak) return false;
          }
        }
        for (let i = 0; i < span; i++) {
          const idx = start + i;
          const key = `${date}-${idx}`;
          const existing = cls.grid[key];
          if (existing && existing.kind !== "empty") return false;
          if (facultyBusy[key] && getFaculties(course).some((f) => facultyBusy[key].has(f)))
            return false;
        }
        return true;
      };

      const unavailableReason = (
        cls: ClassData,
        course: Course,
        date: string,
        start: number,
        span: number,
      ): string => {
        if (!courseAllowedOn(course, date)) return "outside course day/date rules";
        if (!startSlotsFor(course, date).includes(start)) return "period not selected in rules";
        if (opts.strictRules) {
          if (!courseBlockFitsRules(cls.grid, course, s.slots, date, start, span)) {
            return "span violates consecutive rules or selected periods";
          }
        } else {
          for (let i = 0; i < span; i++) {
            const idx = start + i;
            if (idx >= s.slots.length || s.slots[idx].isBreak) {
              return "span crosses a break";
            }
          }
        }
        for (let i = 0; i < span; i++) {
          const idx = start + i;
          const key = `${date}-${idx}`;
          const existing = cls.grid[key];
          if (existing?.kind === "blocked") return "slot is blocked";
          if (existing?.kind === "break") return "slot is a break";
          if (existing?.kind === "course") {
            return existing.courseId === course.id
              ? "selected rule slots already filled"
              : "class already has another course there";
          }
          if (facultyBusy[key] && getFaculties(course).some((f) => facultyBusy[key].has(f))) {
            const busy = classes
              .filter((other) => other.id !== cls.id)
              .map((other) => {
                const busyCell = other.grid[key];
                if (busyCell?.kind !== "course") return null;
                const busyCourse = other.courses.find((c) => c.id === busyCell.courseId);
                if (!busyCourse || !sharesFaculty(busyCourse, course)) return null;
                return `${other.name}${busyCourse.name ? ` (${busyCourse.name})` : ""}`;
              })
              .filter((value): value is string => Boolean(value));
            return busy.length > 0
              ? `faculty busy in ${busy.join(", ")}`
              : "faculty busy in another class";
          }
        }
        return "no open matching slot";
      };

      const countPlacedStarts = (cls: ClassData, course: Course, dateList: string[]) => {
        let placed = 0;
        const perDay: Record<string, number> = {};
        const perSlot: Record<number, number> = {};
        dateList.forEach((d) => {
          perDay[d] = 0;
        });
        countCourseSessionsInDates(cls.grid, course, s.slots, dateList, (d, i) => {
          if (!courseAllowedOn(course, d) || !spanFitsCourse(course, i, d)) return;
          perDay[d] = (perDay[d] ?? 0) + 1;
          perSlot[i] = (perSlot[i] ?? 0) + 1;
          placed++;
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
        dateOrder: Record<string, number>;
        label: string;
      };

      const tasks: AutoTask[] = [];
      const addTask = (
        cls: ClassData,
        course: Course,
        dateList: string[],
        desiredPeriods: number,
        label: string,
      ) => {
        const startsByDate: Record<string, number[]> = {};
        const dateOrder: Record<string, number> = {};
        let possibleStarts = 0;
        dateList.forEach((d, dateIdx) => {
          dateOrder[d] = dateIdx;
          if (!courseAllowedOn(course, d)) return;
          const starts = startSlotsFor(course, d);
          startsByDate[d] = starts;
          possibleStarts += starts.length;
        });
        if (possibleStarts === 0) return;

        // Count placed periods in the date range
        let placedPeriods = 0;
        dateList.forEach((d) => {
          for (let i = 0; i < s.slots.length; i++) {
            const cell = cls.grid[`${d}-${i}`];
            if (cell?.kind === "course" && cell.courseId === course.id) {
              placedPeriods++;
            }
          }
        });

        const remaining = Math.max(0, desiredPeriods - placedPeriods);
        totalTarget += remaining;
        if (remaining > 0) {
          tasks.push({
            cls,
            course,
            remaining,
            perDay: {},
            perSlot: {},
            startsByDate,
            dateOrder,
            label,
          });
        }
      };

      mutableClasses.forEach((cls) => {
        const clsDates = workingDatesByClass.get(cls.id) ?? workingDates;
        const clsState = stateForClass(cls, s);
        const clsWeeks = new Map<string, string[]>();
        clsDates.forEach((d) => {
          const key = weekKey(d);
          const week = clsWeeks.get(key);
          if (week) week.push(d);
          else clsWeeks.set(key, [d]);
        });
        cls.courses.forEach((course) => {
          if (course.disabled) return;
          const totalTarget = courseTotalTarget(course, courseSemesterWeeks(course, clsState));
          if (totalTarget > 0) {
            addTask(cls, course, clsDates, totalTarget, "total");
            return;
          }
          clsWeeks.forEach((weekDates, key) => {
            const desired = courseWeeklyTarget(course);
            addTask(cls, course, weekDates, desired, key);
          });
        });
      });

      const availableCount = (task: AutoTask) => {
        let count = 0;
        const min = courseMinStretch(task.course);
        const max = courseMaxStretch(task.course);
        const effMin = opts.strictRules ? Math.min(min, task.remaining) : 1;
        const effMax = opts.strictRules
          ? Math.min(max, task.remaining)
          : Math.min(cleanDurationSlots(task.course.durationSlots, s.slots), task.remaining);

        Object.entries(task.startsByDate).forEach(([date, starts]) => {
          starts.forEach((start) => {
            for (let span = effMax; span >= effMin; span--) {
              if (canPlace(task.cls, task.course, date, start, span)) {
                count++;
                break;
              }
            }
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
          const aRules =
            (a.course.allowedWeekdays?.length || 7) +
            Object.values(a.course.allowedSlotsByWeekday ?? {}).flat().length +
            (a.course.allowedSlots?.length || s.slots.length);
          const bRules =
            (b.course.allowedWeekdays?.length || 7) +
            Object.values(b.course.allowedSlotsByWeekday ?? {}).flat().length +
            (b.course.allowedSlots?.length || s.slots.length);
          return aRules - bRules;
        });

        let nextPlacement: {
          task: AutoTask;
          date: string;
          slot: number;
          span: number;
          score: number;
        } | null = null;
        for (const task of tasks) {
          if (task.remaining <= 0) continue;
          const min = courseMinStretch(task.course);
          const max = courseMaxStretch(task.course);
          const effMin = opts.strictRules ? Math.min(min, task.remaining) : 1;
          const effMax = opts.strictRules
            ? Math.min(max, task.remaining)
            : Math.min(cleanDurationSlots(task.course.durationSlots, s.slots), task.remaining);

          for (const [date, starts] of Object.entries(task.startsByDate)) {
            for (const sIdx of starts) {
              for (let span = effMax; span >= effMin; span--) {
                if (!canPlace(task.cls, task.course, date, sIdx, span)) continue;
                const taskIndex = tasks.indexOf(task);
                const currentAvailability = availability.get(task) ?? 0;
                const wouldBlockAnotherRequiredSlot = tasks.reduce((risk, other) => {
                  if (other === task || other.remaining <= 0) return risk;
                  const otherAvailability = availability.get(other) ?? 0;
                  if (otherAvailability === 0) return risk;
                  let blockedOptions = 0;
                  const otherMin = courseMinStretch(other.course);
                  const otherMax = courseMaxStretch(other.course);
                  const otherEffMin = opts.strictRules ? Math.min(otherMin, other.remaining) : 1;
                  const otherEffMax = opts.strictRules
                    ? Math.min(otherMax, other.remaining)
                    : Math.min(
                        cleanDurationSlots(other.course.durationSlots, s.slots),
                        other.remaining,
                      );

                  Object.entries(other.startsByDate).forEach(([otherDate, otherStarts]) => {
                    otherStarts.forEach((otherStart) => {
                      let hadValidSpan = false;
                      for (let otherSpan = otherEffMax; otherSpan >= otherEffMin; otherSpan--) {
                        if (canPlace(other.cls, other.course, otherDate, otherStart, otherSpan)) {
                          hadValidSpan = true;
                          break;
                        }
                      }
                      if (!hadValidSpan) return;

                      let overlaps = false;
                      for (let a = 0; a < span; a++) {
                        for (let otherSpan = otherEffMax; otherSpan >= otherEffMin; otherSpan--) {
                          for (let b = 0; b < otherSpan; b++) {
                            if (date === otherDate && sIdx + a === otherStart + b) overlaps = true;
                          }
                        }
                      }
                      if (!overlaps) return;
                      if (
                        other.cls.id === task.cls.id ||
                        sharesFaculty(other.course, task.course)
                      ) {
                        blockedOptions++;
                      }
                    });
                  });
                  return otherAvailability - blockedOptions < other.remaining ? risk + 1 : risk;
                }, 0);
                const avoidableRisk =
                  currentAvailability > task.remaining ? wouldBlockAnotherRequiredSlot : 0;
                const score =
                  avoidableRisk * 1000000000000 +
                  (globalDateOrder.get(date) ?? workingDates.length) * 1000000000 +
                  sIdx * 1000000 +
                  (task.dateOrder[date] ?? workingDates.length) * 10000 +
                  taskIndex * 100 +
                  (task.perDay[date] ?? 0) * 10 +
                  (task.perSlot[sIdx] ?? 0) +
                  classDayLoad(task.cls, date);
                if (!nextPlacement || score < nextPlacement.score) {
                  nextPlacement = { task, date, slot: sIdx, span, score };
                }
              }
            }
          }
        }
        if (!nextPlacement) continue;

        const task = nextPlacement.task;
        const span = nextPlacement.span;
        for (let i = 0; i < span; i++) {
          const key = `${nextPlacement.date}-${nextPlacement.slot + i}`;
          task.cls.grid[key] = { kind: "course", courseId: task.course.id };
          {
            const bucket = (facultyBusy[key] ??= new Set());
            getFaculties(task.course).forEach((f) => bucket.add(f));
          }
        }
        task.perDay[nextPlacement.date] = (task.perDay[nextPlacement.date] ?? 0) + 1;
        task.perSlot[nextPlacement.slot] = (task.perSlot[nextPlacement.slot] ?? 0) + 1;
        task.remaining -= span;
        placedCount += span;
        progressed = true;
      }

      tasks.forEach((task) => {
        if (task.remaining <= 0) return;
        const openNow = availableCount(task);
        const reasonCounts = new Map<string, number>();
        const min = courseMinStretch(task.course);
        const max = courseMaxStretch(task.course);
        const effMin = opts.strictRules ? Math.min(min, task.remaining) : 1;
        const effMax = opts.strictRules
          ? Math.min(max, task.remaining)
          : Math.min(cleanDurationSlots(task.course.durationSlots, s.slots), task.remaining);

        Object.entries(task.startsByDate).forEach(([date, starts]) => {
          starts.forEach((start) => {
            let hasValidSpan = false;
            for (let span = effMax; span >= effMin; span--) {
              if (canPlace(task.cls, task.course, date, start, span)) {
                hasValidSpan = true;
                break;
              }
            }
            if (hasValidSpan) return;
            const reason = unavailableReason(task.cls, task.course, date, start, effMax);
            reasonCounts.set(reason, (reasonCounts.get(reason) ?? 0) + 1);
          });
        });
        const rankedReasons = [...reasonCounts.entries()].sort((a, b) => b[1] - a[1]);
        const topReason =
          rankedReasons.find(([reason]) => reason !== "selected rule slots already filled")?.[0] ??
          rankedReasons[0]?.[0];
        unmet.push(
          `${task.cls.name} · ${task.course.name}: ${task.remaining} left${openNow === 0 ? ` (${topReason ?? "no open rule slots"})` : ` (${openNow} open rule slots)`}`,
        );
      });

      const auditIssues: string[] = [];
      workingDates.forEach((date) => {
        s.slots.forEach((_, slotIdx) => {
          const key = `${date}-${slotIdx}`;
          const facultyAtSlot = new Map<string, string[]>();
          classes.forEach((cls) => {
            const cell = cls.grid[key];
            if (cell?.kind !== "course") return;
            const course = cls.courses.find((c) => c.id === cell.courseId);
            if (!course) {
              auditIssues.push(
                `${cls.name} ${date} ${periodLabelFor(slotIdx)} has an unknown course`,
              );
              return;
            }
            const prev = slotIdx > 0 ? cls.grid[`${date}-${slotIdx - 1}`] : undefined;
            const isStart = !(prev?.kind === "course" && prev.courseId === cell.courseId);
            if (isStart) {
              let L = 0;
              while (slotIdx + L < s.slots.length) {
                const nextCell = cls.grid[`${date}-${slotIdx + L}`];
                if (nextCell?.kind === "course" && nextCell.courseId === cell.courseId) {
                  L++;
                } else {
                  break;
                }
              }
              if (!courseBlockFitsRules(cls.grid, course, s.slots, date, slotIdx, L)) {
                auditIssues.push(
                  `${cls.name} ${course.name} violates rules at ${date} ${periodLabelFor(slotIdx)}`,
                );
              }
            }
            getFaculties(course).forEach((f) => {
              const list = facultyAtSlot.get(f) ?? [];
              list.push(`${cls.name} · ${course.name}`);
              facultyAtSlot.set(f, list);
            });
          });
          facultyAtSlot.forEach((list, faculty) => {
            if (list.length > 1) {
              auditIssues.push(
                `${faculty} overlaps at ${date} ${periodLabelFor(slotIdx)} (${list.join(", ")})`,
              );
            }
          });
        });
      });

      const visibleCourseCounts = classes.map((cls) => {
        const count = workingDates.reduce((sum, date) => {
          return (
            sum +
            s.slots.reduce((slotSum, _, slotIdx) => {
              const cell = cls.grid[`${date}-${slotIdx}`];
              return slotSum + (cell?.kind === "course" ? 1 : 0);
            }, 0)
          );
        }, 0);
        return { id: cls.id, name: cls.name, count };
      });
      const activeVisibleCount =
        visibleCourseCounts.find((item) => item.id === activeClassId)?.count ?? 0;
      const firstVisibleClass = visibleCourseCounts.find((item) => item.count > 0);

      // Diagnostic feedback stays on screen instead of blocking with popups.
      queueMicrotask(() => {
        const mode = opts.strictRules
          ? "Fill by Rules"
          : opts.overwrite
            ? "Regenerate"
            : "Fill Empty";
        setPendingScrollClassId(activeClassId);
        if (totalTarget === 0) {
          const validDates = workingDates.length;
          const courseCount = classes.reduce((sum, cls) => sum + cls.courses.length, 0);
          const nonBreakCount = s.slots.filter((slot) => !slot.isBreak).length;
          const reason = [
            validDates === 0 ? "date range" : "",
            courseCount === 0 ? "courses" : "",
            nonBreakCount === 0 ? "non-break periods" : "",
          ]
            .filter(Boolean)
            .join(", ");
          setAutoFillReport(`${mode}: nothing to place${reason ? ` — check ${reason}.` : "."}`);
        } else if (unmet.length > 0) {
          const debugCourse = classes
            .flatMap((c) => c.courses)
            .find((c) => c.name.includes("DES1203") || c.id.includes("DES1203"));
          const debugInfo = debugCourse
            ? ` | DES1203 debug: from=${debugCourse.fromDate ?? "none"}, to=${debugCourse.toDate ?? "none"}, wds=${JSON.stringify(debugCourse.allowedWeekdays)}, slots=${JSON.stringify(debugCourse.allowedSlots)}, slotsWd=${JSON.stringify(debugCourse.allowedSlotsByWeekday)}, dur=${debugCourse.durationSlots}, min=${debugCourse.minDurationSlots ?? "none"}, max=${debugCourse.maxDurationSlots ?? "none"}, enabled=${debugCourse.stretchRuleEnabled !== false}`
            : "";
          setAutoFillReport(
            `${mode}: placed ${placedCount} of ${totalTarget}. ${safetyRemoved > 0 ? `Removed ${safetyRemoved} unsafe old cell${safetyRemoved === 1 ? "" : "s"}. ` : ""}Remaining: ${unmet.slice(0, 4).join("; ")}${debugInfo}`,
          );
        } else if (auditIssues.length > 0) {
          setAutoFillReport(
            `${mode}: safety audit found issues — ${auditIssues.slice(0, 3).join("; ")}`,
          );
        } else if (placedCount === 0) {
          if (firstVisibleClass) {
            const showingCount =
              activeVisibleCount > 0 ? activeVisibleCount : firstVisibleClass.count;
            const showingClass =
              activeVisibleCount > 0
                ? visibleCourseCounts.find((item) => item.id === activeClassId)?.name
                : firstVisibleClass.name;
            setAutoFillReport(
              `${mode}: already filled and rules verified — showing ${showingCount} course slot${showingCount === 1 ? "" : "s"}${showingClass ? ` in ${showingClass}` : ""}.${safetyRemoved > 0 ? ` Removed ${safetyRemoved} unsafe old cell${safetyRemoved === 1 ? "" : "s"}.` : ""}`,
            );
          } else {
            setAutoFillReport(
              `${mode}: no visible course slots were placed. Check that the selected dates match the course rules and that rule slots are not blocked.`,
            );
          }
        } else {
          setAutoFillReport(
            `${mode}: placed ${placedCount} of ${totalTarget} planned sessions. Rules verified.${safetyRemoved > 0 ? ` Removed ${safetyRemoved} unsafe old cell${safetyRemoved === 1 ? "" : "s"}.` : ""}`,
          );
        }
      });

      return { ...s, classes };
    });
  };

  const runAutoPopulate = (label: string, opts: { overwrite: boolean; strictRules?: boolean }) => {
    // When frozen we still allow auto-fill to run: locked cells are preserved
    // and only empty slots receive new placements.
    setAutoStatus({ label, phase: "Preparing…" });
    // Yield twice so the overlay paints before the synchronous solver runs.
    requestAnimationFrame(() => {
      setAutoStatus({ label, phase: "Placing sessions…" });
      requestAnimationFrame(() => {
        try {
          autoPopulate(opts);
        } finally {
          setAutoStatus({ label, phase: "Finalizing…" });
          setTimeout(() => setAutoStatus(null), 350);
        }
      });
    });
  };

  // ------------ CSV block upload ------------
  const downloadBlockTemplate = () => {
    const nonBreakCount = state.slots.filter((sl) => !sl.isBreak).length || 8;
    const d0 = state.fromDate || isoToday();
    const toDMY = (iso: string) => {
      const [y, m, d] = iso.split("-");
      return `${d}/${m}/${y}`;
    };
    const sample = [
      "# Blocker template. Save as .csv and upload via 'Upload blocker CSV'.",
      "# date    = DD/MM/YYYY",
      "# Reason  = shown in the blocked cell",
      `# Session = 'all' to block the whole day, or comma list of periods (1..${nonBreakCount}), e.g. '1,3,5,6'`,
      "date,Reason,Session",
      `${toDMY(d0)},Holiday,all`,
      `${toDMY(addDays(d0, 1))},Sports,"7,${nonBreakCount}"`,
      `${toDMY(addDays(d0, 2))},Assembly,"1,3,5,6"`,
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
        if (ch === '"' && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else if (ch === '"') inQ = false;
        else cur += ch;
      } else {
        if (ch === '"') inQ = true;
        else if (ch === ",") {
          out.push(cur);
          cur = "";
        } else cur += ch;
      }
    }
    out.push(cur);
    return out.map((v) => v.trim());
  };

  const applyBlockCsv = (text: string) => {
    const rawLines = text
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l && !l.startsWith("#"));
    if (rawLines.length === 0) {
      setBlockReport("CSV is empty.");
      return;
    }
    const header = parseCsvRow(rawLines[0]).map((c) => c.toLowerCase());
    const dateIdx = header.indexOf("date");
    const periodsIdx = (() => {
      const i = header.indexOf("session");
      return i >= 0 ? i : header.indexOf("periods");
    })();
    const labelIdx = (() => {
      const i = header.indexOf("reason");
      return i >= 0 ? i : header.indexOf("label");
    })();
    const scopeIdx = header.indexOf("scope");
    if (dateIdx < 0) {
      setBlockReport("CSV missing required 'date' column.");
      return;
    }

    const normalizeDate = (raw: string): string => {
      const s = (raw || "").trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
      const m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
      if (!m) return "";
      const d = m[1].padStart(2, "0");
      const mo = m[2].padStart(2, "0");
      let y = m[3];
      if (y.length === 2) y = (parseInt(y, 10) > 50 ? "19" : "20") + y;
      return `${y}-${mo}-${d}`;
    };

    const nonBreakIdxs = state.slots
      .map((sl, i) => ({ sl, i }))
      .filter((x) => !x.sl.isBreak)
      .map((x) => x.i);
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

    let applied = 0,
      skipped = 0,
      outOfRange = 0,
      noPeriods = 0;
    const errors: string[] = [];
    setState((s) => {
      const inRange = new Set<string>();
      s.classes.forEach((cls) => classDatesFor(cls, s).forEach((d) => inRange.add(d)));
      daysBetween(s.fromDate, s.toDate).forEach((d) => inRange.add(d));
      const classes: ClassData[] = s.classes.map((cls) => ({ ...cls, grid: { ...cls.grid } }));
      rawLines.slice(1).forEach((line, i) => {
        const cells = parseCsvRow(line);
        const date = normalizeDate(cells[dateIdx]);
        if (!date) {
          errors.push(`Row ${i + 2}: bad date "${cells[dateIdx]}" (use DD/MM/YYYY)`);
          skipped++;
          return;
        }
        const periods = parsePeriods(periodsIdx >= 0 ? cells[periodsIdx] : "all");
        if (periods.length === 0) {
          errors.push(`Row ${i + 2}: no valid periods parsed from "${cells[periodsIdx] ?? ""}"`);
          noPeriods++;
          skipped++;
          return;
        }
        if (!inRange.has(date)) {
          errors.push(
            `Row ${i + 2}: date ${date} is outside timetable range ${s.fromDate}..${s.toDate}`,
          );
          outOfRange++;
          skipped++;
          return;
        }
        const label = (labelIdx >= 0 ? cells[labelIdx] : "") || "Block";
        const scope = ((scopeIdx >= 0 ? cells[scopeIdx] : "") || "all").toLowerCase();
        const targets =
          scope === "all" || scope === "*" || scope === ""
            ? classes
            : classes.filter((c) => c.name.toLowerCase() === scope);
        if (targets.length === 0) {
          errors.push(`Row ${i + 2}: unknown scope "${cells[scopeIdx]}"`);
          skipped++;
          return;
        }
        targets.forEach((cls) => {
          const clsRange = new Set(classDatesFor(cls, s));
          if (!clsRange.has(date)) return;
          periods.forEach((slotIdx) => {
            cls.grid[`${date}-${slotIdx}`] = { kind: "blocked", label };
            applied++;
          });
        });
      });
      return { ...s, classes };
    });
    const extras: string[] = [];
    if (outOfRange)
      extras.push(
        `${outOfRange} row(s) outside timetable date range (extend From/To to include them).`,
      );
    if (noPeriods) extras.push(`${noPeriods} row(s) had no valid periods.`);
    const msg =
      `Applied ${applied} blocked cells.` +
      (skipped ? ` Skipped ${skipped} row(s).` : "") +
      (extras.length ? ` ${extras.join(" ")}` : "") +
      (errors.length ? ` ${errors.slice(0, 4).join(" | ")}` : "");
    setBlockReport(msg);
  };

  // Export
  const buildSheet = (cls: ClassData) => {
    const rows: string[][] = [];
    rows.push(["Week", "Day / Date", ...state.slots.map(slotLabel)]);
    const clsDates = classDatesFor(cls, state);
    clsDates.forEach((date) => {
      const { weekday, date: dstr } = dayLabel(date);
      const weekIndex = weekIndexForDate(cls, state, date);
      const row = [`W${weekIndex}`, `${weekday} ${dstr}`];
      state.slots.forEach((sl, i) => {
        if (sl.isBreak) {
          row.push("Break");
          return;
        }
        const cell = cls.grid[`${date}-${i}`];
        if (!cell) row.push("");
        else if (cell.kind === "break") row.push(`Break: ${cell.label}`);
        else if (cell.kind === "blocked") row.push(`Blocked: ${cell.label}`);
        else if (cell.kind === "course") {
          const c = cls.courses.find((x) => x.id === cell.courseId);
          if (!c) {
            row.push("");
          } else {
            const P = Math.max(0, c.practicalHours ?? 0);
            const LT = Math.max(0, c.lectureHours ?? 0) + Math.max(0, c.tutorialHours ?? 0);
            const isPractical = P > 0 && (LT === 0 || (c.durationSlots ?? 1) >= 2);
            const code = isPractical ? `${c.name}_P` : c.name;
            row.push(`${code} (${c.faculty})`);
          }
        } else row.push("");
      });
      rows.push(row);
    });
    return rows;
  };
  const exportExcel = () => {
    const wb = XLSXStyle.utils.book_new();
    const hexClean = (h: string) =>
      (h || "").replace("#", "").padStart(6, "0").slice(-6).toUpperCase();
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
      const clsDates = classDatesFor(cls, state);
      const ws = XLSXStyle.utils.aoa_to_sheet(data);
      const numCols = data[0].length;
      ws["!cols"] = Array.from({ length: numCols }, (_, i) => ({
        wch: i === 0 ? 8 : i === 1 ? 18 : 20,
      }));
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
          if (r === 0 || c === 0 || c === 1) {
            cellStyle.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } };
            cellStyle.fill = { patternType: "solid", fgColor: { rgb: "0D0D0D" } };
          } else {
            const date = clsDates[r - 1];
            const slotIdx = c - 2;
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
              cellStyle.font = {
                name: "Calibri",
                sz: 11,
                bold: true,
                color: { rgb: textColorFor(bg) },
              };
            }
          }
          (ws[addr] as { s?: unknown }).s = cellStyle;
        }
      }
      ws["!freeze"] = { xSplit: 2, ySplit: 1 };
      XLSXStyle.utils.book_append_sheet(wb, ws, cls.name.slice(0, 31) || "Class");
    });
    XLSXStyle.writeFile(wb, `timetable_${state.fromDate}_to_${state.toDate}.xlsx`);
  };
  // ─── Shared helpers: Summary+Gantt sheet ────────────────────────────────
  const buildSummaryData = (
    entries: Array<{ course: Course; cls: ClassData }>,
    refCls?: ClassData | null,
  ): {
    rows: (string | number)[][];
    ganttStartCol: number;
    weeks: number[];
    courseWeekHits: Set<number>[];
  } => {
    const ref = refCls ?? activeClass ?? state.classes[0];
    const allDates = daysBetween(state.fromDate, state.toDate);
    const weekNums = new Set<number>();
    if (ref) allDates.forEach((d) => weekNums.add(weekIndexForDate(ref, state, d)));
    const weeks = Array.from(weekNums).sort((a, b) => a - b);
    const ganttStartCol = 8;

    const courseWeekHits: Set<number>[] = entries.map(({ course, cls }) => {
      const hits = new Set<number>();
      classDatesFor(cls, state).forEach((date) => {
        state.slots.forEach((_, si) => {
          const cell = cls.grid[`${date}-${si}`];
          if (cell?.kind === "course" && cell.courseId === course.id)
            hits.add(weekIndexForDate(ref ?? cls, state, date));
        });
      });
      return hits;
    });

    const header: (string | number)[] = [
      "#",
      "Course",
      "Faculty",
      "L",
      "T",
      "P",
      "C",
      "Periods/Wk",
      ...weeks.map((w) => `W${w}`),
    ];
    const rows: (string | number)[][] = [header];
    entries.forEach(({ course }, idx) => {
      const L = Math.max(0, course.lectureHours ?? 0);
      const T = Math.max(0, course.tutorialHours ?? 0);
      const P = Math.max(0, course.practicalHours ?? 0);
      const C = Math.max(0, course.credits ?? 0);
      const wt = courseWeeklyTarget(course);
      rows.push([
        idx + 1,
        course.name,
        course.faculty || "",
        L || "",
        T || "",
        P || "",
        C || "",
        wt > 0 ? wt : "",
        ...weeks.map((w) => (courseWeekHits[idx].has(w) ? "\u25A0" : "")),
      ]);
    });
    return { rows, ganttStartCol, weeks, courseWeekHits };
  };

  const applySummaryStyles = (
    ws: ReturnType<typeof XLSXStyle.utils.aoa_to_sheet>,
    rows: (string | number)[][],
    ganttStartCol: number,
    weeks: number[],
    courseWeekHits: Set<number>[],
    entries: Array<{ course: Course }>,
    hexClean: (h: string) => string,
    textColorFor: (hex: string) => string,
  ) => {
    const numCols = rows[0]?.length ?? 0;
    const bdr = { style: "thin", color: { rgb: "CCCCCC" } } as const;
    const bb = { top: bdr, bottom: bdr, left: bdr, right: bdr };
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < numCols; c++) {
        const addr = XLSXStyle.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { t: "s", v: "" };
        const cs: Record<string, unknown> = {
          alignment: {
            horizontal: c >= ganttStartCol ? "center" : c === 0 ? "center" : "left",
            vertical: "center",
            wrapText: true,
          },
          border: bb,
          font: { name: "Calibri", sz: 11 },
        };
        if (r === 0) {
          cs.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } };
          cs.fill = { patternType: "solid", fgColor: { rgb: "0D0D0D" } };
        } else {
          const ei = r - 1;
          if (c >= ganttStartCol && ei < courseWeekHits.length) {
            const wn = weeks[c - ganttStartCol];
            if (wn !== undefined && courseWeekHits[ei].has(wn) && entries[ei]) {
              const bg = hexClean(entries[ei].course.color);
              cs.fill = { patternType: "solid", fgColor: { rgb: bg } };
              cs.font = { name: "Calibri", sz: 14, bold: true, color: { rgb: textColorFor(bg) } };
              cs.alignment = { horizontal: "center", vertical: "center" };
            }
          } else if (c <= 1) {
            cs.font = { name: "Calibri", sz: 11, bold: true };
          }
        }
        (ws[addr] as { s?: unknown }).s = cs;
      }
    }
    const nRows = rows.length;
    (ws as unknown as Record<string, unknown>)["!rows"] = Array.from({ length: nRows }, (_, i) => ({
      hpt: i === 0 ? 26 : 22,
    }));
    (ws as unknown as Record<string, unknown>)["!cols"] = [
      { wch: 5 },
      { wch: 30 },
      { wch: 20 },
      { wch: 5 },
      { wch: 5 },
      { wch: 5 },
      { wch: 5 },
      { wch: 12 },
      ...weeks.map(() => ({ wch: 5 })),
    ];
    (ws as unknown as Record<string, unknown>)["!freeze"] = { xSplit: 0, ySplit: 1 };
  };

  const applyTimetableStyles = (
    ws: ReturnType<typeof XLSXStyle.utils.aoa_to_sheet>,
    data: (string | number)[][],
    cls: ClassData,
    clsDates: string[],
    hexClean: (h: string) => string,
    textColorFor: (hex: string) => string,
  ) => {
    const numCols = data[0].length;
    const bdr = { style: "thin", color: { rgb: "CCCCCC" } } as const;
    const bb = { top: bdr, bottom: bdr, left: bdr, right: bdr };
    (ws as unknown as Record<string, unknown>)["!cols"] = Array.from(
      { length: numCols },
      (_, i) => ({ wch: i === 0 ? 8 : i === 1 ? 18 : 20 }),
    );
    (ws as unknown as Record<string, unknown>)["!rows"] = data.map((_, i) => ({
      hpt: i === 0 ? 24 : 32,
    }));
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < numCols; c++) {
        const addr = XLSXStyle.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { t: "s", v: "" };
        const cs: Record<string, unknown> = {
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: bb,
          font: { name: "Calibri", sz: 11 },
        };
        if (r === 0 || c === 0 || c === 1) {
          cs.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } };
          cs.fill = { patternType: "solid", fgColor: { rgb: "0D0D0D" } };
        } else {
          const date = clsDates[r - 1];
          const slotIdx = c - 2;
          const key = `${date}-${slotIdx}`;
          const isConf = conflicts.has(`${cls.id}:${key}`);
          const sl = state.slots[slotIdx];
          const cell = cls.grid[key];
          if (isConf) {
            cs.fill = { patternType: "solid", fgColor: { rgb: "FECACA" } };
            cs.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "991B1B" } };
          } else if (sl?.isBreak || cell?.kind === "break") {
            cs.fill = { patternType: "solid", fgColor: { rgb: "FDE68A" } };
            cs.font = { name: "Calibri", sz: 11, italic: true, color: { rgb: "78350F" } };
          } else if (cell?.kind === "blocked") {
            cs.fill = { patternType: "solid", fgColor: { rgb: "374151" } };
            cs.font = { name: "Calibri", sz: 11, color: { rgb: "FFFFFF" } };
          } else if (cell?.kind === "course") {
            const course = cls.courses.find((x) => x.id === cell.courseId);
            const bg = hexClean(course?.color ?? "#DDDDDD");
            cs.fill = { patternType: "solid", fgColor: { rgb: bg } };
            cs.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: textColorFor(bg) } };
          }
        }
        (ws[addr] as { s?: unknown }).s = cs;
      }
    }
    (ws as unknown as Record<string, unknown>)["!freeze"] = { xSplit: 2, ySplit: 1 };
  };

  // Export a single class's timetable as Summary + Timetable sheets
  const exportClassTimetable = (cls: ClassData) => {
    const hexClean = (h: string) =>
      (h || "").replace("#", "").padStart(6, "0").slice(-6).toUpperCase();
    const textColorFor = (hex: string) => {
      const h = hexClean(hex);
      const r2 = parseInt(h.slice(0, 2), 16);
      const g2 = parseInt(h.slice(2, 4), 16);
      const b2 = parseInt(h.slice(4, 6), 16);
      const lum = (0.299 * r2 + 0.587 * g2 + 0.114 * b2) / 255;
      return lum > 0.6 ? "111111" : "FFFFFF";
    };
    const wb = XLSXStyle.utils.book_new();
    const entries = cls.courses.map((course) => ({ course, cls }));
    const { rows: sumRows, ganttStartCol, weeks, courseWeekHits } = buildSummaryData(entries, cls);
    const sumWs = XLSXStyle.utils.aoa_to_sheet(sumRows);
    applySummaryStyles(
      sumWs,
      sumRows,
      ganttStartCol,
      weeks,
      courseWeekHits,
      entries,
      hexClean,
      textColorFor,
    );
    XLSXStyle.utils.book_append_sheet(wb, sumWs, "Summary");

    const data = buildSheet(cls);
    const clsDates = classDatesFor(cls, state);
    const ttWs = XLSXStyle.utils.aoa_to_sheet(data);
    applyTimetableStyles(ttWs, data, cls, clsDates, hexClean, textColorFor);
    XLSXStyle.utils.book_append_sheet(wb, ttWs, "Timetable");

    const cleanName =
      cls.name
        .replace(/[\\/?*[\]:]/g, "_")
        .slice(0, 20)
        .trim() || "Class";
    XLSXStyle.writeFile(wb, `class_timetable_${cleanName}.xlsx`);
  };

  // Export a single course's timetable as Summary + Timetable sheets
  const exportCourseTimetable = (cls: ClassData, course: Course) => {
    const hexClean = (h: string) =>
      (h || "").replace("#", "").padStart(6, "0").slice(-6).toUpperCase();
    const textColorFor = (hex: string) => {
      const h = hexClean(hex);
      const r2 = parseInt(h.slice(0, 2), 16);
      const g2 = parseInt(h.slice(2, 4), 16);
      const b2 = parseInt(h.slice(4, 6), 16);
      const lum = (0.299 * r2 + 0.587 * g2 + 0.114 * b2) / 255;
      return lum > 0.6 ? "111111" : "FFFFFF";
    };
    const wb = XLSXStyle.utils.book_new();
    const entries = [{ course, cls }];
    const { rows: sumRows, ganttStartCol, weeks, courseWeekHits } = buildSummaryData(entries, cls);
    const sumWs = XLSXStyle.utils.aoa_to_sheet(sumRows);
    applySummaryStyles(
      sumWs,
      sumRows,
      ganttStartCol,
      weeks,
      courseWeekHits,
      entries,
      hexClean,
      textColorFor,
    );
    XLSXStyle.utils.book_append_sheet(wb, sumWs, "Summary");

    // Course-only timetable sheet
    const clsDates = classDatesFor(cls, state);
    const courseHeader = ["Week", "Day / Date", ...state.slots.map(slotLabel)];
    const courseRows: (string | number)[][] = [courseHeader];
    clsDates.forEach((date) => {
      const { weekday, date: dstr } = dayLabel(date);
      const weekIndex = weekIndexForDate(cls, state, date);
      const row: (string | number)[] = [`W${weekIndex}`, `${weekday} ${dstr}`];
      state.slots.forEach((sl, i) => {
        if (sl.isBreak) {
          row.push("Break");
          return;
        }
        const cell = cls.grid[`${date}-${i}`];
        if (cell?.kind === "course" && cell.courseId === course.id) {
          const P2 = Math.max(0, course.practicalHours ?? 0);
          const LT2 =
            Math.max(0, course.lectureHours ?? 0) + Math.max(0, course.tutorialHours ?? 0);
          const isPractical = P2 > 0 && (LT2 === 0 || (course.durationSlots ?? 1) >= 2);
          const code = isPractical ? `${course.name}_P` : course.name;
          row.push(`${code} (${course.faculty})`);
        } else {
          row.push("");
        }
      });
      courseRows.push(row);
    });
    const ttWs = XLSXStyle.utils.aoa_to_sheet(courseRows);
    const numCols = courseHeader.length;
    const bdr = { style: "thin", color: { rgb: "CCCCCC" } } as const;
    const bb = { top: bdr, bottom: bdr, left: bdr, right: bdr };
    (ttWs as unknown as Record<string, unknown>)["!cols"] = Array.from(
      { length: numCols },
      (_, i) => ({ wch: i === 0 ? 8 : i === 1 ? 18 : 20 }),
    );
    (ttWs as unknown as Record<string, unknown>)["!rows"] = courseRows.map((_, i) => ({
      hpt: i === 0 ? 24 : 32,
    }));
    for (let r = 0; r < courseRows.length; r++) {
      for (let c = 0; c < numCols; c++) {
        const addr = XLSXStyle.utils.encode_cell({ r, c });
        if (!ttWs[addr]) ttWs[addr] = { t: "s", v: "" };
        const cs: Record<string, unknown> = {
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: bb,
          font: { name: "Calibri", sz: 11 },
        };
        if (r === 0 || c === 0 || c === 1) {
          cs.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } };
          cs.fill = { patternType: "solid", fgColor: { rgb: "0D0D0D" } };
        } else {
          const date = clsDates[r - 1];
          const slotIdx = c - 2;
          const sl = state.slots[slotIdx];
          const cell = cls.grid[`${date}-${slotIdx}`];
          if (sl?.isBreak) {
            cs.fill = { patternType: "solid", fgColor: { rgb: "FDE68A" } };
            cs.font = { name: "Calibri", sz: 11, italic: true, color: { rgb: "78350F" } };
          } else if (cell?.kind === "course" && cell.courseId === course.id) {
            const bg = hexClean(course.color ?? "#DDDDDD");
            cs.fill = { patternType: "solid", fgColor: { rgb: bg } };
            cs.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: textColorFor(bg) } };
          } else if (cell?.kind === "blocked") {
            cs.fill = { patternType: "solid", fgColor: { rgb: "F3F4F6" } };
            cs.font = { name: "Calibri", sz: 11, color: { rgb: "9CA3AF" } };
          }
        }
        (ttWs[addr] as { s?: unknown }).s = cs;
      }
    }
    (ttWs as unknown as Record<string, unknown>)["!freeze"] = { xSplit: 2, ySplit: 1 };
    XLSXStyle.utils.book_append_sheet(wb, ttWs, "Timetable");

    const cleanName =
      course.name
        .replace(/[\\/?*[\]:]/g, "_")
        .slice(0, 20)
        .trim() || "Course";
    XLSXStyle.writeFile(wb, `course_timetable_${cleanName}.xlsx`);
  };

  const buildFacultyWorkbook = (faculty: string) => {
    const refClass = activeClass || state.classes[0];
    if (!refClass) return null;

    const wb = XLSXStyle.utils.book_new();
    const hexClean = (h: string) =>
      (h || "").replace("#", "").padStart(6, "0").slice(-6).toUpperCase();
    const textColorFor = (hex: string) => {
      const h = hexClean(hex);
      const r = parseInt(h.slice(0, 2), 16);
      const g = parseInt(h.slice(2, 4), 16);
      const b = parseInt(h.slice(4, 6), 16);
      const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      return lum > 0.6 ? "111111" : "FFFFFF";
    };
    // Build Summary sheet (all courses this faculty teaches)
    const facEntries: Array<{ course: Course; cls: ClassData }> = [];
    state.classes.forEach((cls) => {
      cls.courses.forEach((c) => {
        if (c.faculty) {
          const names = c.faculty.split(",").map((f) => f.trim().toLowerCase());
          if (names.includes(faculty.toLowerCase())) facEntries.push({ course: c, cls });
        }
      });
    });
    if (facEntries.length > 0) {
      const {
        rows: sumRows,
        ganttStartCol,
        weeks,
        courseWeekHits,
      } = buildSummaryData(facEntries, refClass);
      const sumWs = XLSXStyle.utils.aoa_to_sheet(sumRows);
      applySummaryStyles(
        sumWs,
        sumRows,
        ganttStartCol,
        weeks,
        courseWeekHits,
        facEntries,
        hexClean,
        textColorFor,
      );
      XLSXStyle.utils.book_append_sheet(wb, sumWs, "Summary");
    }

    const border = { style: "thin", color: { rgb: "CCCCCC" } } as const;
    const baseBorders = { top: border, bottom: border, left: border, right: border };

    const workingDates = daysBetween(state.fromDate, state.toDate);

    const header = [
      "Week",
      "Day / Date",
      ...state.slots.map((s) => (s.isBreak ? "Break" : slotLabel(s))),
    ];
    const rows: (string | number)[][] = [header];

    workingDates.forEach((date) => {
      const { weekday, date: dstr } = dayLabel(date);
      const weekIndex = weekIndexForDate(refClass, state, date);
      const row: (string | number)[] = [`W${weekIndex}`, `${weekday} ${dstr}`];

      state.slots.forEach((sl, i) => {
        if (sl.isBreak) {
          row.push("Break");
          return;
        }
        const matches: string[] = [];
        const commonMap = new Map<string, { classNames: string[]; courseName: string }>();
        const individualMatches: string[] = [];

        state.classes.forEach((cls) => {
          const cell = cls.grid[`${date}-${i}`];
          if (cell?.kind === "course") {
            const course = cls.courses.find((x) => x.id === cell.courseId);
            if (course && course.faculty) {
              const names = course.faculty.split(",").map((f) => f.trim().toLowerCase());
              if (names.includes(faculty.toLowerCase())) {
                if (course.common) {
                  const key = course.name.trim().toLowerCase();
                  const existing = commonMap.get(key);
                  if (existing) {
                    if (!existing.classNames.includes(cls.name)) {
                      existing.classNames.push(cls.name);
                    }
                  } else {
                    commonMap.set(key, { classNames: [cls.name], courseName: course.name });
                  }
                } else {
                  individualMatches.push(`${cls.name} - ${course.name}`);
                }
              }
            }
          }
        });

        // Format common groups
        commonMap.forEach((val) => {
          matches.push(`${val.classNames.join(", ")} - ${val.courseName}`);
        });
        // Add individual ones
        matches.push(...individualMatches);

        row.push(matches.join(" / "));
      });
      rows.push(row);
    });

    const ws = XLSXStyle.utils.aoa_to_sheet(rows);
    const numCols = header.length;
    ws["!cols"] = Array.from({ length: numCols }, (_, i) => ({
      wch: i === 0 ? 8 : i === 1 ? 18 : 22,
    }));
    ws["!rows"] = rows.map((_, i) => ({ hpt: i === 0 ? 24 : 32 }));

    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < numCols; c++) {
        const addr = XLSXStyle.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { t: "s", v: "" };
        const cellStyle: Record<string, unknown> = {
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: baseBorders,
          font: { name: "Calibri", sz: 11 },
        };

        if (r === 0 || c === 0 || c === 1) {
          cellStyle.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } };
          cellStyle.fill = { patternType: "solid", fgColor: { rgb: "0D0D0D" } };
        } else {
          const date = workingDates[r - 1];
          const slotIdx = c - 2;
          const sl = state.slots[slotIdx];
          if (sl?.isBreak) {
            cellStyle.fill = { patternType: "solid", fgColor: { rgb: "FDE68A" } };
            cellStyle.font = { name: "Calibri", sz: 11, italic: true, color: { rgb: "78350F" } };
          } else {
            let matchColor = "";
            state.classes.forEach((cls) => {
              const cell = cls.grid[`${date}-${slotIdx}`];
              if (cell?.kind === "course") {
                const course = cls.courses.find((x) => x.id === cell.courseId);
                if (course && course.faculty) {
                  const names = course.faculty.split(",").map((f) => f.trim().toLowerCase());
                  if (names.includes(faculty.toLowerCase())) {
                    matchColor = course.color;
                  }
                }
              }
            });
            if (matchColor) {
              const bg = hexClean(matchColor);
              cellStyle.fill = { patternType: "solid", fgColor: { rgb: bg } };
              cellStyle.font = {
                name: "Calibri",
                sz: 11,
                bold: true,
                color: { rgb: textColorFor(bg) },
              };
            }
          }
        }
        (ws[addr] as { s?: unknown }).s = cellStyle;
      }
    }
    ws["!freeze"] = { xSplit: 2, ySplit: 1 };

    let cleanName = faculty.replace(new RegExp("[\\\\/?*\\[\\]:]", "g"), " ").slice(0, 31).trim();
    if (!cleanName) cleanName = "Faculty";

    XLSXStyle.utils.book_append_sheet(wb, ws, cleanName);

    const filename = `faculty_timetable_${cleanName.replace(/\s+/g, "_")}.xlsx`;
    return { wb, filename };
  };

  const exportSingleFacultyExcel = (faculty: string) => {
    const res = buildFacultyWorkbook(faculty);
    if (!res) {
      alert("No classes available to export.");
      return;
    }
    XLSXStyle.writeFile(res.wb, res.filename);
  };

  const exportAllFacultiesZip = async () => {
    const faculties = new Set<string>();
    state.classes.forEach((cls) => {
      cls.courses.forEach((c) => {
        if (c.faculty) {
          c.faculty
            .split(",")
            .map((f) => f.trim())
            .filter(Boolean)
            .forEach((f) => faculties.add(f));
        }
      });
    });
    const facList = Array.from(faculties).sort();

    if (facList.length === 0) {
      alert("No faculties found in any class to export.");
      return;
    }

    const zip = new JSZip();
    facList.forEach((faculty) => {
      const res = buildFacultyWorkbook(faculty);
      if (res) {
        const wbout = XLSXStyle.write(res.wb, { bookType: "xlsx", type: "array" });
        zip.file(res.filename, wbout);
      }
    });

    const content = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(content);
    const a = document.createElement("a");
    a.href = url;
    a.download = `all_faculties_timetables_${state.fromDate}_to_${state.toDate}.zip`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Export School Workload sheet
  const exportSchoolWorkload = () => {
    if (state.classes.length === 0) {
      alert("No classes to export.");
      return;
    }
    const wb = XLSXStyle.utils.book_new();
    const bdr = { style: "thin", color: { rgb: "CCCCCC" } } as const;
    const bb = { top: bdr, bottom: bdr, left: bdr, right: bdr };

    const parseCourseCodeAndName = (fullName: string) => {
      const trimmed = (fullName || "").trim();
      const m = trimmed.match(/^([A-Za-z0-9]+(?:\s*[0-9]+)?)\s*[\s–—:-]\s*(.+)$/);
      if (m) {
        return { code: m[1].trim(), name: m[2].trim() };
      }
      const parts = trimmed.split(/\s+/);
      if (parts.length > 1 && /^[A-Za-z0-9]+$/.test(parts[0])) {
        return { code: parts[0], name: parts.slice(1).join(" ") };
      }
      return { code: trimmed, name: trimmed };
    };

    const workingDates = daysBetween(state.fromDate, state.toDate);
    const dateHeaders = workingDates.map((d) => {
      const { weekday, date: dstr } = dayLabel(d);
      return `${weekday} ${dstr}`;
    });

    const HEADER = [
      "#",
      "Class",
      "Course Code",
      "Course Name",
      "Faculty",
      "L",
      "T",
      "P",
      "C",
      "Starting Date",
      "Ending Date",
      ...dateHeaders,
    ];

    const META_COL_COUNT = 11;

    const rows: (string | number)[][] = [HEADER];
    type CellMeta =
      | { type: "inactive" }
      | { type: "free" }
      | { type: "partial"; freeHours: string }
      | { type: "blocked" };

    const rowMetaList: CellMeta[][] = [];
    let globalIdx = 0;

    const nonBreakSlots = state.slots
      .map((sl, idx) => ({ sl, idx, pNum: periodNumberFor(idx) }))
      .filter((s) => !s.sl.isBreak);
    const totalNonBreak = nonBreakSlots.length;

    state.classes.forEach((cls) => {
      const clsState = stateForClass(cls, state);
      cls.courses.forEach((course) => {
        globalIdx++;
        const L = Math.max(0, course.lectureHours ?? 0);
        const T = Math.max(0, course.tutorialHours ?? 0);
        const P = Math.max(0, course.practicalHours ?? 0);
        const C = Math.max(0, course.credits ?? 0);

        const range = effectiveCourseRange(course, clsState);
        const startDate = range?.from || clsState.fromDate;
        const endDate = range?.to || clsState.toDate;

        const { code, name: courseName } = parseCourseCodeAndName(course.name);
        const facName = course.faculty || "";

        const row: (string | number)[] = [
          globalIdx,
          cls.name,
          code,
          courseName,
          facName,
          L,
          T,
          P,
          C,
          startDate,
          endDate,
        ];

        const rowCellMeta: CellMeta[] = [];

        workingDates.forEach((date) => {
          if (date < startDate || date > endDate || course.disabled) {
            row.push("-");
            rowCellMeta.push({ type: "inactive" });
            return;
          }

          if (!facName) {
            row.push("Free");
            rowCellMeta.push({ type: "free" });
            return;
          }

          // Check free periods for faculty on this date across all classes
          const facLower = facName.toLowerCase();
          const facNames = facLower
            .split(",")
            .map((f) => f.trim())
            .filter(Boolean);

          const freePNums: number[] = [];

          nonBreakSlots.forEach(({ idx: slotIdx, pNum }) => {
            let isBusy = false;
            for (const c of state.classes) {
              const cell = c.grid[`${date}-${slotIdx}`];
              if (cell?.kind === "course") {
                const crs = c.courses.find((x) => x.id === cell.courseId);
                if (crs && crs.faculty) {
                  const names = crs.faculty.split(",").map((f) => f.trim().toLowerCase());
                  if (facNames.some((fn) => names.includes(fn))) {
                    isBusy = true;
                    break;
                  }
                }
              }
            }
            if (!isBusy) {
              freePNums.push(pNum);
            }
          });

          if (freePNums.length === totalNonBreak) {
            row.push("Free");
            rowCellMeta.push({ type: "free" });
          } else if (freePNums.length > 0) {
            const freeHoursStr = freePNums.join(", ");
            row.push(freeHoursStr);
            rowCellMeta.push({ type: "partial", freeHours: freeHoursStr });
          } else {
            row.push("Blocked");
            rowCellMeta.push({ type: "blocked" });
          }
        });

        rows.push(row);
        rowMetaList.push(rowCellMeta);
      });
    });

    const ws = XLSXStyle.utils.aoa_to_sheet(rows);
    const numCols = HEADER.length;

    const metaWidths = [5, 20, 16, 26, 20, 5, 5, 5, 5, 12, 12];
    const dateWidths = workingDates.map(() => 14);

    (ws as unknown as Record<string, unknown>)["!cols"] = [
      ...metaWidths.map((wch) => ({ wch })),
      ...dateWidths.map((wch) => ({ wch })),
    ];
    (ws as unknown as Record<string, unknown>)["!rows"] = rows.map((_, i) => ({
      hpt: i === 0 ? 26 : 22,
    }));
    (ws as unknown as Record<string, unknown>)["!freeze"] = { xSplit: META_COL_COUNT, ySplit: 1 };

    const headerStyle: Record<string, unknown> = {
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: bb,
      font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: "0D0D0D" } },
    };

    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < numCols; c++) {
        const addr = XLSXStyle.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { t: "s", v: "" };
        let cs: Record<string, unknown>;

        if (r === 0) {
          cs = headerStyle;
        } else {
          if (c < META_COL_COUNT) {
            cs = {
              alignment: {
                horizontal: c === 0 || (c >= 5 && c <= 10) ? "center" : "left",
                vertical: "center",
                wrapText: c === 3 || c === 4,
              },
              border: bb,
              font: {
                name: "Calibri",
                sz: 11,
                bold: c <= 3,
                color: { rgb: "111111" },
              },
            };
          } else {
            const dateColIdx = c - META_COL_COUNT;
            const meta = rowMetaList[r - 1]?.[dateColIdx];

            if (meta?.type === "free") {
              cs = {
                alignment: { horizontal: "center", vertical: "center" },
                border: bb,
                font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "166534" } },
                fill: { patternType: "solid", fgColor: { rgb: "DCFCE7" } },
              };
            } else if (meta?.type === "partial") {
              cs = {
                alignment: { horizontal: "center", vertical: "center" },
                border: bb,
                font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "991B1B" } },
                fill: { patternType: "solid", fgColor: { rgb: "FEE2E2" } },
              };
            } else if (meta?.type === "blocked") {
              cs = {
                alignment: { horizontal: "center", vertical: "center" },
                border: bb,
                font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "7F1D1D" } },
                fill: { patternType: "solid", fgColor: { rgb: "FECACA" } },
              };
            } else {
              // inactive date
              cs = {
                alignment: { horizontal: "center", vertical: "center" },
                border: bb,
                font: { name: "Calibri", sz: 11, color: { rgb: "9CA3AF" } },
                fill: { patternType: "solid", fgColor: { rgb: "F3F4F6" } },
              };
            }
          }
        }
        (ws[addr] as { s?: unknown }).s = cs;
      }
    }

    XLSXStyle.utils.book_append_sheet(wb, ws, "School Workload");
    XLSXStyle.writeFile(wb, `school_workload_${state.fromDate}_to_${state.toDate}.xlsx`);
  };

  // Export School Faculty Workload Summary sheet
  const exportFacultyWorkloadSummary = () => {
    const faculties = new Set<string>();
    state.classes.forEach((cls) => {
      cls.courses.forEach((c) => {
        if (c.faculty) {
          c.faculty
            .split(",")
            .map((f) => f.trim())
            .filter(Boolean)
            .forEach((f) => faculties.add(f));
        }
      });
    });
    const facList = Array.from(faculties).sort();

    if (facList.length === 0) {
      alert("No faculties found to export.");
      return;
    }

    const wb = XLSXStyle.utils.book_new();
    const bdr = { style: "thin", color: { rgb: "CCCCCC" } } as const;
    const bb = { top: bdr, bottom: bdr, left: bdr, right: bdr };

    const workingDates = daysBetween(state.fromDate, state.toDate);
    const dateHeaders = workingDates.map((d) => {
      const { weekday, date: dstr } = dayLabel(d);
      return `${weekday} ${dstr}`;
    });

    const HEADER = ["#", "Faculty Name", ...dateHeaders];
    const META_COL_COUNT = 2;

    const rows: (string | number)[][] = [HEADER];
    type CellMeta = { type: "free" } | { type: "partial"; freeHours: string } | { type: "blocked" };

    const rowMetaList: CellMeta[][] = [];

    const nonBreakSlots = state.slots
      .map((sl, idx) => ({ sl, idx, pNum: periodNumberFor(idx) }))
      .filter((s) => !s.sl.isBreak);
    const totalNonBreak = nonBreakSlots.length;

    facList.forEach((facName, idx) => {
      const facLower = facName.toLowerCase();
      const facNames = facLower
        .split(",")
        .map((f) => f.trim())
        .filter(Boolean);

      const row: (string | number)[] = [idx + 1, facName];
      const rowCellMeta: CellMeta[] = [];

      workingDates.forEach((date) => {
        const freePNums: number[] = [];

        nonBreakSlots.forEach(({ idx: slotIdx, pNum }) => {
          let isBusy = false;
          for (const c of state.classes) {
            const cell = c.grid[`${date}-${slotIdx}`];
            if (cell?.kind === "course") {
              const crs = c.courses.find((x) => x.id === cell.courseId);
              if (crs && crs.faculty) {
                const names = crs.faculty.split(",").map((f) => f.trim().toLowerCase());
                if (facNames.some((fn) => names.includes(fn))) {
                  isBusy = true;
                  break;
                }
              }
            }
          }
          if (!isBusy) {
            freePNums.push(pNum);
          }
        });

        if (freePNums.length === totalNonBreak) {
          const freeHoursStr = freePNums.join(", ");
          row.push(freeHoursStr);
          rowCellMeta.push({ type: "free" });
        } else if (freePNums.length > 0) {
          const freeHoursStr = freePNums.join(", ");
          row.push(freeHoursStr);
          rowCellMeta.push({ type: "partial", freeHours: freeHoursStr });
        } else {
          row.push("None");
          rowCellMeta.push({ type: "blocked" });
        }
      });

      rows.push(row);
      rowMetaList.push(rowCellMeta);
    });

    const ws = XLSXStyle.utils.aoa_to_sheet(rows);
    const numCols = HEADER.length;

    const metaWidths = [5, 24];
    const dateWidths = workingDates.map(() => 14);

    (ws as unknown as Record<string, unknown>)["!cols"] = [
      ...metaWidths.map((wch) => ({ wch })),
      ...dateWidths.map((wch) => ({ wch })),
    ];
    (ws as unknown as Record<string, unknown>)["!rows"] = rows.map((_, i) => ({
      hpt: i === 0 ? 26 : 22,
    }));
    (ws as unknown as Record<string, unknown>)["!freeze"] = { xSplit: META_COL_COUNT, ySplit: 1 };

    const headerStyle: Record<string, unknown> = {
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: bb,
      font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: "0D0D0D" } },
    };

    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < numCols; c++) {
        const addr = XLSXStyle.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { t: "s", v: "" };
        let cs: Record<string, unknown>;

        if (r === 0) {
          cs = headerStyle;
        } else {
          if (c < META_COL_COUNT) {
            cs = {
              alignment: {
                horizontal: c === 0 ? "center" : "left",
                vertical: "center",
              },
              border: bb,
              font: {
                name: "Calibri",
                sz: 11,
                bold: c === 1,
                color: { rgb: "111111" },
              },
            };
          } else {
            const dateColIdx = c - META_COL_COUNT;
            const meta = rowMetaList[r - 1]?.[dateColIdx];

            if (meta?.type === "free") {
              cs = {
                alignment: { horizontal: "center", vertical: "center" },
                border: bb,
                font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "166534" } },
                fill: { patternType: "solid", fgColor: { rgb: "DCFCE7" } },
              };
            } else if (meta?.type === "partial") {
              cs = {
                alignment: { horizontal: "center", vertical: "center" },
                border: bb,
                font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "991B1B" } },
                fill: { patternType: "solid", fgColor: { rgb: "FEE2E2" } },
              };
            } else {
              // blocked / 0 free hours
              cs = {
                alignment: { horizontal: "center", vertical: "center" },
                border: bb,
                font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "7F1D1D" } },
                fill: { patternType: "solid", fgColor: { rgb: "FECACA" } },
              };
            }
          }
        }
        (ws[addr] as { s?: unknown }).s = cs;
      }
    }

    XLSXStyle.utils.book_append_sheet(wb, ws, "Faculty Workload Summary");
    XLSXStyle.writeFile(
      wb,
      `school_faculty_workload_summary_${state.fromDate}_to_${state.toDate}.xlsx`,
    );
  };

  // Export a classwise summary report: one sheet per class, listing all courses
  // with LTPC, faculty, total planned periods, placed periods, and missing.
  const exportSummaryReport = () => {
    if (state.classes.length === 0) {
      alert("No classes to export.");
      return;
    }
    const wb = XLSXStyle.utils.book_new();
    const bdr = { style: "thin", color: { rgb: "CCCCCC" } } as const;
    const bb = { top: bdr, bottom: bdr, left: bdr, right: bdr };

    const isoWeekKey = (iso: string) => {
      const d = utcDateFromIso(iso);
      if (!d) return "invalid";
      const day = (d.getUTCDay() + 6) % 7;
      const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() - day + 3));
      const first = new Date(Date.UTC(t.getUTCFullYear(), 0, 4));
      return `${t.getUTCFullYear()}-W${
        1 +
        Math.round(
          ((t.getTime() - first.getTime()) / 86400000 - 3 + ((first.getUTCDay() + 6) % 7)) / 7,
        )
      }`;
    };

    // Group classes by department (undefined/blank → "Unassigned")
    const deptMap = new Map<string, ClassData[]>();
    state.classes.forEach((cls) => {
      const dept = cls.department?.trim() || "Unassigned";
      const list = deptMap.get(dept) ?? [];
      list.push(cls);
      deptMap.set(dept, list);
    });

    // Header has a "Class" column between # and Course Code
    const HEADER = [
      "#",
      "Class",
      "Course Code",
      "Faculty",
      "L",
      "T",
      "P",
      "C",
      "Total Classes",
      "Placed Classes",
      "Missing",
    ];
    const COL_WIDTHS = [5, 22, 30, 22, 5, 5, 5, 5, 14, 14, 12];
    // Column indices for data checks
    const COL_TOTAL = 8;
    const COL_PLACED = 9;
    const COL_MISSING = 10;

    const headerStyle: Record<string, unknown> = {
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: bb,
      font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: "0D0D0D" } },
    };

    const usedSheetNames = new Set<string>();

    deptMap.forEach((classes, dept) => {
      const rows: (string | number)[][] = [HEADER];
      let globalIdx = 0;

      let deptTotalExpected = 0;
      let deptTotalPlaced = 0;

      classes.forEach((cls) => {
        const clsState = stateForClass(cls, state);
        const clsDates = classDatesFor(cls, state);
        const weekCount = new Set(clsDates.map(isoWeekKey)).size;

        cls.courses.forEach((course) => {
          globalIdx++;
          const L = Math.max(0, course.lectureHours ?? 0);
          const T = Math.max(0, course.tutorialHours ?? 0);
          const P = Math.max(0, course.practicalHours ?? 0);
          const C = Math.max(0, course.credits ?? 0);

          const totalT = courseTotalTarget(course, courseSemesterWeeks(course, clsState));
          const total = totalT > 0 ? totalT : courseWeeklyTarget(course) * weekCount;
          const placed = countCoursePeriodsInDates(cls.grid, course, state.slots, clsDates);
          const missing = Math.max(0, total - placed);

          deptTotalExpected += total > 0 ? total : 0;
          deptTotalPlaced += placed;

          rows.push([
            globalIdx,
            cls.name,
            course.name,
            course.faculty || "",
            L,
            T,
            P,
            C,
            total > 0 ? total : 0,
            placed,
            missing > 0 ? missing : "",
          ]);
        });
      });

      // Department totals row
      const deptMissing = Math.max(0, deptTotalExpected - deptTotalPlaced);
      rows.push([
        "",
        "TOTAL",
        "",
        "",
        "",
        "",
        "",
        "",
        deptTotalExpected,
        deptTotalPlaced,
        deptMissing > 0 ? deptMissing : "",
      ]);

      const ws = XLSXStyle.utils.aoa_to_sheet(rows);
      const numCols = HEADER.length;

      (ws as unknown as Record<string, unknown>)["!cols"] = COL_WIDTHS.map((wch) => ({ wch }));
      (ws as unknown as Record<string, unknown>)["!rows"] = rows.map((_, i) => ({
        hpt: i === 0 ? 26 : 20,
      }));
      (ws as unknown as Record<string, unknown>)["!freeze"] = { xSplit: 0, ySplit: 1 };

      const totalsRowIdx = rows.length - 1;

      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < numCols; c++) {
          const addr = XLSXStyle.utils.encode_cell({ r, c });
          if (!ws[addr]) ws[addr] = { t: "s", v: "" };
          let cs: Record<string, unknown>;

          if (r === 0) {
            cs = headerStyle;
          } else if (r === totalsRowIdx) {
            const isMissingTotals = c === COL_MISSING && (rows[r][c] as number) > 0;
            cs = {
              alignment: { horizontal: c <= 2 ? "left" : "center", vertical: "center" },
              border: bb,
              font: {
                name: "Calibri",
                sz: 11,
                bold: true,
                color: { rgb: isMissingTotals ? "B91C1C" : "FFFFFF" },
              },
              fill: {
                patternType: "solid",
                fgColor: { rgb: isMissingTotals ? "FEE2E2" : "374151" },
              },
            };
          } else {
            const missingVal = rows[r][COL_MISSING] as number | string;
            const isMissingCell =
              c === COL_MISSING && missingVal !== "" && (missingVal as number) > 0;
            const isPlacedCell = c === COL_PLACED;
            const total = rows[r][COL_TOTAL] as number;
            const placed2 = rows[r][COL_PLACED] as number;
            const complete = total > 0 && placed2 >= total;

            cs = {
              alignment: {
                horizontal: c <= 3 ? (c === 0 ? "center" : "left") : "center",
                vertical: "center",
                wrapText: c <= 3,
              },
              border: bb,
              font: {
                name: "Calibri",
                sz: 11,
                bold: c <= 2,
                color: isMissingCell
                  ? { rgb: "B91C1C" }
                  : isPlacedCell && complete
                    ? { rgb: "166534" }
                    : { rgb: "111111" },
              },
              fill: isMissingCell
                ? { patternType: "solid", fgColor: { rgb: "FEF2F2" } }
                : isPlacedCell && complete
                  ? { patternType: "solid", fgColor: { rgb: "DCFCE7" } }
                  : undefined,
            };
            if (!cs.fill) delete cs.fill;
          }
          (ws[addr] as { s?: unknown }).s = cs;
        }
      }

      // Safe unique sheet name from department name
      const rawName =
        dept
          .replace(/[\\/?*[\]:]/g, "_")
          .slice(0, 31)
          .trim() || "Unassigned";
      let sheetName = rawName;
      let n = 1;
      while (usedSheetNames.has(sheetName.toLowerCase())) {
        const suffix = ` (${n++})`;
        sheetName = rawName.slice(0, 31 - suffix.length) + suffix;
      }
      usedSheetNames.add(sheetName.toLowerCase());
      XLSXStyle.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSXStyle.writeFile(wb, `department_summary_${state.fromDate}_to_${state.toDate}.xlsx`);
  };

  const exportFacultyExcel = () => {
    const refClass = activeClass || state.classes[0];
    if (!refClass) {
      alert("No classes available to export.");
      return;
    }
    const wb = XLSXStyle.utils.book_new();
    const hexClean = (h: string) =>
      (h || "").replace("#", "").padStart(6, "0").slice(-6).toUpperCase();
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

    const facultiesMap = new Map<string, string>();
    state.classes.forEach((cls) => {
      cls.courses.forEach((c) => {
        if (c.faculty) {
          c.faculty
            .split(",")
            .map((f) => f.trim())
            .filter(Boolean)
            .forEach((f) => {
              const lower = f.toLowerCase();
              if (!facultiesMap.has(lower)) {
                facultiesMap.set(lower, f);
              }
            });
        }
      });
    });
    const sortedFaculties = Array.from(facultiesMap.values()).sort((a, b) => a.localeCompare(b));

    if (sortedFaculties.length === 0) {
      alert("No faculty assignments found to export.");
      return;
    }

    const workingDates = daysBetween(state.fromDate, state.toDate);
    const usedSheetNames = new Set<string>();

    sortedFaculties.forEach((faculty) => {
      const header = [
        "Week",
        "Day / Date",
        ...state.slots.map((s) => (s.isBreak ? "Break" : slotLabel(s))),
      ];
      const rows: (string | number)[][] = [header];

      workingDates.forEach((date) => {
        const { weekday, date: dstr } = dayLabel(date);
        const weekIndex = weekIndexForDate(refClass, state, date);
        const row: (string | number)[] = [`W${weekIndex}`, `${weekday} ${dstr}`];

        state.slots.forEach((sl, i) => {
          if (sl.isBreak) {
            row.push("Break");
            return;
          }
          const matches: string[] = [];
          const commonMap = new Map<string, { classNames: string[]; courseName: string }>();
          const individualMatches: string[] = [];

          state.classes.forEach((cls) => {
            const cell = cls.grid[`${date}-${i}`];
            if (cell?.kind === "course") {
              const course = cls.courses.find((x) => x.id === cell.courseId);
              if (course && course.faculty) {
                const names = course.faculty.split(",").map((f) => f.trim().toLowerCase());
                if (names.includes(faculty.toLowerCase())) {
                  if (course.common) {
                    const key = course.name.trim().toLowerCase();
                    const existing = commonMap.get(key);
                    if (existing) {
                      if (!existing.classNames.includes(cls.name)) {
                        existing.classNames.push(cls.name);
                      }
                    } else {
                      commonMap.set(key, { classNames: [cls.name], courseName: course.name });
                    }
                  } else {
                    individualMatches.push(`${cls.name} - ${course.name}`);
                  }
                }
              }
            }
          });

          // Format common groups
          commonMap.forEach((val) => {
            matches.push(`${val.classNames.join(", ")} - ${val.courseName}`);
          });
          // Add individual ones
          matches.push(...individualMatches);

          row.push(matches.join(" / "));
        });
        rows.push(row);
      });

      const ws = XLSXStyle.utils.aoa_to_sheet(rows);
      const numCols = header.length;
      ws["!cols"] = Array.from({ length: numCols }, (_, i) => ({
        wch: i === 0 ? 8 : i === 1 ? 18 : 22,
      }));
      ws["!rows"] = rows.map((_, i) => ({ hpt: i === 0 ? 24 : 32 }));

      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < numCols; c++) {
          const addr = XLSXStyle.utils.encode_cell({ r, c });
          if (!ws[addr]) ws[addr] = { t: "s", v: "" };
          const cellStyle: Record<string, unknown> = {
            alignment: { horizontal: "center", vertical: "center", wrapText: true },
            border: baseBorders,
            font: { name: "Calibri", sz: 11 },
          };

          if (r === 0 || c === 0 || c === 1) {
            cellStyle.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } };
            cellStyle.fill = { patternType: "solid", fgColor: { rgb: "0D0D0D" } };
          } else {
            const date = workingDates[r - 1];
            const slotIdx = c - 2;
            const sl = state.slots[slotIdx];
            if (sl?.isBreak) {
              cellStyle.fill = { patternType: "solid", fgColor: { rgb: "FDE68A" } };
              cellStyle.font = { name: "Calibri", sz: 11, italic: true, color: { rgb: "78350F" } };
            } else {
              let matchColor = "";
              state.classes.forEach((cls) => {
                const cell = cls.grid[`${date}-${slotIdx}`];
                if (cell?.kind === "course") {
                  const course = cls.courses.find((x) => x.id === cell.courseId);
                  if (course && course.faculty) {
                    const names = course.faculty.split(",").map((f) => f.trim().toLowerCase());
                    if (names.includes(faculty.toLowerCase())) {
                      matchColor = course.color;
                    }
                  }
                }
              });
              if (matchColor) {
                const bg = hexClean(matchColor);
                cellStyle.fill = { patternType: "solid", fgColor: { rgb: bg } };
                cellStyle.font = {
                  name: "Calibri",
                  sz: 11,
                  bold: true,
                  color: { rgb: textColorFor(bg) },
                };
              }
            }
          }
          (ws[addr] as { s?: unknown }).s = cellStyle;
        }
      }
      ws["!freeze"] = { xSplit: 2, ySplit: 1 };

      let cleanName = faculty.replace(new RegExp("[\\\\/?*\\[\\]:]", "g"), " ").slice(0, 31).trim();
      if (!cleanName) cleanName = "Faculty";

      let finalName = cleanName;
      let counter = 1;
      while (usedSheetNames.has(finalName.toLowerCase())) {
        const suffix = ` (${counter})`;
        finalName = cleanName.slice(0, 31 - suffix.length) + suffix;
        counter++;
      }
      usedSheetNames.add(finalName.toLowerCase());

      XLSXStyle.utils.book_append_sheet(wb, ws, finalName);
    });

    XLSXStyle.writeFile(wb, `faculty_timetables_${state.fromDate}_to_${state.toDate}.xlsx`);
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
    // aSc TimeTables "Data to Fill" import workbook. One sheet per class.
    // Lessons/week is the actual number of occupied timetable periods in each
    // week. LTPC is used only to split those placed periods into theory rows
    // (course code) and practical rows (coursecode_P), so totals match the
    // designed timetable exactly.
    const header = [
      "Teacher",
      "Class",
      "Group",
      "Subject",
      "Subject",
      "Length",
      "Lessons/week",
      "Available classrooms",
      "Cycle",
      "weight",
    ];
    const wb = XLSX.utils.book_new();
    state.classes.forEach((cls) => {
      const rows: (string | number)[][] = [header];
      const clsDates = classDatesFor(cls, state);
      const clsState = stateForClass(cls, state);
      cls.courses.forEach((course) => {
        const teacher = course.faculty || "";
        const className = cls.name;
        const group = "Entire class";
        const subjectName = course.name;
        const classroom = course.classroom || "";
        const placedPeriodsByWeek = new Map<number, number>();
        clsDates.forEach((date) => {
          state.slots.forEach((slot, slotIdx) => {
            if (slot.isBreak) return;
            const cell = cls.grid[`${date}-${slotIdx}`];
            if (cell?.kind !== "course" || cell.courseId !== course.id) return;
            const cycle =
              courseCycleForDate(course, clsState, date) ??
              (() => {
                // Fallback: week index relative to global timetable start
                const start = utcDateFromIso(classFromDate(cls, state));
                const cur = utcDateFromIso(date);
                if (!start || !cur) return 1;
                return Math.floor((cur.getTime() - start.getTime()) / (7 * 86400000)) + 1;
              })();
            placedPeriodsByWeek.set(cycle, (placedPeriodsByWeek.get(cycle) ?? 0) + 1);
          });
        });
        if (placedPeriodsByWeek.size === 0) return;

        const totalPlacedPeriods = Array.from(placedPeriodsByWeek.values()).reduce(
          (sum, value) => sum + value,
          0,
        );
        const theoryUnits = Math.max(0, (course.lectureHours ?? 0) + (course.tutorialHours ?? 0));
        const practicalUnits = Math.max(0, course.practicalHours ?? 0);
        const totalLtpUnits = theoryUnits + practicalUnits;
        const desiredTheoryPeriods =
          totalLtpUnits > 0
            ? Math.min(
                totalPlacedPeriods,
                Math.round((totalPlacedPeriods * theoryUnits) / totalLtpUnits),
              )
            : totalPlacedPeriods;

        const weekEntries = Array.from(placedPeriodsByWeek.entries()).sort(([a], [b]) => a - b);
        const theoryByWeek = new Map<number, number>();
        if (desiredTheoryPeriods > 0) {
          const weighted = weekEntries.map(([cycle, periods]) => {
            const exact =
              totalPlacedPeriods > 0 ? (periods * desiredTheoryPeriods) / totalPlacedPeriods : 0;
            const base = Math.min(periods, Math.floor(exact));
            return { cycle, periods, base, fraction: exact - base };
          });
          let assigned = weighted.reduce((sum, item) => sum + item.base, 0);
          weighted.forEach((item) => theoryByWeek.set(item.cycle, item.base));
          weighted
            .slice()
            .sort((a, b) => b.fraction - a.fraction || a.cycle - b.cycle)
            .forEach((item) => {
              if (assigned >= desiredTheoryPeriods) return;
              const current = theoryByWeek.get(item.cycle) ?? 0;
              if (current >= item.periods) return;
              theoryByWeek.set(item.cycle, current + 1);
              assigned++;
            });
        }

        const addRow = (subjectCode: string, cycle: number, lessons: number) => {
          if (lessons <= 0) return;
          const weight = Number((lessons / 18).toFixed(4));
          rows.push([
            teacher,
            className,
            group,
            subjectCode,
            subjectName,
            1,
            lessons,
            classroom,
            `W${cycle}`,
            weight,
          ]);
        };

        weekEntries.forEach(([cycle, placedPeriods]) => {
          const theoryLessons = Math.min(placedPeriods, theoryByWeek.get(cycle) ?? 0);
          const practicalLessons = Math.max(0, placedPeriods - theoryLessons);
          addRow(subjectName, cycle, theoryLessons);
          addRow(`${subjectName}_P`, cycle, practicalLessons);
        });
      });
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"] = [
        { wch: 22 },
        { wch: 10 },
        { wch: 14 },
        { wch: 14 },
        { wch: 32 },
        { wch: 8 },
        { wch: 14 },
        { wch: 22 },
        { wch: 8 },
        { wch: 10 },
        { wch: 10 },
      ];
      const sheetName = `Data to Fill ${cls.name}`.slice(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, `timetable_${state.fromDate}_to_${state.toDate}_ASC.xlsx`);
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
      setAutoFillReport(
        "Loaded .aadhi file. Course span values above the available periods were corrected to 1.",
      );
    } catch {
      alert("Could not read this .aadhi file.");
    }
  };

  // ---- Admin cross-check: load additional .aadhi files as reference and
  // report faculty conflicts between them and the currently open timetable.
  const loadAdminReference = async (file: File) => {
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      const loaded: SavedState | undefined = parsed?.state ?? parsed;
      if (!loaded || !Array.isArray(loaded.classes) || !Array.isArray(loaded.slots)) {
        alert(`"${file.name}" doesn't look like a valid .aadhi file.`);
        return;
      }
      const normalized = normalizeStateSnapshot(loaded);

      // Give imported classes unique IDs and check for clashing names
      const uniqueClasses = normalized.classes.map((cls, idx) => {
        const nameClash = state.classes.some((c) => c.name === cls.name);
        const name = nameClash ? `${cls.name} (${file.name.replace(/\.aadhi$/, "")})` : cls.name;
        return {
          ...cls,
          id: `k${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 7)}`,
          name,
        };
      });

      setState((prev) => ({
        ...prev,
        classes: [...prev.classes, ...uniqueClasses],
      }));

      setAdminRefs((prev) => [
        ...prev,
        {
          id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
          name: file.name,
          state: normalized,
        },
      ]);

      if (uniqueClasses.length > 0) {
        setActiveClassId(uniqueClasses[0].id);
      }

      setWarningsPanelOpen(true);
      setAutoFillReport(
        `Imported ${uniqueClasses.length} class(es) from "${file.name}" as active classes. Check warnings panel for overlaps.`,
      );
    } catch {
      alert(`Could not read "${file.name}".`);
    }
  };

  const adminReport = useMemo(() => {
    // Build (date, slot) → list of { origin, className, courseName, faculty }
    type Placement = { origin: string; className: string; courseName: string; faculty: string };
    const rows: Array<{ key: string; date: string; slotIdx: number; conflicts: Placement[] }> = [];
    if (!adminOpen || adminRefs.length === 0) return rows;
    const sources: Array<{ origin: string; state: State }> = [
      { origin: "This file", state },
      ...adminRefs.map((r) => ({ origin: r.name, state: r.state })),
    ];
    // Group placements by key
    const byKey = new Map<string, Placement[]>();
    sources.forEach(({ origin, state: st }) => {
      st.classes.forEach((cls) => {
        Object.entries(cls.grid).forEach(([key, cell]) => {
          if (cell.kind !== "course") return;
          const course = cls.courses.find((c) => c.id === cell.courseId);
          if (!course) return;
          getFaculties(course).forEach((faculty) => {
            const list = byKey.get(key) ?? [];
            list.push({
              origin,
              className: cls.name,
              courseName: course.name || course.id,
              faculty,
            });
            byKey.set(key, list);
          });
        });
      });
    });
    byKey.forEach((list, key) => {
      const byFaculty = new Map<string, Placement[]>();
      list.forEach((p) => {
        const b = byFaculty.get(p.faculty) ?? [];
        b.push(p);
        byFaculty.set(p.faculty, b);
      });
      const clashes: Placement[] = [];
      byFaculty.forEach((ps) => {
        // Only count as a clash if the same faculty appears in more than one
        // (origin,className) pair at this slot.
        const distinct = new Set(ps.map((p) => `${p.origin}⧫${p.className}`));
        if (distinct.size > 1) clashes.push(...ps);
      });
      if (clashes.length > 0) {
        const [d, sIdxStr] = key.split(/-(?=\d+$)/);
        rows.push({ key, date: d, slotIdx: parseInt(sIdxStr, 10), conflicts: clashes });
      }
    });
    rows.sort((a, b) => (a.date === b.date ? a.slotIdx - b.slotIdx : a.date.localeCompare(b.date)));
    return rows;
  }, [adminOpen, adminRefs, state]);

  const cellDisplay = (cell: Cell | undefined, courses: Course[]) => {
    if (!cell || cell.kind === "empty") return { text: "", bg: "#fff", fg: "#94a3b8" };
    if (cell.kind === "break") return { text: cell.label, bg: "#fef3c7", fg: "#92400e" };
    if (cell.kind === "blocked") return { text: cell.label, bg: "#e5e7eb", fg: "#374151" };
    const course = courses.find((c) => c.id === cell.courseId);
    return {
      text: course
        ? `${course.name}\n${course.faculty}${course.classroom ? ` · ${course.classroom}` : ""}`
        : "?",
      bg: course?.color ?? "#ddd",
      fg: "#1f2937",
    };
  };

  const hasConflicts = conflicts.size > 0;
  const activeVisibleCourseSlots = useMemo(() => {
    if (!activeClass) return 0;
    return dates.reduce((sum, date) => {
      return (
        sum +
        state.slots.reduce((slotSum, _, slotIdx) => {
          const cell = activeClass.grid[`${date}-${slotIdx}`];
          return slotSum + (cell?.kind === "course" ? 1 : 0);
        }, 0)
      );
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
      const wk =
        1 +
        Math.round(
          ((t.getTime() - first.getTime()) / 86400000 - 3 + ((first.getUTCDay() + 6) % 7)) / 7,
        );
      return `${t.getUTCFullYear()}-W${wk}`;
    };
    const countPlaced = (cls: ClassData) => {
      const clsDates = classDatesFor(cls, state);
      let n = 0;
      cls.courses.forEach((course) => {
        n += countCoursePeriodsInDates(cls.grid, course, state.slots, clsDates);
      });
      return n;
    };
    const planFor = (cls: ClassData) => {
      const clsState = stateForClass(cls, state);
      const clsDates = classDatesFor(cls, state);
      const weekCount = new Set(clsDates.map(isoWeekKey)).size;
      return cls.courses.reduce((sum, c) => {
        if (c.disabled) return sum;
        const totalT = courseTotalTarget(c, courseSemesterWeeks(c, clsState));
        const requested = totalT > 0 ? totalT : courseWeeklyTarget(c) * weekCount;
        const capacity = countCourseRuleCapacity(c, state.slots, clsDates);
        if (requested <= 0) return sum + capacity;
        return sum + requested;
      }, 0);
    };
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

  // Per-course assigned session counts for the active class.
  const coursePlacementCounts = useMemo(() => {
    const map = new Map<string, number>();
    if (!activeClass) return map;
    const clsDates = classDatesFor(activeClass, state);
    activeClass.courses.forEach((course) => {
      const count = countCoursePeriodsInDates(activeClass.grid, course, state.slots, clsDates);
      if (count > 0) map.set(course.id, count);
    });
    return map;
  }, [activeClass, state.slots, state.fromDate, state.toDate]);
  const globalFaculties = useMemo(() => {
    const set = new Set<string>();
    state.classes.forEach((cls) => {
      cls.courses.forEach((c) => {
        if (c.faculty) {
          c.faculty
            .split(",")
            .map((f) => f.trim())
            .filter(Boolean)
            .forEach((f) => set.add(f));
        }
      });
    });
    return Array.from(set).sort();
  }, [state.classes]);

  const selectedFacultyCourses = useMemo(() => {
    if (!selectedFaculty) return [];
    const commonGroups = new Map<
      string,
      { classNames: string[]; courseName: string; weeklyPeriods: number }
    >();
    const individualList: Array<{ className: string; courseName: string; weeklyPeriods: number }> =
      [];

    state.classes.forEach((cls) => {
      cls.courses.forEach((c) => {
        if (c.faculty) {
          const names = c.faculty.split(",").map((f) => f.trim().toLowerCase());
          if (names.includes(selectedFaculty.toLowerCase())) {
            if (c.common) {
              const key = c.name.trim().toLowerCase();
              const existing = commonGroups.get(key);
              if (existing) {
                if (!existing.classNames.includes(cls.name)) {
                  existing.classNames.push(cls.name);
                }
              } else {
                commonGroups.set(key, {
                  classNames: [cls.name],
                  courseName: c.name,
                  weeklyPeriods: courseWeeklyTarget(c),
                });
              }
            } else {
              individualList.push({
                className: cls.name,
                courseName: c.name,
                weeklyPeriods: courseWeeklyTarget(c),
              });
            }
          }
        }
      });
    });

    const groupedCommon = Array.from(commonGroups.values()).map((g) => ({
      className: g.classNames.join(", "),
      courseName: g.courseName + " (Common)",
      weeklyPeriods: g.weeklyPeriods,
    }));

    return [...groupedCommon, ...individualList];
  }, [selectedFaculty, state.classes]);

  const commitFacultyRename = (oldName: string, newName: string) => {
    if (!newName || oldName === newName) {
      setEditingFaculty(null);
      return;
    }
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        return {
          ...cls,
          courses: cls.courses.map((course) => {
            if (!course.faculty) return course;
            const names = course.faculty.split(",").map((f) => f.trim());
            const updatedNames = names.map((name) => {
              return name.toLowerCase() === oldName.toLowerCase() ? newName : name;
            });
            return {
              ...course,
              faculty: updatedNames.join(", "),
            };
          }),
        };
      }),
    }));
    setEditingFaculty(null);
  };
  const handleWarningClick = (warning: (typeof conflictDetailsList)[0]) => {
    setActiveClassId(warning.classId);
    setPendingScroll({
      classId: warning.classId,
      date: warning.date,
      slotIdx: warning.slotIdx,
    });
    if (window.innerWidth < 768) {
      setWarningsPanelOpen(false);
    }
  };

  return (
    <div
      className="min-h-screen w-full bg-[#f5f3ee] text-[#2d2d2d]"
      style={{ fontFamily: "'Manrope', system-ui, sans-serif" }}
      onMouseLeave={() => setIsPainting(false)}
      onMouseMove={(e) => {
        if (armedTool?.kind === "course") setCursorPos({ x: e.clientX, y: e.clientY });
      }}
    >
      {armedTool?.kind === "course" &&
        cursorPos &&
        activeClass &&
        (() => {
          const course = activeClass.courses.find((c) => c.id === armedTool.courseId);
          if (!course) return null;
          const clsState = stateForClass(activeClass, state);
          const total = courseTotalTarget(course, courseSemesterWeeks(course, clsState));
          const placed = coursePlacementCounts.get(course.id) ?? 0;
          const remaining = Math.max(0, total - placed);
          return (
            <div
              className="pointer-events-none fixed z-[9999] flex items-center gap-2 border-2 border-[#0d0d0d] bg-white px-2 py-1 text-[10px] font-bold uppercase tracking-wider shadow-lg"
              style={{ left: cursorPos.x + 16, top: cursorPos.y + 16 }}
            >
              <span
                className="inline-block h-2.5 w-2.5 border border-[#0d0d0d]"
                style={{ backgroundColor: course.color }}
              />
              <span>{course.name}</span>
              <span className="text-[#0d0d0d]/60">·</span>
              <span className={remaining === 0 ? "text-emerald-700" : "text-red-700"}>
                {remaining} left
              </span>
              <span className="text-[#0d0d0d]/60">/ {total}</span>
            </div>
          );
        })()}
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
                    <div
                      key={cls.id}
                      className="flex items-stretch gap-1"
                      onContextMenu={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        setClassContextMenu({
                          classId: cls.id,
                          x: e.clientX,
                          y: e.clientY,
                        });
                      }}
                    >
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
                  <div
                    key={c.id}
                    className="border border-[#0d0d0d]/30 bg-white"
                    onContextMenu={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      setCourseContextMenu({
                        courseId: c.id,
                        x: e.clientX,
                        y: e.clientY,
                      });
                    }}
                  >
                    <div
                      className={
                        "flex items-center gap-2 border-b border-[#0d0d0d]/10 px-3 py-2 " +
                        (c.disabled ? "opacity-60" : "")
                      }
                    >
                      <label
                        className="h-3.5 w-3.5 shrink-0 cursor-pointer border border-[#0d0d0d]/20 relative rounded overflow-hidden"
                        title="Change course color"
                        onClick={(e) => e.stopPropagation()}
                      >
                        <span className="absolute inset-0" style={{ backgroundColor: c.color }} />
                        <input
                          type="color"
                          value={c.color.startsWith("#") ? c.color : hslToHex(c.color)}
                          onChange={(e) => updateCourse(c.id, { color: e.target.value })}
                          className="absolute inset-0 opacity-0 cursor-pointer w-full h-full p-0 border-none"
                        />
                      </label>
                      <input
                        value={c.name}
                        onChange={(e) => updateCourse(c.id, { name: e.target.value })}
                        placeholder="Course"
                        className={
                          "min-w-0 flex-1 bg-transparent text-sm font-bold outline-none " +
                          (c.disabled ? "line-through" : "")
                        }
                      />
                      {c.common && (
                        <span className="shrink-0 rounded bg-sky-100 border border-sky-300 px-1 py-0.5 text-[8px] font-bold uppercase tracking-wider text-sky-800">
                          Common
                        </span>
                      )}
                      <button
                        onClick={() =>
                          updateCourse(c.id, { disabled: c.disabled ? undefined : true })
                        }
                        title={
                          c.disabled
                            ? "Course is disabled — excluded from Fill by Rules. Click to enable."
                            : "Disable this course (excluded from Fill by Rules)."
                        }
                        className={
                          "flex items-center gap-1 border px-1.5 py-0.5 text-[10px] font-bold uppercase tracking-wider " +
                          (c.disabled
                            ? "border-[#0d0d0d]/40 bg-[#0d0d0d] text-white"
                            : "border-[#0d0d0d]/30 bg-white text-[#2d2d2d]/70 hover:border-[#0d0d0d]")
                        }
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        {c.disabled ? "Off" : "On"}
                      </button>
                      <button
                        onClick={() => clearCoursePlacements(c.id)}
                        title="Clear all placed sessions of this course from this class"
                        className="border border-red-200 bg-red-50 px-1.5 py-0.5 text-[10px] font-bold uppercase tracking-wider text-red-700 hover:bg-red-100 hover:border-red-300"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        Clear
                      </button>
                      <button
                        onClick={() => removeCourse(c.id)}
                        className="text-xs text-[#2d2d2d]/40 hover:text-red-600"
                        aria-label="Remove course"
                      >
                        ×
                      </button>
                    </div>
                    <div className="grid grid-cols-2 gap-2 px-3 pt-2">
                      <input
                        value={c.faculty}
                        onChange={(e) => updateCourse(c.id, { faculty: e.target.value })}
                        placeholder="Faculty (comma-sep for multiple)"
                        title="One or more faculty names. Separate co-teachers with commas (e.g. Dr. Smith, Dr. Jones). Any shared name across classes counts as a conflict."
                        className="min-w-0 border-b border-dashed border-[#0d0d0d]/20 bg-transparent text-xs text-[#2d2d2d]/70 outline-none focus:border-[#0d0d0d]"
                      />
                      <input
                        value={c.classroom ?? ""}
                        onChange={(e) => updateCourse(c.id, { classroom: e.target.value })}
                        placeholder="Room"
                        title="Classroom / room name"
                        className="min-w-0 border-b border-dashed border-[#0d0d0d]/20 bg-transparent text-xs text-[#2d2d2d]/70 outline-none focus:border-[#0d0d0d]"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      />
                    </div>
                    <div className="flex flex-wrap items-center gap-x-3 gap-y-2 px-3 py-2">
                      {(
                        [
                          ["lectureHours", "L", "Lecture hours per week"],
                          ["tutorialHours", "T", "Tutorial hours per week"],
                          ["practicalHours", "P", "Practical sessions per week"],
                          ["credits", "C", "Credits (informational)"],
                        ] as [keyof Course, string, string][]
                      ).map(([field, label, tip]) => (
                        <label
                          key={field as string}
                          title={tip}
                          className="flex items-center gap-1 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
                          style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                        >
                          <span>{label}</span>
                          <input
                            type="number"
                            min={0}
                            value={(c[field] as number | undefined) ?? 0}
                            onChange={(e) => {
                              const n = Math.max(0, parseInt(e.target.value || "0", 10));
                              updateCourse(c.id, {
                                [field]: n > 0 ? n : undefined,
                              } as Partial<Course>);
                            }}
                            className="w-10 border border-[#0d0d0d]/20 bg-white px-1 py-0.5 text-center text-xs"
                          />
                        </label>
                      ))}
                      <label
                        className="flex items-center gap-1 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        <select
                          value={c.durationSlots}
                          onChange={(e) =>
                            updateCourse(c.id, {
                              durationSlots: cleanDurationSlots(e.target.value, state.slots),
                            })
                          }
                          className="border border-[#0d0d0d]/20 bg-white px-1 py-0.5 text-xs text-[#2d2d2d] outline-none"
                        >
                          {Array.from(
                            { length: Math.min(7, Math.max(1, nonBreakCount(state.slots))) },
                            (_, idx) => idx + 1,
                          ).map((num) => (
                            <option key={num} value={num}>
                              {num}
                            </option>
                          ))}
                        </select>
                        <span title="Consecutive classes at a stretch in a day">at a stretch</span>
                      </label>
                      <label
                        title="Sessions per week (auto-fill target). Derived from LTPC (L+T+P) when any of L/T/P is set. Total uses this course date range inside the timetable."
                        className="flex items-center gap-1.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
                        style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                      >
                        {(c.lectureHours ?? 0) + (c.tutorialHours ?? 0) + (c.practicalHours ?? 0) >
                        0 ? (
                          <span className="rounded bg-[#0d0d0d]/5 px-1.5 py-0.5 text-xs">
                            {courseWeeklyTarget(c)}
                          </span>
                        ) : (
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
                        )}
                        <span>/wk</span>
                      </label>
                      <label
                        title="Total sessions across this course date range inside the timetable. Auto-derived from LTPC as (L+T+P) × course weeks when blank; type a value to override."
                        className="flex items-center gap-1.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60"
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
                        <span>total</span>
                      </label>
                    </div>
                    {/* Placed progress row */}
                    <div className="px-3 pb-2">
                      {(() => {
                        const placed = coursePlacementCounts.get(c.id) ?? 0;
                        const target = courseTotalTarget(
                          c,
                          courseSemesterWeeks(
                            c,
                            activeClass ? stateForClass(activeClass, state) : state,
                          ),
                        );
                        const pct =
                          target > 0 ? Math.min(100, Math.round((placed / target) * 100)) : 0;
                        const done = target > 0 && placed >= target;
                        return (
                          <div
                            className={
                              "flex items-center justify-between gap-2 border-2 px-2 py-1 text-[10px] uppercase tracking-wider " +
                              (done
                                ? "border-emerald-700 bg-emerald-50 text-emerald-800"
                                : "border-[#0d0d0d]/15 bg-[#f5f3ee] text-[#2d2d2d]/70")
                            }
                            style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                            title="Placed sessions in the current date range"
                          >
                            <span className="font-bold">Placed</span>
                            <span className="flex-1">
                              {target > 0 ? (
                                <span className="relative block h-1.5 w-full overflow-hidden bg-[#0d0d0d]/10">
                                  <span
                                    className={
                                      "absolute inset-y-0 left-0 " +
                                      (done ? "bg-emerald-600" : "bg-[#0d0d0d]")
                                    }
                                    style={{ width: pct + "%" }}
                                  />
                                </span>
                              ) : null}
                            </span>
                            <span className="font-bold tabular-nums">
                              {target > 0 ? `${placed} / ${target}` : `${placed}`}
                            </span>
                          </div>
                        );
                      })()}
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
                          {slot.isBreak && <span className="ml-1 text-[#b45309]">·break</span>}
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
                Auto-fill packs each course into the week using its rules. If <b>/wk</b> is 0, Fill
                by Rules uses every allowed weekday and period opportunity.
              </p>
              <div className="mb-3 grid grid-cols-2 gap-1">
                <button
                  onClick={() => runAutoPopulate("Fill Empty", { overwrite: false })}
                  className="border-2 border-[#0d0d0d] bg-white px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
                >
                  Fill Empty
                </button>
                <button
                  onClick={() => {
                    if (confirm("Clear all courses and re-generate?"))
                      runAutoPopulate("Regenerate", { overwrite: true });
                  }}
                  className="border-2 border-[#0d0d0d] bg-[#0d0d0d] px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider text-[#f5f3ee] hover:opacity-90"
                >
                  Regenerate
                </button>
              </div>
              <div className="mb-3 grid grid-cols-2 gap-1">
                <button
                  onClick={() =>
                    runAutoPopulate("Fill by Rules", { overwrite: false, strictRules: true })
                  }
                  className="border-2 border-[#0d0d0d] bg-amber-200 px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-amber-300"
                  title="Places courses in their rule periods. If /wk is 0, uses every allowed period opportunity."
                >
                  Fill by Rules
                </button>
                <button
                  onClick={() => {
                    if (confirm("Clear all courses and fill only by course rules?")) {
                      runAutoPopulate("Regen Rules", { overwrite: true, strictRules: true });
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
                      onChange={async (e) => {
                        const f = e.target.files?.[0];
                        if (f) {
                          const text = await f.text();
                          setPendingBlockCsv({ name: f.name, text });
                          setBlockReport(`Loaded "${f.name}". Click Block to apply.`);
                        }
                        e.target.value = "";
                      }}
                    />
                  </label>
                  <button
                    onClick={() => {
                      if (!pendingBlockCsv) {
                        setBlockReport("Upload a CSV first.");
                        return;
                      }
                      applyBlockCsv(pendingBlockCsv.text);
                    }}
                    disabled={!pendingBlockCsv}
                    className="border border-[#0d0d0d] bg-rose-300 px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-rose-400 disabled:cursor-not-allowed disabled:opacity-40"
                  >
                    Block
                  </button>
                  {pendingBlockCsv && (
                    <button
                      onClick={() => {
                        setPendingBlockCsv(null);
                        setBlockReport("");
                      }}
                      className="border border-[#0d0d0d]/60 bg-white px-2 py-1.5 text-[10px] font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
                    >
                      Clear
                    </button>
                  )}
                </div>
                {pendingBlockCsv && (
                  <p className="mt-2 text-[10px] font-bold text-[#2d2d2d]/80">
                    Loaded: <code>{pendingBlockCsv.name}</code>
                  </p>
                )}
                {blockReport && (
                  <p className="mt-2 whitespace-pre-wrap text-[10px] text-[#2d2d2d]/80">
                    {blockReport}
                  </p>
                )}
                <p className="mt-2 text-[10px] text-[#2d2d2d]/60">
                  Columns: <code>date, Reason, Session</code>. Date is <code>DD/MM/YYYY</code>.
                  Session is <code>all</code> or period numbers like <code>1,3,5,6</code>.
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
                  From · {activeClass?.name ?? ""}
                </div>
                <input
                  type="date"
                  value={activeClass ? classFromDate(activeClass, state) : state.fromDate}
                  onChange={(e) => {
                    const v = e.target.value;
                    if (!activeClass) {
                      setState((s) => ({ ...s, fromDate: v || s.fromDate }));
                      return;
                    }
                    setState((s) => ({
                      ...s,
                      classes: s.classes.map((c) =>
                        c.id === activeClass.id ? { ...c, fromDate: v || undefined } : c,
                      ),
                    }));
                  }}
                  className="border-b border-[#0d0d0d] bg-transparent py-0.5 text-sm font-semibold outline-none"
                  style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                />
              </div>
              <span className="pb-1 text-lg text-[#2d2d2d]/30">/</span>
              <div>
                <div className="text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/50">
                  To · {activeClass?.name ?? ""}
                </div>
                <input
                  type="date"
                  value={activeClass ? classToDate(activeClass, state) : state.toDate}
                  onChange={(e) => {
                    const v = e.target.value;
                    if (!activeClass) {
                      setState((s) => ({ ...s, toDate: v || s.toDate }));
                      return;
                    }
                    setState((s) => ({
                      ...s,
                      classes: s.classes.map((c) =>
                        c.id === activeClass.id ? { ...c, toDate: v || undefined } : c,
                      ),
                    }));
                  }}
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
                onClick={clearCurrentClass}
                className="border-2 border-red-700 bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider text-red-700 transition-transform hover:bg-red-50 active:translate-y-0.5"
                title="Clear course assignments in the current class only"
              >
                Clear Current
              </button>
              <button
                onClick={undo}
                disabled={past.length === 0}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5 disabled:opacity-40 disabled:cursor-not-allowed"
                title="Undo (Cmd+Z / Ctrl+Z)"
              >
                ↩ Undo
              </button>
              <button
                onClick={redo}
                disabled={future.length === 0}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5 disabled:opacity-40 disabled:cursor-not-allowed"
                title="Redo (Cmd+R / Ctrl+R / Ctrl+Y)"
              >
                ↪ Redo
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
                onClick={() => setAdminOpen(true)}
                className="border-2 border-indigo-700 bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider text-indigo-700 transition-transform hover:bg-indigo-50 active:translate-y-0.5"
                title="Admin mode: load other classes' .aadhi files to cross-check all timetables"
              >
                Admin
              </button>
              <button
                onClick={() => setFacultyPanelOpen(true)}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5"
                title="Manage all faculty names across classes"
              >
                👥 Faculties
              </button>
              <button
                onClick={() => setControlPanelOpen(true)}
                className="border-2 border-[#7c3aed] bg-[#ede9fe] px-3 py-2 text-[11px] font-bold uppercase tracking-wider text-[#7c3aed] transition-transform hover:bg-[#ddd6fe] active:translate-y-0.5"
                title="Control panel: manage all classes & courses, toggle common, download timetables"
              >
                📋 Control Panel
              </button>
              {conflictDetailsList.length > 0 && (
                <button
                  onClick={() => setWarningsPanelOpen(true)}
                  className="border-2 border-red-600 bg-red-50 px-3 py-2 text-[11px] font-bold uppercase tracking-wider text-red-700 transition-transform hover:bg-red-100 active:translate-y-0.5"
                  title="Show overlap & rule warnings panel"
                >
                  ⚠️ Warnings ({conflictDetailsList.length})
                </button>
              )}
              <button
                onClick={exportCSV}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5"
              >
                CSV
              </button>
              <button
                onClick={exportASC}
                className="border-2 border-[#0d0d0d] bg-white px-3 py-2 text-[11px] font-bold uppercase tracking-wider transition-transform hover:bg-[#e8e4dd] active:translate-y-0.5"
                title="Export aSc TimeTables 'Data to Fill' workbook (.xlsx)"
              >
                ASC
              </button>
              <button
                onClick={exportExcel}
                className="border-2 border-[#0d0d0d] bg-[#0d0d0d] px-3 py-2 text-[11px] font-bold uppercase tracking-wider text-[#f5f3ee] transition-transform hover:opacity-90 active:translate-y-0.5"
              >
                Export Excel
              </button>
              <button
                onClick={exportFacultyExcel}
                className="border-2 border-[#0369a1] bg-[#e0f2fe] hover:bg-[#bae6fd] px-3 py-2 text-[11px] font-bold uppercase tracking-wider text-[#0369a1] transition-transform active:translate-y-0.5"
                title="Export each individual faculty's timetable into sheets in an Excel file"
              >
                Faculty Timetable
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
              {dates.length} day{dates.length === 1 ? "" : "s"} · {state.slots.length} slots ·{" "}
              {activeVisibleCourseSlots} filled
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
                {activeClass?.name ?? "Class"}: {sessionStats.activePlaced}/
                {sessionStats.activePlanned}
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
                <span className="ml-1 text-[#2d2d2d]/60">· {sessionStats.totalRemaining} left</span>
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
                          (activeClass?.courses ?? []).find((c) => c.id === armedTool.courseId)
                            ?.color ?? "#ddd",
                      }}
                    />
                    <span className="font-bold">
                      {(activeClass?.courses ?? []).find((c) => c.id === armedTool.courseId)
                        ?.name ?? "?"}
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
            {conflictDetailsList.length > 0 && (
              <button
                onClick={() => setWarningsPanelOpen(true)}
                className="w-full text-left flex items-center justify-between border-2 border-red-600 bg-red-50 px-3 py-2 text-xs font-semibold text-red-700 hover:bg-red-100 transition shadow-[2px_2px_0px_0px_#dc2626]"
              >
                <span className="flex items-center gap-2">
                  <span className="text-sm">⚠️</span>
                  <span>
                    Faculty Overlaps / Rule Violations —{" "}
                    <b>
                      {conflictDetailsList.length} warning
                      {conflictDetailsList.length === 1 ? "" : "s"}
                    </b>{" "}
                    detected. Click to view list and locate conflicts.
                  </span>
                </span>
                <span className="underline text-[10px] font-bold uppercase tracking-wider bg-white border border-red-600 px-2 py-0.5 text-red-700 shadow-[1px_1px_0_#dc2626] transition hover:bg-red-50">
                  Open Panel
                </span>
              </button>
            )}
            {dates.length === 0 && (
              <div className="border-2 border-[#d97706] bg-[#fef3c7] px-3 py-2 text-xs font-semibold text-[#b45309]">
                Pick a valid date range.
              </div>
            )}
            <p className="text-[11px] text-[#2d2d2d]/60">
              Click any cell to pick a course. Click-drag to paint. Right-click to change tool.
              Alt+right-click to erase.
            </p>
          </div>

          {/* Grid */}
          <div className="flex-1 overflow-auto p-4 sm:p-8">
            <div className="inline-block min-w-full" ref={gridRef}>
              <table className="w-full border-collapse border-2 border-[#0d0d0d] text-left">
                <thead>
                  <tr className="bg-[#0d0d0d] text-[#f5f3ee]">
                    <th
                      className="sticky left-0 z-10 w-16 border border-[#f5f3ee]/20 bg-[#0d0d0d] p-3 text-[10px] font-bold uppercase tracking-widest text-center"
                      style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                    >
                      Week
                    </th>
                    <th
                      className="sticky left-16 z-10 w-32 border border-[#f5f3ee]/20 bg-[#0d0d0d] p-3 text-[10px] font-bold uppercase tracking-widest"
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
                      ? state.slots.some(
                          (_, slotIdx) => activeClass.grid[`${date}-${slotIdx}`]?.kind === "course",
                        )
                      : false;
                    const weekIndex = activeClass ? weekIndexForDate(activeClass, state, date) : 1;
                    return (
                      <tr key={date} data-filled-row={rowHasCourse ? "true" : undefined}>
                        <td
                          className="sticky left-0 z-10 border-2 border-[#0d0d0d] bg-[#e8e4dd] p-3 text-center align-middle font-bold text-xs w-16"
                          style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                        >
                          W{weekIndex}
                        </td>
                        <th
                          className="sticky left-16 z-10 border-2 border-[#0d0d0d] bg-[#e8e4dd] p-3 text-left align-middle w-32"
                          style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                        >
                          <div className="text-xs font-bold uppercase tracking-wider">
                            {weekday}
                          </div>
                          <div className="text-[10px] font-medium text-[#2d2d2d]/60">{dstr}</div>
                        </th>
                        {state.slots.map((sl, i) => {
                          if (!activeClass) return null;
                          const key = `${date}-${i}`;
                          const cell = activeClass.grid[key];
                          const isConflict = conflicts.has(`${activeClass.id}:${key}`);
                          const isConflictIgnored = state.ignoredConflicts?.includes(
                            `${activeClass.id}:${key}`,
                          );

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
                              id={`cell-${activeClass.id}-${date}-${i}`}
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
                                    {isConflictIgnored && (
                                      <span
                                        className="ml-1 text-[9px] font-normal text-gray-500 opacity-60"
                                        title="Conflict Ignored"
                                      >
                                        ⚠️🚫
                                      </span>
                                    )}
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

          {/* Bottom tools bar: course chips, eraser, override toggle */}
          <div className="sticky bottom-0 z-30 border-t-2 border-[#0d0d0d] bg-[#f5f3ee]/95 px-4 py-2 backdrop-blur sm:px-8">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/60">
                Tools
              </span>
              <div className="flex flex-wrap items-center gap-1">
                {(activeClass?.courses ?? []).map((c) => {
                  const active = armedTool?.kind === "course" && armedTool.courseId === c.id;
                  return (
                    <button
                      key={c.id}
                      onClick={() =>
                        setArmedTool(active ? null : { kind: "course", courseId: c.id })
                      }
                      className={
                        "flex items-center gap-1 border-2 px-2 py-1 text-[11px] font-bold transition " +
                        (active
                          ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee] shadow-[2px_2px_0px_0px_#0d0d0d]"
                          : "border-[#0d0d0d]/60 bg-white text-[#0d0d0d] hover:border-[#0d0d0d]")
                      }
                      title={`${c.name}${c.faculty ? " · " + c.faculty : ""}`}
                    >
                      <span
                        className="inline-block h-3 w-3 border border-[#0d0d0d]/40"
                        style={{ backgroundColor: c.color }}
                      />
                      <span className="max-w-[10rem] truncate">{c.name || "Untitled"}</span>
                    </button>
                  );
                })}
                {(activeClass?.courses ?? []).length === 0 && (
                  <span className="text-[11px] italic text-[#2d2d2d]/50">
                    Add a course in the sidebar
                  </span>
                )}
              </div>
              <span className="mx-1 h-6 w-px bg-[#0d0d0d]/20" />
              <button
                onClick={() => setArmedTool(armedTool?.kind === "erase" ? null : { kind: "erase" })}
                className={
                  "border-2 px-2 py-1 text-[11px] font-bold uppercase tracking-wider transition " +
                  (armedTool?.kind === "erase"
                    ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee] shadow-[2px_2px_0px_0px_#0d0d0d]"
                    : "border-[#0d0d0d]/60 bg-white hover:border-[#0d0d0d]")
                }
                title="Eraser — click cells to clear"
              >
                Eraser
              </button>
              <button
                onClick={() => setArmedTool(armedTool?.kind === "break" ? null : { kind: "break" })}
                className={
                  "border-2 px-2 py-1 text-[11px] font-bold uppercase tracking-wider transition " +
                  (armedTool?.kind === "break"
                    ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee] shadow-[2px_2px_0px_0px_#0d0d0d]"
                    : "border-[#0d0d0d]/60 bg-white hover:border-[#0d0d0d]")
                }
              >
                Break
              </button>
              <button
                onClick={() =>
                  setArmedTool(armedTool?.kind === "blocked" ? null : { kind: "blocked" })
                }
                className={
                  "border-2 px-2 py-1 text-[11px] font-bold uppercase tracking-wider transition " +
                  (armedTool?.kind === "blocked"
                    ? "border-[#0d0d0d] bg-[#0d0d0d] text-[#f5f3ee] shadow-[2px_2px_0px_0px_#0d0d0d]"
                    : "border-[#0d0d0d]/60 bg-white hover:border-[#0d0d0d]")
                }
              >
                Block
              </button>
              <span className="mx-1 h-6 w-px bg-[#0d0d0d]/20" />
              <button
                onClick={() => setOverrideMode((v) => !v)}
                className={
                  "flex items-center gap-2 border-2 px-2 py-1 text-[11px] font-bold uppercase tracking-wider transition " +
                  (overrideMode
                    ? "border-red-700 bg-red-600 text-white shadow-[2px_2px_0px_0px_#0d0d0d]"
                    : "border-[#0d0d0d]/60 bg-white hover:border-[#0d0d0d]")
                }
                title="Bypass faculty-busy and course-rule checks when placing manually"
              >
                <span
                  className={
                    "inline-block h-3 w-6 border-2 " +
                    (overrideMode ? "border-white bg-white/30" : "border-[#0d0d0d]/50 bg-[#f5f3ee]")
                  }
                >
                  <span
                    className={
                      "block h-full w-1/2 " +
                      (overrideMode ? "translate-x-full bg-white" : "bg-[#0d0d0d]/50")
                    }
                  />
                </span>
                Override {overrideMode ? "ON" : "OFF"}
              </button>
              {overrideMode && (
                <span className="text-[10px] font-semibold text-red-700">
                  Rules & faculty conflicts bypassed for manual placement
                </span>
              )}
              <span className="mx-1 h-6 w-px bg-[#0d0d0d]/20" />
              <button
                onClick={() => {
                  if (!state.frozen) {
                    if (
                      !confirm(
                        "Freeze the timetable? This locks all cells, hides conflict warnings, and makes overrides permanent. You can unfreeze later.",
                      )
                    )
                      return;
                  }
                  setState((s) => {
                    const nextFrozen = !s.frozen;
                    if (!nextFrozen) return { ...s, frozen: false };
                    // Lock every existing course placement so future auto-fills
                    // (and any subsequent unfreeze + edit) cannot silently
                    // remove or overwrite the manually reviewed timetable.
                    const classes = s.classes.map((cls) => {
                      const grid: Record<string, Cell> = {};
                      Object.entries(cls.grid).forEach(([k, cell]) => {
                        grid[k] = cell.kind === "course" ? { ...cell, locked: true } : cell;
                      });
                      return { ...cls, grid };
                    });
                    return { ...s, frozen: true, classes };
                  });
                }}
                className={
                  "flex items-center gap-2 border-2 px-2 py-1 text-[11px] font-bold uppercase tracking-wider transition " +
                  (state.frozen
                    ? "border-sky-700 bg-sky-600 text-white shadow-[2px_2px_0px_0px_#0d0d0d]"
                    : "border-[#0d0d0d]/60 bg-white hover:border-[#0d0d0d]")
                }
                title="Lock the timetable, hide conflict warnings, and make overrides permanent"
              >
                <span aria-hidden>{state.frozen ? "🔒" : "❄"}</span>
                {state.frozen ? "Frozen — Unfreeze" : "Freeze"}
              </button>
              {state.frozen && (
                <span className="text-[10px] font-semibold text-sky-700">
                  Timetable locked · overrides permanent
                </span>
              )}
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
                  // Count placed sessions of this course in the class date range
                  const placedForCourse = coursePlacementCounts.get(c.id) ?? 0;
                  const dateRange =
                    c.fromDate || c.toDate
                      ? `${c.fromDate ?? "start"} → ${c.toDate ?? "end"}`
                      : null;
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
                            ? c.allowedSlots.map((i) => `P${periodNumberFor(i)}`).join(",")
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
                          {(() => {
                            const t = courseTotalTarget(
                              c,
                              courseSemesterWeeks(
                                c,
                                activeClass ? stateForClass(activeClass, state) : state,
                              ),
                            );
                            return t > 0
                              ? `${placedForCourse} / ${t} sessions`
                              : `${placedForCourse} placed · no total set`;
                          })()}
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

              {/* Ignore conflict option */}
              <div className="mt-4 border-t border-[#0d0d0d]/10 pt-3">
                <label className="flex items-center gap-1.5 text-[10px] font-bold uppercase tracking-wider text-[#2d2d2d]/60 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={
                      state.ignoredConflicts?.includes(
                        `${activeClass.id}:${picker.date}-${picker.slotIdx}`,
                      ) ?? false
                    }
                    onChange={(e) => {
                      const keyStr = `${activeClass.id}:${picker.date}-${picker.slotIdx}`;
                      setState((s) => {
                        const list = s.ignoredConflicts ? [...s.ignoredConflicts] : [];
                        if (e.target.checked) {
                          if (!list.includes(keyStr)) list.push(keyStr);
                        } else {
                          const idx = list.indexOf(keyStr);
                          if (idx !== -1) list.splice(idx, 1);
                        }
                        return { ...s, ignoredConflicts: list };
                      });
                    }}
                    className="accent-[#0d0d0d]"
                  />
                  <span>Ignore conflict on this slot</span>
                </label>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Rules editor modal */}
      {rulesFor &&
        activeClass &&
        (() => {
          const course = activeClass.courses.find((c) => c.id === rulesFor);
          if (!course) return null;
          const wdRule = course.allowedWeekdays ?? [];
          const wdAll = wdRule.length === 0;
          const slotRule = course.allowedSlots ?? [];
          const slotAll = slotRule.length === 0;
          const nonBreakIdxs = state.slots.map((sl, i) => ({ sl, i })).filter((x) => !x.sl.isBreak);
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
                        onClick={() =>
                          updateCourse(course.id, { fromDate: undefined, toDate: undefined })
                        }
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
                          onChange={(e) =>
                            updateCourse(course.id, { fromDate: e.target.value || undefined })
                          }
                          className="border border-[#0d0d0d]/20 bg-white px-2 py-1 text-xs outline-none focus:border-[#0d0d0d]"
                        />
                      </label>
                      <label className="flex flex-col gap-0.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60">
                        <span>To</span>
                        <input
                          type="date"
                          value={course.toDate ?? ""}
                          onChange={(e) =>
                            updateCourse(course.id, { toDate: e.target.value || undefined })
                          }
                          className="border border-[#0d0d0d]/20 bg-white px-2 py-1 text-xs outline-none focus:border-[#0d0d0d]"
                        />
                      </label>
                    </div>
                    <p className="mt-1 text-[10px] text-[#2d2d2d]/50">
                      Restrict this course to a specific date window. Leave blank to use the full
                      timetable range.
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
                      Pick the weekdays this faculty is available. Deselect all to treat every day
                      as allowed.
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
                        const periodNum = state.slots
                          .slice(0, i + 1)
                          .filter((x) => !x.isBreak).length;
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
                                allowedSlots: next.length === allIdxs.length ? [] : next,
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
                      Pick the periods this course can be scheduled in. Break slots are excluded.
                    </p>
                  </div>

                  <div>
                    <div className="mb-2 flex items-center justify-between">
                      <span className="text-[11px] font-bold uppercase tracking-widest text-[#2d2d2d]/70">
                        Available periods (per-day)
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
                      Optional. Override the default periods above for specific weekdays (e.g. Mon
                      P1–P2, Thu P5–P6). Unset weekdays fall back to the default.
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
                          <div key={wd} className="border border-[#0d0d0d]/15 bg-white px-2 py-1.5">
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
                                    style={{
                                      fontFamily: "'JetBrains Mono', ui-monospace, monospace",
                                    }}
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

                    <div>
                      <div className="mb-2 flex items-center justify-between mt-4">
                        <span className="text-[11px] font-bold uppercase tracking-widest text-[#2d2d2d]/70">
                          Classes at a stretch
                        </span>
                        <label className="flex items-center gap-1.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60 font-semibold cursor-pointer">
                          <input
                            type="checkbox"
                            checked={course.stretchRuleEnabled !== false}
                            onChange={(e) =>
                              updateCourse(course.id, {
                                stretchRuleEnabled: e.target.checked,
                              })
                            }
                            className="accent-[#0d0d0d]"
                          />
                          <span>Enabled</span>
                        </label>
                      </div>

                      {course.stretchRuleEnabled !== false ? (
                        <>
                          <div className="grid grid-cols-2 gap-2">
                            <label className="flex flex-col gap-0.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60">
                              <span>Min consecutive</span>
                              <select
                                value={course.minDurationSlots ?? course.durationSlots}
                                onChange={(e) => {
                                  const min = parseInt(e.target.value, 10);
                                  const max = course.maxDurationSlots ?? course.durationSlots;
                                  const nextMax = max < min ? min : max;
                                  const currentDur = course.durationSlots;
                                  const nextDur =
                                    currentDur < min
                                      ? min
                                      : currentDur > nextMax
                                        ? nextMax
                                        : currentDur;
                                  updateCourse(course.id, {
                                    minDurationSlots: min,
                                    maxDurationSlots: nextMax,
                                    durationSlots: nextDur,
                                  });
                                }}
                                className="border border-[#0d0d0d]/20 bg-white px-2 py-1 text-xs outline-none focus:border-[#0d0d0d]"
                                style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                              >
                                {Array.from(
                                  { length: Math.min(7, Math.max(1, nonBreakCount(state.slots))) },
                                  (_, idx) => idx + 1,
                                ).map((num) => (
                                  <option key={num} value={num}>
                                    {num} period{num > 1 ? "s" : ""}
                                  </option>
                                ))}
                              </select>
                            </label>
                            <label className="flex flex-col gap-0.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60">
                              <span>Max consecutive</span>
                              <select
                                value={course.maxDurationSlots ?? course.durationSlots}
                                onChange={(e) => {
                                  const max = parseInt(e.target.value, 10);
                                  const min = course.minDurationSlots ?? course.durationSlots;
                                  const nextMin = min > max ? max : min;
                                  const currentDur = course.durationSlots;
                                  const nextDur =
                                    currentDur < nextMin
                                      ? nextMin
                                      : currentDur > max
                                        ? max
                                        : currentDur;
                                  updateCourse(course.id, {
                                    maxDurationSlots: max,
                                    minDurationSlots: nextMin,
                                    durationSlots: nextDur,
                                  });
                                }}
                                className="border border-[#0d0d0d]/20 bg-white px-2 py-1 text-xs outline-none focus:border-[#0d0d0d]"
                                style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                              >
                                {Array.from(
                                  { length: Math.min(7, Math.max(1, nonBreakCount(state.slots))) },
                                  (_, idx) => idx + 1,
                                ).map((num) => (
                                  <option key={num} value={num}>
                                    {num} period{num > 1 ? "s" : ""}
                                  </option>
                                ))}
                              </select>
                            </label>
                          </div>
                          <p className="mt-1 text-[10px] text-[#2d2d2d]/50">
                            Specify the minimum and maximum consecutive classes of this course that
                            can be scheduled in a single stretch (1 to 7).
                          </p>
                        </>
                      ) : (
                        <p className="text-[10px] text-[#2d2d2d]/50 italic">
                          Stretch limits are disabled. This course can be scheduled for any
                          consecutive number of periods.
                        </p>
                      )}
                    </div>

                    {/* Common Course Toggle */}
                    <div className="mt-4 border-t border-dashed border-[#0d0d0d]/15 pt-4">
                      <div className="flex items-center justify-between">
                        <span className="text-[11px] font-bold uppercase tracking-widest text-[#2d2d2d]/70">
                          Combined / Common Subject
                        </span>
                        <label className="flex items-center gap-1.5 text-[10px] uppercase tracking-wider text-[#2d2d2d]/60 font-semibold cursor-pointer">
                          <input
                            type="checkbox"
                            checked={course.common === true}
                            onChange={(e) =>
                              updateCourse(course.id, {
                                common: e.target.checked || undefined,
                              })
                            }
                            className="accent-[#0d0d0d]"
                          />
                          <span>Common</span>
                        </label>
                      </div>
                      <p className="mt-1 text-[10px] text-[#2d2d2d]/50">
                        When enabled, classes taking this same course (matching name) at the same
                        time will not trigger a faculty schedule conflict (treated as a combined
                        lecture).
                      </p>
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
      {autoStatus && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-[#0d0d0d]/60 backdrop-blur-sm">
          <div className="w-[min(420px,90vw)] border-2 border-[#0d0d0d] bg-[#f5f3ee] p-5 shadow-[6px_6px_0_#0d0d0d]">
            <div className="mb-1 text-[10px] font-bold uppercase tracking-[0.2em] text-[#2d2d2d]/70">
              Auto-fill
            </div>
            <div className="mb-3 text-lg font-bold text-[#0d0d0d]">{autoStatus.label}</div>
            <div className="mb-2 h-2 w-full overflow-hidden border-2 border-[#0d0d0d] bg-white">
              <div className="h-full w-1/3 animate-[autofill_1.1s_ease-in-out_infinite] bg-amber-400" />
            </div>
            <div className="text-[11px] text-[#2d2d2d]/80">{autoStatus.phase}</div>
          </div>
          <style>{`
            @keyframes autofill {
              0% { transform: translateX(-100%); }
              100% { transform: translateX(400%); }
            }
            @keyframes cellHighlight {
              0% {
                outline: 4px solid #dc2626;
                outline-offset: 2px;
                box-shadow: 0 0 0 8px rgba(220, 38, 38, 0.4);
                background-color: #fef2f2;
              }
              50% {
                outline: 4px solid #f87171;
                outline-offset: 2px;
                box-shadow: 0 0 0 12px rgba(248, 113, 113, 0.2);
                background-color: #fee2e2;
              }
              100% {
                outline: 4px solid transparent;
                outline-offset: 0px;
                box-shadow: 0 0 0 0px transparent;
              }
            }
            .cell-highlight-flash {
              animation: cellHighlight 3s ease-out;
            }
          `}</style>
        </div>
      )}
      {adminOpen && (
        <div
          className="fixed inset-0 z-[90] bg-[#0d0d0d]/50 p-4 backdrop-blur-sm"
          onClick={() => setAdminOpen(false)}
        >
          <div
            className="absolute left-1/2 top-1/2 flex max-h-[85vh] w-[min(760px,calc(100%-2rem))] -translate-x-1/2 -translate-y-1/2 flex-col border-2 border-[#0d0d0d] bg-[#f5f3ee] shadow-[8px_8px_0px_0px_#0d0d0d]"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between border-b-2 border-[#0d0d0d] bg-indigo-700 px-4 py-3 text-white">
              <div>
                <div className="text-[10px] font-bold uppercase tracking-widest text-white/70">
                  Admin Mode
                </div>
                <div
                  className="text-sm font-bold"
                  style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                >
                  Cross-check timetables across classes
                </div>
              </div>
              <button
                onClick={() => setAdminOpen(false)}
                className="border-2 border-white bg-transparent px-3 py-1 text-[11px] font-bold uppercase tracking-wider hover:bg-white hover:text-indigo-700"
              >
                Close
              </button>
            </div>
            <div className="flex-1 overflow-y-auto p-4">
              <input
                ref={adminInputRef}
                type="file"
                accept=".aadhi,application/json"
                multiple
                className="hidden"
                onChange={async (e) => {
                  const files = Array.from(e.target.files ?? []);
                  for (const f of files) await loadAdminReference(f);
                  e.target.value = "";
                }}
              />
              <div className="mb-3 flex flex-wrap items-center gap-2">
                <button
                  onClick={() => adminInputRef.current?.click()}
                  className="border-2 border-[#0d0d0d] bg-white px-3 py-1.5 text-[11px] font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
                >
                  Load reference .aadhi
                </button>
                {adminRefs.length > 0 && (
                  <button
                    onClick={() => setAdminRefs([])}
                    className="border-2 border-red-700 bg-white px-3 py-1.5 text-[11px] font-bold uppercase tracking-wider text-red-700 hover:bg-red-50"
                  >
                    Clear all
                  </button>
                )}
                <span className="text-[11px] text-[#2d2d2d]/70">
                  Loaded references are held in memory only — nothing is written back.
                </span>
              </div>
              <div className="mb-4">
                <div className="mb-1 text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/70">
                  Loaded references ({adminRefs.length})
                </div>
                {adminRefs.length === 0 ? (
                  <div className="border-2 border-dashed border-[#0d0d0d]/30 bg-white p-3 text-xs text-[#2d2d2d]/60">
                    No reference files loaded. Add other classes' saved .aadhi files to compare.
                  </div>
                ) : (
                  <ul className="space-y-1">
                    {adminRefs.map((r) => (
                      <li
                        key={r.id}
                        className="flex items-center justify-between border border-[#0d0d0d]/30 bg-white px-2 py-1 text-xs"
                      >
                        <span className="truncate">
                          <span className="font-semibold">{r.name}</span>
                          <span className="ml-2 text-[#2d2d2d]/60">
                            {r.state.classes.length} class{r.state.classes.length === 1 ? "" : "es"}{" "}
                            · {r.state.fromDate} → {r.state.toDate}
                          </span>
                        </span>
                        <button
                          onClick={() => setAdminRefs((prev) => prev.filter((x) => x.id !== r.id))}
                          className="ml-2 border border-red-700 px-1.5 py-0.5 text-[10px] font-bold uppercase text-red-700 hover:bg-red-50"
                        >
                          Remove
                        </button>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
              <div>
                <div className="mb-1 flex items-center justify-between">
                  <div className="text-[10px] font-bold uppercase tracking-widest text-[#2d2d2d]/70">
                    Faculty conflicts across all timetables
                  </div>
                  <span
                    className={
                      "border-2 px-2 py-0.5 text-[10px] font-bold uppercase " +
                      (adminReport.length === 0
                        ? "border-green-700 bg-green-50 text-green-700"
                        : "border-red-700 bg-red-50 text-red-700")
                    }
                  >
                    {adminReport.length === 0
                      ? "No conflicts"
                      : `${adminReport.length} conflict${adminReport.length === 1 ? "" : "s"}`}
                  </span>
                </div>
                {adminReport.length > 0 && (
                  <div className="max-h-[40vh] overflow-y-auto border-2 border-[#0d0d0d] bg-white">
                    <table className="w-full border-collapse text-[11px]">
                      <thead className="sticky top-0 bg-[#0d0d0d] text-white">
                        <tr>
                          <th className="border border-[#0d0d0d]/40 px-2 py-1 text-left">Date</th>
                          <th className="border border-[#0d0d0d]/40 px-2 py-1 text-left">Period</th>
                          <th className="border border-[#0d0d0d]/40 px-2 py-1 text-left">
                            Faculty conflict
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {adminReport.map((row) => {
                          const facMap = new Map<string, typeof row.conflicts>();
                          row.conflicts.forEach((c) => {
                            const list = facMap.get(c.faculty) ?? [];
                            list.push(c);
                            facMap.set(c.faculty, list);
                          });
                          return (
                            <tr key={row.key} className="odd:bg-[#f5f3ee]">
                              <td
                                className="border border-[#0d0d0d]/20 px-2 py-1 align-top"
                                style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                              >
                                {row.date}
                              </td>
                              <td
                                className="border border-[#0d0d0d]/20 px-2 py-1 align-top"
                                style={{ fontFamily: "'JetBrains Mono', ui-monospace, monospace" }}
                              >
                                {periodLabelFor(row.slotIdx)}
                              </td>
                              <td className="border border-[#0d0d0d]/20 px-2 py-1 align-top">
                                {Array.from(facMap.entries()).map(([faculty, entries]) => (
                                  <div key={faculty} className="mb-1 last:mb-0">
                                    <span className="font-bold">{faculty}</span>
                                    <span className="text-[#2d2d2d]/70"> — busy in </span>
                                    {entries.map((e, i) => (
                                      <span
                                        key={i}
                                        className="mr-1 inline-block border border-[#0d0d0d]/40 bg-white px-1"
                                      >
                                        {e.origin} / {e.className} · {e.courseName}
                                      </span>
                                    ))}
                                  </div>
                                ))}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
                {adminReport.length === 0 && adminRefs.length > 0 && (
                  <div className="border-2 border-green-700 bg-green-50 p-3 text-xs text-green-800">
                    All references cross-checked — no shared faculty is double-booked in the same
                    date + period.
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
      {/* Warnings Panel Backdrop Overlay */}
      {warningsPanelOpen && (
        <div
          className="fixed inset-0 z-[90] bg-[#0d0d0d]/35 backdrop-blur-[2px]"
          onClick={() => setWarningsPanelOpen(false)}
        />
      )}
      {/* Warnings Panel Drawer */}
      <div
        className={`fixed right-0 top-0 h-full w-[min(450px,100vw)] bg-[#f5f3ee] border-l-2 border-[#0d0d0d] shadow-[0_0_50px_rgba(0,0,0,0.3)] z-[95] flex flex-col transition-transform duration-300 ease-in-out ${
          warningsPanelOpen ? "translate-x-0" : "translate-x-full"
        }`}
      >
        <div className="flex items-center justify-between border-b-2 border-[#0d0d0d] bg-red-700 px-4 py-3 text-white">
          <div>
            <div className="text-[10px] font-bold uppercase tracking-widest text-white/70">
              Timetable Checker
            </div>
            <div
              className="text-sm font-bold"
              style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
            >
              Overlap & Rule Warnings ({conflictDetailsList.length})
            </div>
          </div>
          <button
            onClick={() => setWarningsPanelOpen(false)}
            className="border-2 border-white bg-transparent px-3 py-1 text-[11px] font-bold uppercase tracking-wider hover:bg-white hover:text-red-700 transition"
          >
            Close
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {conflictDetailsList.length === 0 ? (
            <div className="border-2 border-dashed border-green-700 bg-green-50 p-6 text-center text-green-800">
              <span className="text-2xl block mb-2">🎉</span>
              <div className="font-bold text-sm">No Conflicts Found</div>
              <div className="text-xs text-green-700/80 mt-1">
                All classes look good! No faculty double-bookings or rule violations.
              </div>
            </div>
          ) : (
            conflictDetailsList.map((warn) => (
              <button
                key={warn.id}
                onClick={() => handleWarningClick(warn)}
                className="w-full text-left border-2 border-[#0d0d0d] bg-white p-3 hover:bg-[#e8e4dd] transition-all hover:-translate-y-0.5 hover:shadow-[3px_3px_0px_0px_#0d0d0d] flex flex-col gap-2 group active:translate-y-0 active:shadow-[1px_1px_0px_0px_#0d0d0d]"
              >
                <div className="flex items-center justify-between w-full">
                  <span
                    className={`px-2 py-0.5 text-[9px] font-bold uppercase border ${
                      warn.type === "faculty"
                        ? "border-red-600 bg-red-50 text-red-700"
                        : "border-amber-600 bg-amber-50 text-amber-700"
                    }`}
                  >
                    {warn.type === "faculty" ? "Faculty Overlap" : "Rule Violation"}
                  </span>
                  <span className="text-[10px] font-bold uppercase tracking-wider bg-[#0d0d0d]/5 px-1.5 py-0.5 rounded">
                    {warn.className}
                  </span>
                </div>

                <div className="text-xs font-semibold text-[#0d0d0d]">{warn.description}</div>

                <div className="flex items-center justify-between text-[10px] text-[#2d2d2d]/60 font-mono mt-1 border-t border-dashed border-[#0d0d0d]/10 pt-2 w-full">
                  <span>
                    {warn.weekday}, {warn.date} · {warn.periodLabel}
                  </span>
                  <span className="text-[9px] font-bold uppercase text-indigo-700 group-hover:underline flex items-center gap-1">
                    Locate Cell &rarr;
                  </span>
                </div>
              </button>
            ))
          )}
        </div>
      </div>

      {/* Control Panel Backdrop Overlay */}
      {controlPanelOpen && (
        <div
          className="fixed inset-0 z-[90] bg-[#0d0d0d]/35 backdrop-blur-[2px]"
          onClick={() => setControlPanelOpen(false)}
        />
      )}
      {/* Control Panel Drawer */}
      <div
        className={`fixed right-0 top-0 h-full w-[min(600px,100vw)] bg-[#faf8ff] border-l-2 border-[#7c3aed] shadow-[0_0_50px_rgba(124,58,237,0.2)] z-[95] flex flex-col transition-transform duration-300 ease-in-out ${
          controlPanelOpen ? "translate-x-0" : "translate-x-full"
        }`}
      >
        {/* Header */}
        <div className="flex items-center justify-between border-b-2 border-[#7c3aed] bg-[#7c3aed] px-4 py-3 text-[#f5f3ee]">
          <div>
            <div className="text-[10px] font-bold uppercase tracking-widest text-[#ede9fe]/60">
              Download Manager
            </div>
            <div
              className="text-sm font-bold"
              style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
            >
              📋 Control Panel
            </div>
          </div>
          <button
            onClick={() => setControlPanelOpen(false)}
            className="border-2 border-[#ede9fe] bg-transparent px-3 py-1 text-[11px] font-bold uppercase tracking-wider text-[#ede9fe] hover:bg-[#ede9fe] hover:text-[#7c3aed] transition"
          >
            Close
          </button>
        </div>

        {/* Legend + global download */}
        <div className="flex items-center flex-wrap gap-3 px-4 py-2 border-b border-[#7c3aed]/10 bg-[#ede9fe]/30">
          <span className="flex items-center gap-1 text-[9px] font-bold uppercase tracking-wider text-[#7c3aed]/60">
            <span>📊</span> Class
          </span>
          <span className="flex items-center gap-1 text-[9px] font-bold uppercase tracking-wider text-[#7c3aed]/60">
            <span>📥</span> Course
          </span>
          <span className="flex items-center gap-1 text-[9px] font-bold uppercase tracking-wider text-[#7c3aed]/60">
            <span>👤</span> Faculty
          </span>
          <span className="text-[9px] italic font-normal text-[#7c3aed]/40">
            Each: Summary + Timetable sheets
          </span>
          <button
            onClick={() => exportSummaryReport()}
            className="ml-auto shrink-0 flex items-center gap-1.5 border-2 border-[#7c3aed] bg-white px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-[#7c3aed] hover:bg-[#ede9fe] transition active:translate-y-0.5 shadow-[2px_2px_0_0_#7c3aed]"
            title="Download a classwise summary report with LTPC, faculty, placed vs. total sessions"
          >
            📋 Download Summary
          </button>
          <button
            onClick={() => exportAllFacultiesZip()}
            className="shrink-0 flex items-center gap-1.5 border-2 border-[#0369a1] bg-[#e0f2fe] px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-[#0369a1] hover:bg-[#bae6fd] transition active:translate-y-0.5 shadow-[2px_2px_0_0_#0369a1]"
            title="Generate individual Excel files for all faculties and download them as a ZIP archive"
          >
            📦 Download All Faculties (ZIP)
          </button>
          <button
            onClick={() => exportSchoolWorkload()}
            className="shrink-0 flex items-center gap-1.5 border-2 border-[#15803d] bg-[#dcfce7] px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-[#15803d] hover:bg-[#bbf7d0] transition active:translate-y-0.5 shadow-[2px_2px_0_0_#15803d]"
            title="Download School Workload matrix showing course details, dates, and faculty free hours on occupied dates"
          >
            🏫 School Workload
          </button>
          <button
            onClick={() => exportFacultyWorkloadSummary()}
            className="shrink-0 flex items-center gap-1.5 border-2 border-[#b45309] bg-[#fef3c7] px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-[#b45309] hover:bg-[#fde68a] transition active:translate-y-0.5 shadow-[2px_2px_0_0_#b45309]"
            title="Download School Faculty Workload Summary matrix showing free hour numbers for each faculty member across all dates"
          >
            👨‍🏫 Faculty Workload Summary
          </button>
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {state.classes.length === 0 ? (
            <div className="text-xs text-[#2d2d2d]/50 italic p-4 border-2 border-dashed border-[#7c3aed]/20 bg-white text-center">
              No classes yet. Add a class to get started.
            </div>
          ) : (
            <div className="space-y-3">
              {state.classes.map((cls) => {
                const isExpanded = expandedControlClass === cls.id;
                return (
                  <div
                    key={cls.id}
                    className="border border-[#7c3aed]/20 bg-white shadow-[2px_2px_0px_0px_rgba(124,58,237,0.12)]"
                  >
                    {/* Class row */}
                    <div className="flex flex-col border-b border-[#7c3aed]/10">
                      <div className="flex items-center gap-2 px-3 py-2 bg-[#ede9fe]/40">
                        <button
                          onClick={() => setExpandedControlClass(isExpanded ? null : cls.id)}
                          className="min-w-0 flex-1 flex items-center gap-2 text-left"
                        >
                          <span className="text-[10px] text-[#7c3aed]/50 font-mono shrink-0">
                            {isExpanded ? "▼" : "▶"}
                          </span>
                          <span
                            className="text-xs font-bold text-[#2d2d2d] truncate"
                            style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
                          >
                            {cls.name}
                          </span>
                          <span className="text-[9px] text-[#2d2d2d]/40 font-normal shrink-0">
                            ({cls.courses.length} course{cls.courses.length === 1 ? "" : "s"})
                          </span>
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            exportClassTimetable(cls);
                          }}
                          className="shrink-0 flex items-center gap-1 border border-[#7c3aed]/40 bg-[#ede9fe] px-2 py-1 text-[9px] font-bold uppercase tracking-wider text-[#7c3aed] hover:bg-[#ddd6fe] transition active:translate-y-0.5"
                          title="Download class timetable (Summary + Timetable)"
                        >
                          📊 Class
                        </button>
                      </div>
                      {/* Department row */}
                      <div className="flex items-center gap-2 px-3 pb-2 bg-[#ede9fe]/20">
                        <span className="text-[9px] font-bold uppercase tracking-wider text-[#7c3aed]/50 shrink-0">
                          Dept:
                        </span>
                        <input
                          type="text"
                          value={cls.department ?? ""}
                          onChange={(e) => {
                            const val = e.target.value;
                            setState((s) => ({
                              ...s,
                              classes: s.classes.map((cl) =>
                                cl.id === cls.id ? { ...cl, department: val } : cl,
                              ),
                            }));
                          }}
                          placeholder="e.g. Computer Science"
                          className="min-w-0 flex-1 bg-white border border-[#7c3aed]/20 px-2 py-0.5 text-[10px] text-[#2d2d2d] placeholder:text-[#2d2d2d]/30 outline-none focus:border-[#7c3aed]/60 focus:ring-0 rounded-none"
                        />
                        {cls.department?.trim() && (
                          <span className="shrink-0 text-[9px] font-bold px-1.5 py-0.5 bg-[#7c3aed] text-white rounded">
                            {cls.department.trim()}
                          </span>
                        )}
                      </div>
                    </div>

                    {/* Courses list */}
                    {isExpanded && (
                      <div className="divide-y divide-[#7c3aed]/5">
                        {cls.courses.length === 0 ? (
                          <div className="px-4 py-3 text-[11px] text-[#2d2d2d]/40 italic">
                            No courses in this class.
                          </div>
                        ) : (
                          cls.courses.map((course) => {
                            const courseFaculties = course.faculty
                              ? course.faculty
                                  .split(",")
                                  .map((f) => f.trim())
                                  .filter(Boolean)
                              : [];
                            const hasLTPC =
                              (course.lectureHours ?? 0) +
                                (course.tutorialHours ?? 0) +
                                (course.practicalHours ?? 0) >
                              0;
                            return (
                              <div
                                key={course.id}
                                className="flex items-start gap-2 px-3 py-2.5 hover:bg-[#faf8ff] transition"
                              >
                                {/* Color swatch */}
                                <div
                                  className="w-3 h-3 rounded-full shrink-0 mt-0.5 border border-[#0d0d0d]/10"
                                  style={{ backgroundColor: course.color }}
                                />
                                {/* Info */}
                                <div className="min-w-0 flex-1">
                                  <div className="text-[11px] font-bold text-[#2d2d2d] truncate">
                                    {course.name}
                                    {course.common && (
                                      <span className="ml-1.5 text-[9px] font-bold text-[#0369a1] bg-[#e0f2fe] px-1 py-0.5 rounded">
                                        Common
                                      </span>
                                    )}
                                  </div>
                                  <div className="text-[9px] text-[#2d2d2d]/50 truncate mt-0.5">
                                    {course.faculty || <span className="italic">No faculty</span>}
                                    {hasLTPC && (
                                      <span className="ml-1.5 font-mono text-[#7c3aed]/70">
                                        L{course.lectureHours ?? 0}T{course.tutorialHours ?? 0}P
                                        {course.practicalHours ?? 0}C{course.credits ?? 0}
                                      </span>
                                    )}
                                  </div>
                                </div>
                                {/* Common toggle */}
                                <label
                                  className="shrink-0 flex items-center gap-1 cursor-pointer mt-0.5"
                                  title="Toggle as common / combined course across classes"
                                  onClick={(e) => e.stopPropagation()}
                                >
                                  <input
                                    type="checkbox"
                                    checked={!!course.common}
                                    onChange={(e) => {
                                      e.stopPropagation();
                                      setState((s) => ({
                                        ...s,
                                        classes: s.classes.map((cl) =>
                                          cl.id === cls.id
                                            ? {
                                                ...cl,
                                                courses: cl.courses.map((co) =>
                                                  co.id === course.id
                                                    ? { ...co, common: e.target.checked }
                                                    : co,
                                                ),
                                              }
                                            : cl,
                                        ),
                                      }));
                                    }}
                                    className="w-3 h-3 accent-[#7c3aed]"
                                  />
                                  <span className="text-[9px] font-bold uppercase tracking-wider text-[#7c3aed]/60">
                                    Common
                                  </span>
                                </label>
                                {/* Download buttons */}
                                <div className="shrink-0 flex items-center gap-1 mt-0.5">
                                  <button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      exportCourseTimetable(cls, course);
                                    }}
                                    className="flex items-center gap-0.5 border border-[#0d0d0d]/20 bg-white px-1.5 py-0.5 text-[9px] font-bold uppercase tracking-wider text-[#2d2d2d] hover:bg-[#e8e4dd] transition active:translate-y-0.5"
                                    title="Download this course's timetable (Summary + Timetable)"
                                  >
                                    📥 Course
                                  </button>
                                  {courseFaculties.length > 0 && (
                                    <button
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        courseFaculties.forEach((f) => exportSingleFacultyExcel(f));
                                      }}
                                      className="flex items-center gap-0.5 border border-[#0369a1]/30 bg-[#e0f2fe] px-1.5 py-0.5 text-[9px] font-bold uppercase tracking-wider text-[#0369a1] hover:bg-[#bae6fd] transition active:translate-y-0.5"
                                      title="Download faculty timetable for this course's instructor(s)"
                                    >
                                      👤 Faculty
                                    </button>
                                  )}
                                </div>
                              </div>
                            );
                          })
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {/* Faculty Panel Backdrop Overlay */}
      {facultyPanelOpen && (
        <div
          className="fixed inset-0 z-[90] bg-[#0d0d0d]/35 backdrop-blur-[2px]"
          onClick={() => setFacultyPanelOpen(false)}
        />
      )}
      {/* Faculty Panel Drawer */}
      <div
        className={`fixed right-0 top-0 h-full w-[min(450px,100vw)] bg-[#f5f3ee] border-l-2 border-[#0d0d0d] shadow-[0_0_50px_rgba(0,0,0,0.3)] z-[95] flex flex-col transition-transform duration-300 ease-in-out ${
          facultyPanelOpen ? "translate-x-0" : "translate-x-full"
        }`}
      >
        <div className="flex items-center justify-between border-b-2 border-[#0d0d0d] bg-[#0d0d0d] px-4 py-3 text-[#f5f3ee]">
          <div>
            <div className="text-[10px] font-bold uppercase tracking-widest text-[#f5f3ee]/60">
              Control Panel
            </div>
            <div
              className="text-sm font-bold"
              style={{ fontFamily: "'Sora', system-ui, sans-serif" }}
            >
              Faculty List ({globalFaculties.length})
            </div>
          </div>
          <button
            onClick={() => setFacultyPanelOpen(false)}
            className="border-2 border-[#f5f3ee] bg-transparent px-3 py-1 text-[11px] font-bold uppercase tracking-wider text-[#f5f3ee] hover:bg-[#f5f3ee] hover:text-[#0d0d0d] transition"
          >
            Close
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          <p className="text-xs text-[#2d2d2d]/60 mb-2">
            Below is the list of all unique faculty members across all classes. Click **Rename** to
            update a faculty name globally.
          </p>
          {globalFaculties.length === 0 ? (
            <div className="text-xs text-[#2d2d2d]/50 italic p-4 border-2 border-dashed border-[#0d0d0d]/10 bg-white text-center">
              No faculties assigned yet. Set them on courses.
            </div>
          ) : (
            <div className="space-y-3">
              {globalFaculties.map((fac) => {
                const isSelected = selectedFaculty?.toLowerCase() === fac.toLowerCase();
                const coursesHandled = (() => {
                  const commonGroups = new Map<
                    string,
                    { classNames: string[]; courseName: string; weeklyPeriods: number }
                  >();
                  const individualList: Array<{
                    className: string;
                    courseName: string;
                    weeklyPeriods: number;
                  }> = [];

                  state.classes.forEach((cls) => {
                    cls.courses.forEach((c) => {
                      if (c.faculty) {
                        const names = c.faculty.split(",").map((f) => f.trim().toLowerCase());
                        if (names.includes(fac.toLowerCase())) {
                          if (c.common) {
                            const key = c.name.trim().toLowerCase();
                            const existing = commonGroups.get(key);
                            if (existing) {
                              if (!existing.classNames.includes(cls.name)) {
                                existing.classNames.push(cls.name);
                              }
                            } else {
                              commonGroups.set(key, {
                                classNames: [cls.name],
                                courseName: c.name,
                                weeklyPeriods: courseWeeklyTarget(c),
                              });
                            }
                          } else {
                            individualList.push({
                              className: cls.name,
                              courseName: c.name,
                              weeklyPeriods: courseWeeklyTarget(c),
                            });
                          }
                        }
                      }
                    });
                  });

                  return [
                    ...Array.from(commonGroups.values()).map((g) => ({
                      className: g.classNames.join(", "),
                      courseName: g.courseName + " (Common)",
                      weeklyPeriods: g.weeklyPeriods,
                    })),
                    ...individualList,
                  ];
                })();

                return (
                  <div
                    key={fac}
                    className="flex flex-col border border-[#0d0d0d]/20 bg-white shadow-[3px_3px_0px_0px_rgba(13,13,13,0.1)] transition-all"
                  >
                    <div
                      className="flex items-center gap-2 px-3 py-2 cursor-pointer hover:bg-[#e8e4dd]/30"
                      onClick={() => setSelectedFaculty(isSelected ? null : fac)}
                    >
                      {editingFaculty === fac ? (
                        <input
                          value={facultyEditVal}
                          onChange={(e) => setFacultyEditVal(e.target.value)}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") commitFacultyRename(fac, facultyEditVal.trim());
                            if (e.key === "Escape") setEditingFaculty(null);
                          }}
                          onClick={(e) => e.stopPropagation()}
                          className="min-w-0 flex-1 bg-transparent text-xs font-bold outline-none border-b border-[#0d0d0d]"
                          autoFocus
                        />
                      ) : (
                        <span className="min-w-0 flex-1 text-xs font-bold text-[#2d2d2d] truncate flex items-center gap-1.5">
                          <span className="text-[10px] text-[#2d2d2d]/40 font-mono">
                            {isSelected ? "▼" : "▶"}
                          </span>
                          <span>{fac}</span>
                          <span className="text-[9px] text-[#2d2d2d]/40 font-normal">
                            ({coursesHandled.length} course
                            {coursesHandled.length === 1 ? "" : "s"})
                          </span>
                        </span>
                      )}

                      {editingFaculty === fac ? (
                        <div
                          className="flex items-center gap-1.5 shrink-0"
                          onClick={(e) => e.stopPropagation()}
                        >
                          <button
                            onClick={() => commitFacultyRename(fac, facultyEditVal.trim())}
                            className="text-[10px] font-bold text-green-700 hover:text-green-900 border border-green-700/20 px-2 py-0.5 bg-green-50"
                          >
                            Save
                          </button>
                          <button
                            onClick={() => setEditingFaculty(null)}
                            className="text-[10px] font-bold text-red-700 hover:text-red-900 border border-red-700/20 px-2 py-0.5 bg-red-50"
                          >
                            Cancel
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            setEditingFaculty(fac);
                            setFacultyEditVal(fac);
                          }}
                          className="text-[10px] uppercase font-bold tracking-wider text-[#2d2d2d]/40 hover:text-[#0d0d0d] hover:underline shrink-0"
                        >
                          Rename
                        </button>
                      )}
                    </div>

                    {/* Collapsible Details Area */}
                    {isSelected && (
                      <div className="border-t border-[#0d0d0d]/10 bg-[#f5f3ee]/30 p-3 space-y-2 text-xs">
                        <div className="font-bold text-[#0d0d0d]/70 text-[10px] uppercase tracking-wider">
                          Handled Courses
                        </div>
                        {coursesHandled.length === 0 ? (
                          <div className="text-[11px] text-[#2d2d2d]/50 italic">
                            No courses currently configured for this faculty.
                          </div>
                        ) : (
                          <ul className="space-y-1 pl-1">
                            {coursesHandled.map((ch, idx) => (
                              <li
                                key={idx}
                                className="list-disc list-inside text-[11px] text-[#2d2d2d] leading-relaxed"
                              >
                                <span className="font-mono bg-[#0d0d0d]/5 px-1 rounded text-[10px] font-bold mr-1">
                                  {ch.className}
                                </span>{" "}
                                <span>{ch.courseName}</span>{" "}
                                <span className="text-[#2d2d2d]/60 font-semibold">
                                  ({ch.weeklyPeriods} periods/wk)
                                </span>
                              </li>
                            ))}
                          </ul>
                        )}

                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            exportSingleFacultyExcel(fac);
                          }}
                          className="mt-2 w-full flex items-center justify-center gap-1.5 border border-[#0d0d0d] bg-white px-2.5 py-1.5 text-[10px] font-bold uppercase tracking-wider text-[#0d0d0d] hover:bg-[#e8e4dd] transition active:translate-y-0.5 shadow-[1px_1px_0_0_#0d0d0d]"
                        >
                          📥 Download {fac}'s Excel
                        </button>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {courseContextMenu && (
        <div
          className="fixed z-[100] min-w-[140px] border-2 border-[#0d0d0d] bg-[#f5f3ee] py-1 shadow-[4px_4px_0px_0px_#0d0d0d]"
          style={{ left: courseContextMenu.x, top: courseContextMenu.y }}
          onClick={(e) => e.stopPropagation()}
        >
          <button
            onClick={() => {
              duplicateCourse(courseContextMenu.courseId);
              setCourseContextMenu(null);
            }}
            className="flex w-full items-center px-3 py-1.5 text-left text-xs font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
          >
            Duplicate
          </button>
          <button
            onClick={() => {
              clearCoursePlacements(courseContextMenu.courseId);
              setCourseContextMenu(null);
            }}
            className="flex w-full items-center px-3 py-1.5 text-left text-xs font-bold uppercase tracking-wider text-red-700 hover:bg-[#e8e4dd]"
          >
            Clear Placed
          </button>
          <button
            onClick={() => {
              setRulesFor(courseContextMenu.courseId);
              setCourseContextMenu(null);
            }}
            className="flex w-full items-center px-3 py-1.5 text-left text-xs font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
          >
            Rules
          </button>
          <button
            onClick={() => {
              if (confirm("Remove this course and delete all its placements?")) {
                removeCourse(courseContextMenu.courseId);
              }
              setCourseContextMenu(null);
            }}
            className="flex w-full items-center border-t border-dashed border-[#0d0d0d]/10 px-3 py-1.5 text-left text-xs font-bold uppercase tracking-wider text-red-600 hover:bg-[#e8e4dd]"
          >
            Delete
          </button>
        </div>
      )}

      {classContextMenu && (
        <div
          className="fixed z-[100] min-w-[140px] border-2 border-[#0d0d0d] bg-[#f5f3ee] py-1 shadow-[4px_4px_0px_0px_#0d0d0d]"
          style={{ left: classContextMenu.x, top: classContextMenu.y }}
          onClick={(e) => e.stopPropagation()}
        >
          <button
            onClick={() => {
              duplicateClass(classContextMenu.classId);
              setClassContextMenu(null);
            }}
            className="flex w-full items-center px-3 py-1.5 text-left text-xs font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
          >
            Duplicate Class
          </button>
          <button
            onClick={() => {
              const name = prompt("Enter new name for the class:");
              if (name && name.trim()) {
                renameClass(classContextMenu.classId, name.trim());
              }
              setClassContextMenu(null);
            }}
            className="flex w-full items-center px-3 py-1.5 text-left text-xs font-bold uppercase tracking-wider hover:bg-[#e8e4dd]"
          >
            Rename Class
          </button>
          <button
            onClick={() => {
              if (confirm("Are you sure you want to delete this class and all its data?")) {
                removeClass(classContextMenu.classId);
              }
              setClassContextMenu(null);
            }}
            className="flex w-full items-center border-t border-dashed border-[#0d0d0d]/10 px-3 py-1.5 text-left text-xs font-bold uppercase tracking-wider text-red-600 hover:bg-[#e8e4dd]"
          >
            Delete Class
          </button>
        </div>
      )}
    </div>
  );
}
