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

type Course = { id: string; name: string; faculty: string; color: string };
type ClassData = { id: string; name: string; grid: Record<string, Cell> };
type State = {
  fromDate: string; // YYYY-MM-DD
  toDate: string;
  slots: string[]; // column labels
  courses: Course[];
  classes: ClassData[];
};

const COLORS = [
  "#fca5a5", "#fdba74", "#fcd34d", "#86efac",
  "#67e8f9", "#93c5fd", "#c4b5fd", "#f9a8d4",
];
const STORAGE_KEY = "timetable-maker-v2";

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

function defaultState(): State {
  const from = isoToday();
  const to = addDays(from, 4);
  const slots = [
    "09:00-10:00", "10:00-11:00", "11:00-12:00",
    "12:00-13:00", "13:00-14:00", "14:00-15:00", "15:00-16:00",
  ];
  const courses: Course[] = [
    { id: "c1", name: "Mathematics", faculty: "Dr. Smith", color: COLORS[0] },
    { id: "c2", name: "Physics", faculty: "Dr. Jones", color: COLORS[3] },
    { id: "c3", name: "Chemistry", faculty: "Dr. Patel", color: COLORS[5] },
  ];
  const mkGrid = (): Record<string, Cell> => ({});
  return {
    fromDate: from,
    toDate: to,
    slots,
    courses,
    classes: [
      { id: "k1", name: "Class A", grid: mkGrid() },
      { id: "k2", name: "Class B", grid: mkGrid() },
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
  const [dragging, setDragging] = useState<Tool | null>(null);
  const gridRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setHydrated(true);
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as State;
        setState(parsed);
        setActiveClassId(parsed.classes[0]?.id ?? "");
      }
    } catch {}
  }, []);
  useEffect(() => {
    if (hydrated) localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, [state, hydrated]);

  useEffect(() => {
    const up = () => setDragging(null);
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
            const course = state.courses.find((c) => c.id === cell.courseId);
            if (!course) return;
            (facultyToClass[course.faculty] ??= []).push(cls.id);
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
    setState((s) => ({
      ...s,
      classes: s.classes.map((cls) => {
        if (cls.id !== activeClassId) return cls;
        const grid = { ...cls.grid };
        const key = `${date}-${slotIdx}`;
        if (tool.kind === "erase") delete grid[key];
        else if (tool.kind === "break") grid[key] = { kind: "break", label: "Break" };
        else if (tool.kind === "blocked") grid[key] = { kind: "blocked", label: "Blocked" };
        else grid[key] = { kind: "course", courseId: tool.courseId };
        return { ...cls, grid };
      }),
    }));
  };

  const onCellClick = (date: string, slotIdx: number) => {
    setPicker({ date, slotIdx });
  };
  const onCellMouseDown = (date: string, slotIdx: number, e: React.MouseEvent) => {
    // Shift-drag or right-click drag would be nice, but keep simple:
    // If popover open elsewhere, close. Left click begins picker; drag begins only after picking.
    // Here we support drag AFTER a tool has been chosen via picker (dragging state set).
    if (dragging) {
      e.preventDefault();
      applyTool(date, slotIdx, dragging);
    }
  };
  const onCellEnter = (date: string, slotIdx: number) => {
    if (dragging) applyTool(date, slotIdx, dragging);
  };

  const pickTool = (tool: Tool) => {
    if (!picker) return;
    applyTool(picker.date, picker.slotIdx, tool);
    setDragging(tool); // enable drag-to-paint on subsequent cells
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
          { id, name: `Class ${String.fromCharCode(65 + s.classes.length)}`, grid: {} },
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

  const addCourse = () =>
    setState((s) => ({
      ...s,
      courses: [
        ...s.courses,
        {
          id: `c${Date.now()}`,
          name: "New Course",
          faculty: "Faculty",
          color: COLORS[s.courses.length % COLORS.length],
        },
      ],
    }));
  const updateCourse = (id: string, patch: Partial<Course>) =>
    setState((s) => ({ ...s, courses: s.courses.map((c) => (c.id === id ? { ...c, ...patch } : c)) }));
  const removeCourse = (id: string) =>
    setState((s) => ({
      ...s,
      courses: s.courses.filter((c) => c.id !== id),
      classes: s.classes.map((cls) => {
        const grid = { ...cls.grid };
        Object.keys(grid).forEach((k) => {
          const cell = grid[k];
          if (cell.kind === "course" && cell.courseId === id) delete grid[k];
        });
        return { ...cls, grid };
      }),
    }));

  const addSlot = () =>
    setState((s) => ({ ...s, slots: [...s.slots, `Slot ${s.slots.length + 1}`] }));
  const removeSlot = (i: number) =>
    setState((s) => {
      const slots = s.slots.filter((_, idx) => idx !== i);
      const classes = s.classes.map((cls) => {
        const grid: Record<string, Cell> = {};
        Object.entries(cls.grid).forEach(([k, v]) => {
          const m = k.match(/^(.+)-(\d+)$/);
          if (!m) return;
          const oldIdx = parseInt(m[2], 10);
          if (oldIdx === i) return;
          const newIdx = oldIdx > i ? oldIdx - 1 : oldIdx;
          grid[`${m[1]}-${newIdx}`] = v;
        });
        return { ...cls, grid };
      });
      return { ...s, slots, classes };
    });
  const updateSlot = (i: number, val: string) =>
    setState((s) => {
      const slots = [...s.slots];
      slots[i] = val;
      return { ...s, slots };
    });

  // Export
  const buildSheet = (cls: ClassData) => {
    const rows: string[][] = [];
    rows.push(["Day / Date", ...state.slots]);
    dates.forEach((date) => {
      const { weekday, date: dstr } = dayLabel(date);
      const row = [`${weekday} ${dstr}`];
      state.slots.forEach((_, i) => {
        const cell = cls.grid[`${date}-${i}`];
        if (!cell) row.push("");
        else if (cell.kind === "break") row.push(`Break: ${cell.label}`);
        else if (cell.kind === "blocked") row.push(`Blocked: ${cell.label}`);
        else if (cell.kind === "course") {
          const c = state.courses.find((x) => x.id === cell.courseId);
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

  const cellDisplay = (cell: Cell | undefined) => {
    if (!cell || cell.kind === "empty") return { text: "", bg: "#fff", fg: "#94a3b8" };
    if (cell.kind === "break") return { text: cell.label, bg: "#fef3c7", fg: "#92400e" };
    if (cell.kind === "blocked") return { text: cell.label, bg: "#e5e7eb", fg: "#374151" };
    const course = state.courses.find((c) => c.id === cell.courseId);
    return {
      text: course ? `${course.name}\n${course.faculty}` : "?",
      bg: course?.color ?? "#ddd",
      fg: "#1f2937",
    };
  };

  const hasConflicts = conflicts.size > 0;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900" onMouseLeave={() => setDragging(null)}>
      <header className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex max-w-[1400px] flex-wrap items-center justify-between gap-3 px-6 py-4">
          <div>
            <h1 className="text-2xl font-bold tracking-tight">Timetable Maker</h1>
            <p className="text-sm text-slate-500">
              Pick a date range, click a slot to assign, then drag to fill more.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-3">
            <label className="flex items-center gap-2 text-xs font-medium text-slate-600">
              From
              <input
                type="date"
                value={state.fromDate}
                onChange={(e) => setState((s) => ({ ...s, fromDate: e.target.value }))}
                className="rounded border border-slate-300 px-2 py-1 text-sm"
              />
            </label>
            <label className="flex items-center gap-2 text-xs font-medium text-slate-600">
              To
              <input
                type="date"
                value={state.toDate}
                onChange={(e) => setState((s) => ({ ...s, toDate: e.target.value }))}
                className="rounded border border-slate-300 px-2 py-1 text-sm"
              />
            </label>
            <button
              onClick={exportCSV}
              className="rounded-md border border-slate-300 bg-white px-3 py-1.5 text-sm font-medium hover:bg-slate-100"
            >
              Export CSV
            </button>
            <button
              onClick={exportExcel}
              className="rounded-md bg-slate-900 px-3 py-1.5 text-sm font-medium text-white hover:bg-slate-800"
            >
              Export Excel
            </button>
          </div>
        </div>
      </header>

      <main className="mx-auto grid max-w-[1400px] gap-6 px-6 py-6 lg:grid-cols-[280px_1fr]">
        <aside className="space-y-4">
          <section className="rounded-lg border border-slate-200 bg-white p-4">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-sm font-semibold uppercase tracking-wide text-slate-600">Classes</h2>
              <button onClick={addClass} className="rounded bg-slate-900 px-2 py-1 text-xs text-white hover:bg-slate-800">
                + Add
              </button>
            </div>
            <div className="space-y-2">
              {state.classes.map((cls) => (
                <div key={cls.id} className="flex items-center gap-1">
                  <button
                    onClick={() => setActiveClassId(cls.id)}
                    className={`flex-1 rounded-md border px-2 py-2 text-left text-sm ${
                      cls.id === activeClassId
                        ? "border-slate-900 bg-slate-900 text-white"
                        : "border-slate-200 bg-white hover:bg-slate-50"
                    }`}
                  >
                    <input
                      value={cls.name}
                      onClick={(e) => e.stopPropagation()}
                      onChange={(e) => renameClass(cls.id, e.target.value)}
                      className={`w-full bg-transparent outline-none ${
                        cls.id === activeClassId ? "text-white" : ""
                      }`}
                    />
                  </button>
                  {state.classes.length > 1 && (
                    <button
                      onClick={() => removeClass(cls.id)}
                      className="rounded px-2 py-1 text-xs text-slate-400 hover:bg-red-50 hover:text-red-600"
                    >
                      ×
                    </button>
                  )}
                </div>
              ))}
            </div>
          </section>

          <section className="rounded-lg border border-slate-200 bg-white p-4">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-sm font-semibold uppercase tracking-wide text-slate-600">Courses</h2>
              <button onClick={addCourse} className="rounded bg-slate-900 px-2 py-1 text-xs text-white hover:bg-slate-800">
                + Add
              </button>
            </div>
            <div className="space-y-2">
              {state.courses.map((c) => (
                <div
                  key={c.id}
                  className="rounded-md border border-slate-200 p-2"
                  style={{ borderLeftWidth: 4, borderLeftColor: c.color }}
                >
                  <div className="flex items-start gap-1">
                    <div className="flex-1 space-y-1">
                      <input
                        value={c.name}
                        onChange={(e) => updateCourse(c.id, { name: e.target.value })}
                        placeholder="Course"
                        className="w-full rounded border border-slate-200 px-2 py-1 text-sm"
                      />
                      <input
                        value={c.faculty}
                        onChange={(e) => updateCourse(c.id, { faculty: e.target.value })}
                        placeholder="Faculty"
                        className="w-full rounded border border-slate-200 px-2 py-1 text-xs text-slate-600"
                      />
                    </div>
                    <button
                      onClick={() => removeCourse(c.id)}
                      className="rounded px-2 py-1 text-xs text-slate-400 hover:bg-red-50 hover:text-red-600"
                    >
                      ×
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </section>

          <section className="rounded-lg border border-slate-200 bg-white p-4 text-xs text-slate-600">
            <p className="font-semibold text-slate-700">How to use</p>
            <ol className="mt-2 list-decimal space-y-1 pl-4">
              <li>Click any empty cell → pick a course, break, or block.</li>
              <li>Then click-and-drag across cells to paint the same choice.</li>
              <li>Faculty double-booked across classes gets flagged red.</li>
            </ol>
          </section>
        </aside>

        <section className="space-y-3">
          {hasConflicts && (
            <div className="rounded-md border border-red-200 bg-red-50 px-4 py-2 text-sm text-red-700">
              Faculty conflict — same teacher scheduled in two classes at the same time (highlighted red).
            </div>
          )}
          {dates.length === 0 && (
            <div className="rounded-md border border-amber-200 bg-amber-50 px-4 py-2 text-sm text-amber-700">
              Pick a valid date range.
            </div>
          )}

          <div className="rounded-lg border border-slate-200 bg-white p-4">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-lg font-semibold">{activeClass?.name}</h2>
              <button
                onClick={addSlot}
                className="rounded border border-slate-300 px-2 py-1 text-xs hover:bg-slate-50"
              >
                + Add time slot
              </button>
            </div>

            <div className="overflow-x-auto" ref={gridRef}>
              <table className="w-full border-collapse select-none text-sm">
                <thead>
                  <tr>
                    <th className="sticky left-0 z-10 w-44 border border-slate-200 bg-slate-100 p-2 text-left text-xs font-semibold uppercase tracking-wide text-slate-600">
                      Day
                    </th>
                    {state.slots.map((slot, i) => (
                      <th
                        key={i}
                        className="border border-slate-200 bg-slate-100 p-1 text-xs font-semibold text-slate-600"
                        style={{ minWidth: 120 }}
                      >
                        <div className="flex items-center gap-1">
                          <input
                            value={slot}
                            onChange={(e) => updateSlot(i, e.target.value)}
                            className="w-full rounded bg-white px-1 py-1 text-xs"
                          />
                          {state.slots.length > 1 && (
                            <button
                              onClick={() => removeSlot(i)}
                              className="rounded px-1 text-xs text-slate-400 hover:text-red-600"
                            >
                              ×
                            </button>
                          )}
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {dates.map((date) => {
                    const { weekday, date: dstr } = dayLabel(date);
                    return (
                      <tr key={date}>
                        <th className="sticky left-0 z-10 border border-slate-200 bg-slate-50 p-2 text-left align-middle text-xs">
                          <div className="font-semibold text-slate-800">{weekday}</div>
                          <div className="text-slate-500">{dstr}</div>
                        </th>
                        {state.slots.map((_, i) => {
                          if (!activeClass) return null;
                          const key = `${date}-${i}`;
                          const cell = activeClass.grid[key];
                          const disp = cellDisplay(cell);
                          const isConflict = conflicts.has(`${activeClass.id}:${key}`);
                          return (
                            <td
                              key={i}
                              onClick={() => onCellClick(date, i)}
                              onMouseDown={(e) => onCellMouseDown(date, i, e)}
                              onMouseEnter={() => onCellEnter(date, i)}
                              className={`cursor-pointer border p-2 text-xs align-middle ${
                                isConflict ? "border-red-500 ring-2 ring-red-400" : "border-slate-200"
                              }`}
                              style={{
                                backgroundColor: disp.bg,
                                color: disp.fg,
                                minWidth: 120,
                                height: 56,
                              }}
                            >
                              <div className="whitespace-pre-line font-medium leading-tight">
                                {disp.text || "+"}
                              </div>
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
        </section>
      </main>

      {picker && (
        <div
          className="fixed inset-0 z-40 bg-black/30"
          onClick={() => setPicker(null)}
        >
          <div
            className="absolute left-1/2 top-1/2 w-[360px] -translate-x-1/2 -translate-y-1/2 rounded-lg border border-slate-200 bg-white p-4 shadow-xl"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="mb-3">
              <div className="text-xs uppercase tracking-wide text-slate-500">Assign slot</div>
              <div className="text-sm font-semibold text-slate-800">
                {dayLabel(picker.date).weekday} {dayLabel(picker.date).date} · {state.slots[picker.slotIdx]}
              </div>
              <div className="text-xs text-slate-500">
                Tip: after picking, drag across cells to fill more with the same choice.
              </div>
            </div>
            <div className="mb-3 space-y-1">
              <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Courses</div>
              {state.courses.length === 0 && (
                <div className="text-xs text-slate-500">No courses yet. Add one from the sidebar.</div>
              )}
              {state.courses.map((c) => (
                <button
                  key={c.id}
                  onClick={() => pickTool({ kind: "course", courseId: c.id })}
                  className="flex w-full items-center gap-2 rounded-md border border-slate-200 px-3 py-2 text-left text-sm hover:bg-slate-50"
                >
                  <span className="h-4 w-4 rounded" style={{ backgroundColor: c.color }} />
                  <span className="flex-1">
                    <span className="font-medium">{c.name}</span>
                    <span className="ml-2 text-xs text-slate-500">{c.faculty}</span>
                  </span>
                </button>
              ))}
            </div>
            <div className="grid grid-cols-3 gap-2">
              <button
                onClick={() => pickTool({ kind: "break" })}
                className="rounded-md border border-slate-200 px-2 py-2 text-xs font-medium"
                style={{ backgroundColor: "#fef3c7", color: "#92400e" }}
              >
                Break
              </button>
              <button
                onClick={() => pickTool({ kind: "blocked" })}
                className="rounded-md border border-slate-200 px-2 py-2 text-xs font-medium"
                style={{ backgroundColor: "#e5e7eb", color: "#374151" }}
              >
                Block
              </button>
              <button
                onClick={() => pickTool({ kind: "erase" })}
                className="rounded-md border border-slate-200 bg-white px-2 py-2 text-xs font-medium hover:bg-slate-50"
              >
                Clear
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
