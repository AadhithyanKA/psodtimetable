import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Timetable Maker — Plan classes, faculty & breaks" },
      {
        name: "description",
        content:
          "Build weekly class timetables with faculty allocation, break scheduling, blocked slots, and Excel/CSV export.",
      },
      { property: "og:title", content: "Timetable Maker" },
      {
        property: "og:description",
        content:
          "Allocate courses to faculty across multiple classes with conflict detection. Export as Excel or CSV.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

const DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"] as const;
type Day = (typeof DAYS)[number];

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
};

type ClassData = {
  id: string;
  name: string;
  // key: `${day}-${slotIndex}` -> Cell
  grid: Record<string, Cell>;
};

type State = {
  slots: string[]; // time slot labels e.g. "09:00-10:00"
  days: Day[];
  courses: Course[];
  classes: ClassData[];
};

const COLORS = [
  "#fca5a5",
  "#fdba74",
  "#fcd34d",
  "#86efac",
  "#67e8f9",
  "#93c5fd",
  "#c4b5fd",
  "#f9a8d4",
];

const STORAGE_KEY = "timetable-maker-v1";

function defaultState(): State {
  const slots = [
    "09:00-10:00",
    "10:00-11:00",
    "11:00-12:00",
    "12:00-13:00",
    "13:00-14:00",
    "14:00-15:00",
    "15:00-16:00",
  ];
  const days: Day[] = ["Mon", "Tue", "Wed", "Thu", "Fri"];
  const courses: Course[] = [
    { id: "c1", name: "Mathematics", faculty: "Dr. Smith", color: COLORS[0] },
    { id: "c2", name: "Physics", faculty: "Dr. Jones", color: COLORS[3] },
    { id: "c3", name: "Chemistry", faculty: "Dr. Patel", color: COLORS[5] },
  ];
  const mkGrid = (): Record<string, Cell> => {
    const g: Record<string, Cell> = {};
    days.forEach((d) => {
      slots.forEach((_, i) => {
        g[`${d}-${i}`] = i === 3 ? { kind: "break", label: "Lunch" } : { kind: "empty" };
      });
    });
    return g;
  };
  return {
    slots,
    days,
    courses,
    classes: [
      { id: "k1", name: "Class A", grid: mkGrid() },
      { id: "k2", name: "Class B", grid: mkGrid() },
    ],
  };
}

function useHydrated() {
  const [h, setH] = useState(false);
  useEffect(() => setH(true), []);
  return h;
}

function Index() {
  const hydrated = useHydrated();
  const [state, setState] = useState<State>(defaultState);
  const [activeClassId, setActiveClassId] = useState<string>("k1");
  const [tool, setTool] = useState<
    | { kind: "course"; courseId: string }
    | { kind: "break" }
    | { kind: "blocked" }
    | { kind: "erase" }
  >({ kind: "erase" });

  useEffect(() => {
    if (!hydrated) return;
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as State;
        setState(parsed);
        setActiveClassId(parsed.classes[0]?.id ?? "");
      }
    } catch {
      /* ignore */
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [hydrated]);

  useEffect(() => {
    if (!hydrated) return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, [state, hydrated]);

  const activeClass = state.classes.find((c) => c.id === activeClassId) ?? state.classes[0];

  // Conflict map: for each `${day}-${slot}`, which faculty are used across classes
  const conflicts = useMemo(() => {
    const map: Record<string, Set<string>> = {};
    const conflictSet = new Set<string>(); // `${classId}:${day}-${slot}`
    state.classes.forEach((cls) => {
      Object.entries(cls.grid).forEach(([key, cell]) => {
        if (cell.kind !== "course") return;
        const course = state.courses.find((c) => c.id === cell.courseId);
        if (!course) return;
        if (!map[key]) map[key] = new Set();
        if (map[key].has(course.faculty)) {
          conflictSet.add(`${cls.id}:${key}`);
          // also mark previous class that used this faculty
          state.classes.forEach((c2) => {
            const cell2 = c2.grid[key];
            if (
              cell2 &&
              cell2.kind === "course" &&
              state.courses.find((c) => c.id === cell2.courseId)?.faculty === course.faculty
            ) {
              conflictSet.add(`${c2.id}:${key}`);
            }
          });
        }
        map[key].add(course.faculty);
      });
    });
    return conflictSet;
  }, [state]);

  const setCell = (day: Day, slotIdx: number) => {
    setState((s) => {
      const key = `${day}-${slotIdx}`;
      const classes = s.classes.map((cls) => {
        if (cls.id !== activeClassId) return cls;
        const grid = { ...cls.grid };
        if (tool.kind === "erase") grid[key] = { kind: "empty" };
        else if (tool.kind === "break") grid[key] = { kind: "break", label: "Break" };
        else if (tool.kind === "blocked") grid[key] = { kind: "blocked", label: "Blocked" };
        else grid[key] = { kind: "course", courseId: tool.courseId };
        return { ...cls, grid };
      });
      return { ...s, classes };
    });
  };

  const addClass = () => {
    setState((s) => {
      const id = `k${Date.now()}`;
      const grid: Record<string, Cell> = {};
      s.days.forEach((d) => s.slots.forEach((_, i) => (grid[`${d}-${i}`] = { kind: "empty" })));
      return {
        ...s,
        classes: [...s.classes, { id, name: `Class ${String.fromCharCode(65 + s.classes.length)}`, grid }],
      };
    });
  };

  const removeClass = (id: string) => {
    setState((s) => {
      const classes = s.classes.filter((c) => c.id !== id);
      if (classes.length === 0) return s;
      return { ...s, classes };
    });
    if (activeClassId === id) {
      const next = state.classes.find((c) => c.id !== id);
      if (next) setActiveClassId(next.id);
    }
  };

  const addCourse = () => {
    setState((s) => {
      const id = `c${Date.now()}`;
      const color = COLORS[s.courses.length % COLORS.length];
      return {
        ...s,
        courses: [...s.courses, { id, name: "New Course", faculty: "Faculty", color }],
      };
    });
  };

  const updateCourse = (id: string, patch: Partial<Course>) => {
    setState((s) => ({
      ...s,
      courses: s.courses.map((c) => (c.id === id ? { ...c, ...patch } : c)),
    }));
  };

  const removeCourse = (id: string) => {
    setState((s) => ({
      ...s,
      courses: s.courses.filter((c) => c.id !== id),
      classes: s.classes.map((cls) => {
        const grid = { ...cls.grid };
        Object.keys(grid).forEach((k) => {
          const cell = grid[k];
          if (cell.kind === "course" && cell.courseId === id) grid[k] = { kind: "empty" };
        });
        return { ...cls, grid };
      }),
    }));
    if (tool.kind === "course" && tool.courseId === id) setTool({ kind: "erase" });
  };

  const addSlot = () => {
    setState((s) => {
      const newSlot = `Slot ${s.slots.length + 1}`;
      const slots = [...s.slots, newSlot];
      const classes = s.classes.map((cls) => {
        const grid = { ...cls.grid };
        s.days.forEach((d) => (grid[`${d}-${slots.length - 1}`] = { kind: "empty" }));
        return { ...cls, grid };
      });
      return { ...s, slots, classes };
    });
  };

  const removeSlot = (i: number) => {
    setState((s) => {
      const slots = s.slots.filter((_, idx) => idx !== i);
      const classes = s.classes.map((cls) => {
        const grid: Record<string, Cell> = {};
        s.days.forEach((d) => {
          let newIdx = 0;
          s.slots.forEach((_, oldIdx) => {
            if (oldIdx === i) return;
            grid[`${d}-${newIdx}`] = cls.grid[`${d}-${oldIdx}`] ?? { kind: "empty" };
            newIdx++;
          });
        });
        return { ...cls, grid };
      });
      return { ...s, slots, classes };
    });
  };

  const updateSlot = (i: number, val: string) => {
    setState((s) => {
      const slots = [...s.slots];
      slots[i] = val;
      return { ...s, slots };
    });
  };

  const cellDisplay = (cell: Cell) => {
    if (cell.kind === "empty") return { text: "", bg: "transparent", fg: "#94a3b8" };
    if (cell.kind === "break") return { text: cell.label, bg: "#fef3c7", fg: "#92400e" };
    if (cell.kind === "blocked") return { text: cell.label, bg: "#e5e7eb", fg: "#374151" };
    const course = state.courses.find((c) => c.id === cell.courseId);
    return {
      text: course ? `${course.name}\n${course.faculty}` : "?",
      bg: course?.color ?? "#ddd",
      fg: "#1f2937",
    };
  };

  const buildSheet = (cls: ClassData) => {
    const rows: (string | number)[][] = [];
    rows.push(["Time", ...state.days]);
    state.slots.forEach((slot, i) => {
      const row: string[] = [slot];
      state.days.forEach((d) => {
        const cell = cls.grid[`${d}-${i}`];
        if (!cell || cell.kind === "empty") row.push("");
        else if (cell.kind === "break") row.push(`Break: ${cell.label}`);
        else if (cell.kind === "blocked") row.push(`Blocked: ${cell.label}`);
        else {
          const c = state.courses.find((x) => x.id === cell.courseId);
          row.push(c ? `${c.name} (${c.faculty})` : "");
        }
      });
      rows.push(row);
    });
    return rows;
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    state.classes.forEach((cls) => {
      const ws = XLSX.utils.aoa_to_sheet(buildSheet(cls));
      XLSX.utils.book_append_sheet(wb, ws, cls.name.slice(0, 31));
    });
    XLSX.writeFile(wb, "timetable.xlsx");
  };

  const exportCSV = () => {
    state.classes.forEach((cls) => {
      const rows = buildSheet(cls);
      const csv = rows
        .map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","))
        .join("\n");
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `timetable-${cls.name}.csv`;
      a.click();
      URL.revokeObjectURL(url);
    });
  };

  const hasConflicts = conflicts.size > 0;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <header className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-3 px-6 py-4">
          <div>
            <h1 className="text-2xl font-bold tracking-tight">Timetable Maker</h1>
            <p className="text-sm text-slate-500">
              Allocate courses, block time, add breaks. Faculty conflicts across classes are auto-flagged.
            </p>
          </div>
          <div className="flex gap-2">
            <button
              onClick={exportCSV}
              className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm font-medium hover:bg-slate-100"
            >
              Export CSV
            </button>
            <button
              onClick={exportExcel}
              className="rounded-md bg-slate-900 px-4 py-2 text-sm font-medium text-white hover:bg-slate-800"
            >
              Export Excel
            </button>
          </div>
        </div>
      </header>

      <main className="mx-auto grid max-w-7xl gap-6 px-6 py-6 lg:grid-cols-[280px_1fr]">
        {/* Sidebar */}
        <aside className="space-y-6">
          <section className="rounded-lg border border-slate-200 bg-white p-4">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-sm font-semibold uppercase tracking-wide text-slate-600">Classes</h2>
              <button
                onClick={addClass}
                className="rounded bg-slate-900 px-2 py-1 text-xs font-medium text-white hover:bg-slate-800"
              >
                + Add
              </button>
            </div>
            <div className="space-y-2">
              {state.classes.map((cls) => (
                <div key={cls.id} className="flex items-center gap-2">
                  <button
                    onClick={() => setActiveClassId(cls.id)}
                    className={`flex-1 rounded-md border px-3 py-2 text-left text-sm ${
                      cls.id === activeClassId
                        ? "border-slate-900 bg-slate-900 text-white"
                        : "border-slate-200 bg-white hover:bg-slate-50"
                    }`}
                  >
                    <input
                      value={cls.name}
                      onClick={(e) => e.stopPropagation()}
                      onChange={(e) =>
                        setState((s) => ({
                          ...s,
                          classes: s.classes.map((c) =>
                            c.id === cls.id ? { ...c, name: e.target.value } : c,
                          ),
                        }))
                      }
                      className={`w-full bg-transparent outline-none ${
                        cls.id === activeClassId ? "text-white" : ""
                      }`}
                    />
                  </button>
                  {state.classes.length > 1 && (
                    <button
                      onClick={() => removeClass(cls.id)}
                      className="rounded px-2 py-1 text-xs text-slate-500 hover:bg-red-50 hover:text-red-600"
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
              <button
                onClick={addCourse}
                className="rounded bg-slate-900 px-2 py-1 text-xs font-medium text-white hover:bg-slate-800"
              >
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
                  <div className="flex items-start gap-2">
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
                  <button
                    onClick={() => setTool({ kind: "course", courseId: c.id })}
                    className={`mt-2 w-full rounded px-2 py-1 text-xs font-medium ${
                      tool.kind === "course" && tool.courseId === c.id
                        ? "bg-slate-900 text-white"
                        : "bg-slate-100 hover:bg-slate-200"
                    }`}
                  >
                    {tool.kind === "course" && tool.courseId === c.id ? "Selected" : "Use"}
                  </button>
                </div>
              ))}
            </div>
          </section>

          <section className="rounded-lg border border-slate-200 bg-white p-4">
            <h2 className="mb-3 text-sm font-semibold uppercase tracking-wide text-slate-600">Tools</h2>
            <div className="grid grid-cols-3 gap-2">
              {(
                [
                  { k: "break", label: "Break", bg: "#fef3c7" },
                  { k: "blocked", label: "Block", bg: "#e5e7eb" },
                  { k: "erase", label: "Erase", bg: "#fff" },
                ] as const
              ).map((t) => (
                <button
                  key={t.k}
                  onClick={() => setTool({ kind: t.k })}
                  className={`rounded border px-2 py-2 text-xs font-medium ${
                    tool.kind === t.k ? "border-slate-900 ring-2 ring-slate-900" : "border-slate-200"
                  }`}
                  style={{ backgroundColor: t.bg }}
                >
                  {t.label}
                </button>
              ))}
            </div>
            <p className="mt-3 text-xs text-slate-500">
              Pick a tool or course, then click a grid cell to apply.
            </p>
          </section>
        </aside>

        {/* Grid */}
        <section className="space-y-3">
          {hasConflicts && (
            <div className="rounded-md border border-red-200 bg-red-50 px-4 py-2 text-sm text-red-700">
              Faculty conflict detected — same teacher scheduled in two classes at the same time (highlighted red).
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
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr>
                    <th className="w-40 border border-slate-200 bg-slate-50 p-2 text-left text-xs font-semibold uppercase tracking-wide text-slate-600">
                      Time
                    </th>
                    {state.days.map((d) => (
                      <th
                        key={d}
                        className="border border-slate-200 bg-slate-50 p-2 text-xs font-semibold uppercase tracking-wide text-slate-600"
                      >
                        {d}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {state.slots.map((slot, i) => (
                    <tr key={i}>
                      <td className="border border-slate-200 p-1">
                        <div className="flex items-center gap-1">
                          <input
                            value={slot}
                            onChange={(e) => updateSlot(i, e.target.value)}
                            className="w-full rounded px-1 py-1 text-xs"
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
                      </td>
                      {state.days.map((d) => {
                        if (!activeClass) return null;
                        const cell = activeClass.grid[`${d}-${i}`] ?? { kind: "empty" };
                        const disp = cellDisplay(cell);
                        const isConflict = conflicts.has(`${activeClass.id}:${d}-${i}`);
                        return (
                          <td
                            key={d}
                            onClick={() => setCell(d, i)}
                            className={`cursor-pointer border p-2 text-xs ${
                              isConflict ? "border-red-500 ring-2 ring-red-400" : "border-slate-200"
                            }`}
                            style={{ backgroundColor: disp.bg, color: disp.fg, minWidth: 110, height: 60 }}
                          >
                            <div className="whitespace-pre-line font-medium">{disp.text || "—"}</div>
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </section>
      </main>
    </div>
  );
}
