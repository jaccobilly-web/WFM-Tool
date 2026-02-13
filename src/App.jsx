import { useState, useCallback } from "react";
import { exportToExcel } from "./export";

let _id = 0;
const uid = () => `id_${++_id}_${Date.now()}`;

const makeCategory = (name = "") => ({ id: uid(), name, weight: 0, criteria: [], hasSubcriteria: true });
const makeCriterion = (name = "") => ({ id: uid(), name, weight: 0, description: "" });

const DEFAULT_CATEGORIES = [
  { ...makeCategory(""), criteria: [makeCriterion("")] },
  { ...makeCategory(""), criteria: [makeCriterion("")] },
  { ...makeCategory(""), criteria: [makeCriterion("")] },
];

function WeightBar({ value, total, color = "emerald" }) {
  const pct = total > 0 ? Math.min((value / total) * 100, 100) : 0;
  const colors = { emerald: "bg-emerald-500", amber: "bg-amber-500", red: "bg-red-500" };
  return (
    <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
      <div className={`h-full rounded-full transition-all duration-300 ${colors[color]}`} style={{ width: `${pct}%` }} />
    </div>
  );
}

function WeightSummary({ categories }) {
  const total = categories.reduce((s, c) => s + c.weight, 0);
  const balanced = total === 100;
  const color = balanced ? "emerald" : total > 100 ? "red" : "amber";
  const label = balanced ? "Balanced" : total > 100 ? `${total - 100}% over` : `${100 - total}% remaining`;

  return (
    <div className={`rounded-xl p-4 border-2 transition-colors ${balanced ? "border-emerald-300 bg-emerald-50" : total > 100 ? "border-red-300 bg-red-50" : "border-amber-300 bg-amber-50"}`}>
      <div className="flex items-center justify-between mb-2">
        <span className="text-sm font-semibold text-slate-700">Category Weights</span>
        <div className="flex items-center gap-2">
          <span className={`text-2xl font-bold ${balanced ? "text-emerald-700" : total > 100 ? "text-red-700" : "text-amber-700"}`}>{total}%</span>
          <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${balanced ? "bg-emerald-200 text-emerald-800" : total > 100 ? "bg-red-200 text-red-800" : "bg-amber-200 text-amber-800"}`}>{label}</span>
        </div>
      </div>
      <WeightBar value={total} total={100} color={color} />
      <div className="flex gap-1 mt-3">
        {categories.map(cat => {
          const pct = cat.weight;
          const lbl = cat.name || "Untitled";
          return pct > 0 ? (
            <div key={cat.id} className="rounded-sm transition-all duration-300 overflow-hidden flex flex-col items-center justify-center py-1"
              style={{ width: `${pct}%`, backgroundColor: balanced ? "#10b981" : total > 100 ? "#ef4444" : "#f59e0b", opacity: 0.6 + (pct / 200), minHeight: "36px" }}>
              <span className="text-[10px] text-white font-semibold leading-tight truncate w-full text-center px-1">{lbl}</span>
              <span className="text-[9px] text-white/80 font-bold">{pct}%</span>
            </div>
          ) : null;
        })}
      </div>
    </div>
  );
}

function CriterionRow({ criterion, onChange, onRemove }) {
  return (
    <div className="flex items-start gap-3 py-2.5 group">
      <div className="flex-1 min-w-0">
        <input type="text" value={criterion.name} placeholder="Criterion name"
          onChange={e => onChange({ ...criterion, name: e.target.value })}
          className="w-full text-sm font-medium text-slate-800 bg-transparent border-0 border-b border-transparent hover:border-slate-200 focus:border-slate-400 focus:outline-none px-0 py-0.5 transition-colors" />
        <input type="text" value={criterion.description} placeholder="Brief description (optional)"
          onChange={e => onChange({ ...criterion, description: e.target.value })}
          className="w-full text-xs text-slate-400 bg-transparent border-0 border-b border-transparent hover:border-slate-200 focus:border-slate-400 focus:outline-none px-0 py-0.5 mt-0.5 transition-colors" />
      </div>
      <div className="flex items-center gap-2 shrink-0">
        <input type="number" min="0" max="100" value={criterion.weight}
          onChange={e => onChange({ ...criterion, weight: Math.max(0, Math.min(100, parseInt(e.target.value) || 0)) })}
          className="w-16 text-sm text-right font-semibold text-slate-700 border border-slate-200 rounded-lg px-2 py-1.5 focus:outline-none focus:border-slate-400 focus:ring-1 focus:ring-slate-200" />
        <span className="text-xs text-slate-400 w-3">%</span>
        <button onClick={onRemove} className="opacity-0 group-hover:opacity-100 text-slate-300 hover:text-red-500 transition-all p-1">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M18 6L6 18M6 6l12 12"/></svg>
        </button>
      </div>
    </div>
  );
}

function CategoryCard({ category, onChange, onRemove, index }) {
  const criteriaTotal = category.criteria.reduce((s, c) => s + c.weight, 0);
  const criteriaBalanced = criteriaTotal === 100;
  const criteriaColor = criteriaBalanced ? "emerald" : criteriaTotal > 100 ? "red" : "amber";

  const toggleSubcriteria = () => {
    if (category.hasSubcriteria) {
      onChange({ ...category, hasSubcriteria: false, criteria: [] });
    } else {
      onChange({ ...category, hasSubcriteria: true, criteria: [makeCriterion("")] });
    }
  };

  const updateCriterion = (critId, updated) => onChange({ ...category, criteria: category.criteria.map(c => c.id === critId ? updated : c) });
  const removeCriterion = (critId) => onChange({ ...category, criteria: category.criteria.filter(c => c.id !== critId) });
  const addCriterion = () => onChange({ ...category, criteria: [...category.criteria, makeCriterion()] });
  const autoBalanceCriteria = () => {
    const count = category.criteria.length;
    if (count === 0) return;
    const base = Math.floor(100 / count);
    const remainder = 100 - base * count;
    onChange({ ...category, criteria: category.criteria.map((c, i) => ({ ...c, weight: base + (i < remainder ? 1 : 0) })) });
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
      <div className="px-5 py-4 bg-slate-50 border-b border-slate-200">
        <div className="flex items-center gap-3">
          <span className="text-xs font-bold text-slate-400 bg-slate-200 w-6 h-6 rounded-full flex items-center justify-center">{index + 1}</span>
          <input type="text" value={category.name} placeholder="Category name"
            onChange={e => onChange({ ...category, name: e.target.value })}
            className="flex-1 text-base font-semibold text-slate-800 bg-transparent border-0 border-b-2 border-transparent hover:border-slate-300 focus:border-slate-500 focus:outline-none px-0 py-0.5 transition-colors" />
          <div className="flex items-center gap-2 shrink-0">
            <span className="text-xs text-slate-500">Category weight:</span>
            <input type="number" min="0" max="100" value={category.weight}
              onChange={e => onChange({ ...category, weight: Math.max(0, Math.min(100, parseInt(e.target.value) || 0)) })}
              className="w-16 text-sm text-right font-bold text-slate-700 border border-slate-200 rounded-lg px-2 py-1.5 focus:outline-none focus:border-slate-400 focus:ring-1 focus:ring-slate-200" />
            <span className="text-xs text-slate-400">%</span>
          </div>
          <button onClick={onRemove} className="text-slate-300 hover:text-red-500 transition-colors p-1 ml-1">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M18 6L6 18M6 6l12 12"/></svg>
          </button>
        </div>
        <div className="mt-2">
          <label className="flex items-center gap-2 cursor-pointer">
            <input type="checkbox" checked={category.hasSubcriteria} onChange={toggleSubcriteria}
              className="rounded border-slate-300 text-slate-600 focus:ring-slate-400" />
            <span className="text-xs text-slate-500">Has sub-criteria</span>
          </label>
        </div>
      </div>

      {category.hasSubcriteria && (
        <div className="px-5 py-3">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs text-slate-500">Criteria weights within this category</span>
            <div className="flex items-center gap-3">
              <button onClick={autoBalanceCriteria} className="text-xs text-slate-500 hover:text-slate-700 underline decoration-dotted">Auto-balance</button>
              <span className={`text-xs font-bold ${criteriaBalanced ? "text-emerald-600" : criteriaTotal > 100 ? "text-red-600" : "text-amber-600"}`}>{criteriaTotal}%</span>
            </div>
          </div>
          <WeightBar value={criteriaTotal} total={100} color={criteriaColor} />
          <div className="mt-3 divide-y divide-slate-100">
            {category.criteria.map(crit => (
              <CriterionRow key={crit.id} criterion={crit}
                onChange={updated => updateCriterion(crit.id, updated)}
                onRemove={() => removeCriterion(crit.id)} />
            ))}
          </div>
          <button onClick={addCriterion} className="mt-3 text-xs text-slate-500 hover:text-slate-700 flex items-center gap-1 py-1.5">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 5v14M5 12h14"/></svg>
            Add criterion
          </button>
        </div>
      )}

      {!category.hasSubcriteria && (
        <div className="px-5 py-3">
          <p className="text-xs text-slate-400 italic">This category has no sub-criteria. It will appear as a single column in the spreadsheet.</p>
        </div>
      )}
    </div>
  );
}

function EffectiveWeightPreview({ categories }) {
  const totalCatWeight = categories.reduce((s, c) => s + c.weight, 0);
  const allCriteria = [];
  categories.forEach(cat => {
    if (cat.hasSubcriteria && cat.criteria.length > 0) {
      const critTotal = cat.criteria.reduce((s, c) => s + c.weight, 0);
      cat.criteria.forEach(crit => {
        const ew = critTotal > 0 ? (cat.weight / (totalCatWeight || 100)) * (crit.weight / critTotal) * 100 : 0;
        allCriteria.push({ category: cat.name || "Untitled", name: crit.name || "Untitled", effectiveWeight: Math.round(ew * 10) / 10 });
      });
    } else {
      const ew = totalCatWeight > 0 ? (cat.weight / totalCatWeight) * 100 : 0;
      allCriteria.push({ category: cat.name || "Untitled", name: cat.name || "Untitled", effectiveWeight: Math.round(ew * 10) / 10 });
    }
  });
  allCriteria.sort((a, b) => b.effectiveWeight - a.effectiveWeight);
  const maxWeight = Math.max(...allCriteria.map(c => c.effectiveWeight), 1);

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-5">
      <h3 className="text-sm font-semibold text-slate-700 mb-3">Effective Weights Preview</h3>
      <p className="text-xs text-slate-400 mb-4">How each criterion contributes to the overall score (category weight x criterion weight, normalised to 100%)</p>
      <div className="space-y-2">
        {allCriteria.map((c, i) => (
          <div key={i} className="flex items-center gap-3">
            <span className="text-xs text-slate-500 w-28 shrink-0 truncate" title={c.category}>{c.category}</span>
            <span className="text-xs font-medium text-slate-700 w-36 shrink-0 truncate" title={c.name}>{c.name}</span>
            <div className="flex-1">
              <div className="h-5 bg-slate-50 rounded-full overflow-hidden">
                <div className="h-full rounded-full bg-emerald-500 transition-all duration-300 flex items-center justify-end pr-2"
                  style={{ width: `${Math.max((c.effectiveWeight / maxWeight) * 100, 8)}%`, opacity: 0.5 + (c.effectiveWeight / maxWeight) * 0.5 }}>
                  {c.effectiveWeight >= 3 && <span className="text-[10px] font-bold text-white">{c.effectiveWeight}%</span>}
                </div>
              </div>
            </div>
            <span className="text-xs font-bold text-slate-700 w-12 text-right">{c.effectiveWeight}%</span>
          </div>
        ))}
      </div>
      {allCriteria.length === 0 && <p className="text-xs text-slate-400 italic">Add categories and criteria to see effective weights</p>}
    </div>
  );
}

function OptionNamer({ numOptions, optionNames, setOptionNames }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
      <button onClick={() => setOpen(!open)} className="w-full flex items-center justify-between px-5 py-3 hover:bg-slate-50 transition-colors">
        <span className="text-sm font-medium text-slate-700">Name your options (optional)</span>
        <svg className={`w-4 h-4 text-slate-400 transition-transform ${open ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M6 9l6 6 6-6"/></svg>
      </button>
      {open && (
        <div className="px-5 pb-4 space-y-2">
          <p className="text-xs text-slate-400 mb-2">Leave blank to use "Option 1", "Option 2", etc.</p>
          {Array.from({ length: numOptions }, (_, i) => (
            <div key={i} className="flex items-center gap-3">
              <span className="text-xs text-slate-400 w-16 shrink-0">Option {i + 1}</span>
              <input type="text" value={optionNames[i] || ""} placeholder={`Option ${i + 1}`}
                onChange={e => {
                  const next = [...optionNames];
                  next[i] = e.target.value;
                  setOptionNames(next);
                }}
                className="flex-1 text-sm text-slate-700 border border-slate-200 rounded-lg px-3 py-1.5 focus:outline-none focus:border-slate-400 focus:ring-1 focus:ring-slate-200" />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [categories, setCategories] = useState(DEFAULT_CATEGORIES);
  const [modelName, setModelName] = useState("Weighted Factor Model");
  const [numOptions, setNumOptions] = useState(8);
  const [optionNames, setOptionNames] = useState([]);
  const [showPreview, setShowPreview] = useState(true);

  const totalWeight = categories.reduce((s, c) => s + c.weight, 0);
  const allBalanced = totalWeight === 100 && categories.every(cat => {
    if (!cat.hasSubcriteria) return true;
    const t = cat.criteria.reduce((s, c) => s + c.weight, 0);
    return t === 100 || cat.criteria.length === 0;
  });
  const hasContent = categories.some(c => c.name.trim());

  const updateCategory = useCallback((catId, updated) => setCategories(cats => cats.map(c => c.id === catId ? updated : c)), []);
  const removeCategory = useCallback((catId) => setCategories(cats => cats.filter(c => c.id !== catId)), []);
  const addCategory = useCallback(() => setCategories(cats => [...cats, { ...makeCategory(""), criteria: [makeCriterion("")] }]), []);
  const autoBalanceCategories = useCallback(() => {
    setCategories(cats => {
      const count = cats.length;
      if (count === 0) return cats;
      const base = Math.floor(100 / count);
      const remainder = 100 - base * count;
      return cats.map((c, i) => ({ ...c, weight: base + (i < remainder ? 1 : 0) }));
    });
  }, []);

  const handleExport = async () => {
    await exportToExcel(categories, modelName, numOptions, optionNames);
  };

  return (
    <div style={{ minHeight: "100vh" }}>
      <div style={{ maxWidth: "900px", margin: "0 auto", padding: "32px 16px" }}>
        <div className="mb-8">
          <input type="text" value={modelName} onChange={e => setModelName(e.target.value)}
            className="text-2xl font-bold text-slate-800 bg-transparent border-0 border-b-2 border-transparent hover:border-slate-300 focus:border-slate-500 focus:outline-none px-0 py-1 w-full transition-colors" />
          <p className="text-sm text-slate-400 mt-1">Define your categories, criteria, and weightings. Export to Excel when ready.</p>
        </div>

        <div className="flex flex-wrap items-center gap-4 mb-6">
          <div className="flex items-center gap-2">
            <label className="text-xs text-slate-500">Number of options:</label>
            <input type="number" min="2" max="30" value={numOptions}
              onChange={e => setNumOptions(Math.max(2, Math.min(30, parseInt(e.target.value) || 8)))}
              className="w-16 text-sm text-center border border-slate-200 rounded-lg px-2 py-1.5 bg-white focus:outline-none focus:border-slate-400" />
          </div>
          <div className="flex-1" />
          <button onClick={() => setShowPreview(!showPreview)} className="text-xs text-slate-500 hover:text-slate-700 underline decoration-dotted">
            {showPreview ? "Hide" : "Show"} effective weights
          </button>
          <button onClick={handleExport} disabled={!allBalanced || !hasContent}
            className={`px-5 py-2.5 rounded-xl text-sm font-semibold transition-all shadow-sm ${allBalanced && hasContent ? "bg-slate-800 text-white hover:bg-slate-700 hover:shadow-md" : "bg-slate-200 text-slate-400 cursor-not-allowed"}`}>
            Export to Excel
          </button>
        </div>

        <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg text-xs text-blue-700">
          The exported .xlsx file works in both Excel and Google Sheets. To use in Google Sheets: upload to Drive, then open with Google Sheets.
        </div>

        <div className="mb-6"><WeightSummary categories={categories} /></div>

        <div className="space-y-4 mb-6">
          {categories.map((cat, i) => (
            <CategoryCard key={cat.id} category={cat} index={i}
              onChange={updated => updateCategory(cat.id, updated)}
              onRemove={() => removeCategory(cat.id)} />
          ))}
        </div>

        <div className="flex items-center gap-4 mb-6">
          <button onClick={addCategory}
            className="flex items-center gap-2 px-4 py-2.5 rounded-xl border-2 border-dashed border-slate-300 text-sm font-medium text-slate-500 hover:border-slate-400 hover:text-slate-700 transition-colors">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 5v14M5 12h14"/></svg>
            Add category
          </button>
          <button onClick={autoBalanceCategories} className="text-xs text-slate-500 hover:text-slate-700 underline decoration-dotted">
            Auto-balance categories equally
          </button>
        </div>

        <div className="mb-6">
          <OptionNamer numOptions={numOptions} optionNames={optionNames} setOptionNames={setOptionNames} />
        </div>

        {showPreview && <div className="mb-6"><EffectiveWeightPreview categories={categories} /></div>}

        {!allBalanced && hasContent && (
          <div className="mt-6 p-4 bg-amber-50 border border-amber-200 rounded-xl text-sm text-amber-800">
            <p className="font-semibold mb-1">Fix weightings before exporting:</p>
            <ul className="list-disc list-inside space-y-0.5 text-xs">
              {totalWeight !== 100 && <li>Category weights sum to {totalWeight}% (need 100%)</li>}
              {categories.filter(c => {
                if (!c.hasSubcriteria) return false;
                const t = c.criteria.reduce((s, cr) => s + cr.weight, 0);
                return c.criteria.length > 0 && t !== 100;
              }).map(c => (
                <li key={c.id}>"{c.name || "Untitled"}" criteria sum to {c.criteria.reduce((s, cr) => s + cr.weight, 0)}% (need 100%)</li>
              ))}
            </ul>
          </div>
        )}

        <div className="mt-12 text-center text-xs text-slate-400">Weighted Factor Model Builder</div>
      </div>
    </div>
  );
}
