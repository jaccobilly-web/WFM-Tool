import { useState, useCallback } from "react";
import { exportToExcel } from "./export";

let _id = 0;
const uid = () => `id_${++_id}_${Date.now()}`;

const makeCriterion = (name = "") => ({ id: uid(), name, weight: 0, criteria: [], hasSubcriteria: true });
const makeSubcriterion = (name = "") => ({ id: uid(), name, weight: 0, description: "" });

function WeightBar({ value, total, color = "emerald" }) {
  const pct = total > 0 ? Math.min((value / total) * 100, 100) : 0;
  const colors = { emerald: "bg-emerald-500", amber: "bg-amber-500", red: "bg-red-500" };
  return (
    <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
      <div className={`h-full rounded-full transition-all duration-300 ${colors[color]}`} style={{ width: `${pct}%` }} />
    </div>
  );
}

function StepDots({ current, total }) {
  return (
    <div className="flex items-center gap-2 justify-center mb-8">
      {Array.from({ length: total }, (_, i) => (
        <div key={i} className={`h-2 rounded-full transition-all duration-300 ${i === current ? "w-8 bg-slate-700" : i < current ? "w-2 bg-emerald-400" : "w-2 bg-slate-200"}`} />
      ))}
    </div>
  );
}

function NavButtons({ onBack, onNext, nextLabel = "Continue", nextDisabled = false }) {
  return (
    <div className="flex items-center justify-between mt-8">
      {onBack ? (
        <button onClick={onBack} className="text-sm text-slate-500 hover:text-slate-700 flex items-center gap-1">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>
          Back
        </button>
      ) : <div />}
      <button onClick={onNext} disabled={nextDisabled}
        className={`px-6 py-2.5 rounded-xl text-sm font-semibold transition-all ${nextDisabled ? "bg-slate-200 text-slate-400 cursor-not-allowed" : "bg-slate-800 text-white hover:bg-slate-700 shadow-sm hover:shadow-md"}`}>
        {nextLabel}
      </button>
    </div>
  );
}

function WelcomeStep({ onNext }) {
  return (
    <div className="max-w-xl mx-auto text-center">
      <div className="mb-6">
        <div className="w-16 h-16 bg-slate-800 rounded-2xl flex items-center justify-center mx-auto mb-4">
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.5"><path d="M3 6h18M3 12h18M3 18h18"/><circle cx="7" cy="6" r="1.5" fill="white"/><circle cx="14" cy="12" r="1.5" fill="white"/><circle cx="10" cy="18" r="1.5" fill="white"/></svg>
        </div>
        <h1 className="text-2xl font-bold text-slate-800 mb-3">Weighted Factor Model Builder</h1>
        <p className="text-slate-500 text-sm leading-relaxed max-w-md mx-auto">
          A structured way to compare options and make better decisions. Define your criteria here, then score your options in the generated spreadsheet.
        </p>
      </div>
      <div className="text-left bg-white rounded-xl border border-slate-200 p-5 mb-6">
        <div className="space-y-3">
          <p className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">Build your model</p>
          {[
            { n: "1", title: "Name your decision", desc: "What are you trying to choose between?" },
            { n: "2", title: "List your options", desc: "The things you're comparing." },
            { n: "3", title: "Define your criteria and weights", desc: "What factors matter, and how much?" },
          ].map((item, i) => (
            <div key={i} className="flex gap-3">
              <span className="w-5 h-5 rounded-full bg-slate-800 text-white text-[10px] font-bold flex items-center justify-center shrink-0 mt-0.5">{item.n}</span>
              <div>
                <p className="text-sm font-medium text-slate-700">{item.title}</p>
                <p className="text-xs text-slate-400">{item.desc}</p>
              </div>
            </div>
          ))}
          <div className="flex justify-center py-2">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2"><path d="M12 5v14M5 12l7 7 7-7"/></svg>
          </div>
          <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-4">
            <div className="flex gap-3">
              <div className="w-10 h-10 bg-emerald-600 rounded-lg flex items-center justify-center shrink-0">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
              </div>
              <div>
                <p className="text-sm font-semibold text-emerald-800">Export to Google Sheets or Excel</p>
                <p className="text-xs text-emerald-600 leading-relaxed">The app generates a fully formatted spreadsheet with your criteria, weights, and formulas built in. Score each option in the blue cells and results are calculated automatically using z-score normalisation.</p>
              </div>
            </div>
          </div>
        </div>
      </div>
      <button onClick={onNext} className="px-8 py-3 bg-slate-800 text-white rounded-xl text-sm font-semibold hover:bg-slate-700 shadow-sm hover:shadow-md transition-all">
        Get started
      </button>
    </div>
  );
}

function NameStep({ modelName, setModelName, modelDescription, setModelDescription, onBack, onNext }) {
  return (
    <div className="max-w-lg mx-auto">
      <h2 className="text-xl font-bold text-slate-800 mb-1">Name your model</h2>
      <p className="text-sm text-slate-400 mb-6">Give it a name and describe what decision you're trying to make.</p>
      <div className="space-y-4">
        <div>
          <label className="text-xs font-medium text-slate-500 mb-1.5 block">Model name</label>
          <input type="text" value={modelName} onChange={e => setModelName(e.target.value)}
            placeholder="e.g. Career Options Analysis"
            className="w-full text-lg font-semibold text-slate-800 bg-white border border-slate-200 rounded-lg px-4 py-3 focus:outline-none focus:border-slate-400 focus:ring-1 focus:ring-slate-200" />
        </div>
        <div>
          <label className="text-xs font-medium text-slate-500 mb-1.5 block">What decision is this helping you make? <span className="text-slate-300 font-normal">(optional)</span></label>
          <textarea value={modelDescription} onChange={e => setModelDescription(e.target.value)}
            placeholder="e.g. Choosing the best next step for my career, comparing across impact, personal fit, and logistics"
            rows={3}
            className="w-full text-sm text-slate-600 bg-white border border-slate-200 rounded-lg px-4 py-3 focus:outline-none focus:border-slate-400 focus:ring-1 focus:ring-slate-200 resize-none" />
        </div>
      </div>
      <NavButtons onBack={onBack} onNext={onNext} nextDisabled={!modelName.trim()} />
    </div>
  );
}

function OptionsStep({ numOptions, setNumOptions, optionNames, setOptionNames, onBack, onNext }) {
  const [knowOptions, setKnowOptions] = useState(optionNames.some(n => n && n.trim()) ? "yes" : null);
  const filledCount = optionNames.filter(n => n && n.trim()).length;
  const handleChange = (i, value) => {
    const next = [...optionNames];
    next[i] = value;
    setOptionNames(next);
    if (i >= numOptions - 1 && value.trim()) setNumOptions(Math.max(numOptions, i + 2));
  };
  const handleRemove = (i) => {
    const next = [...optionNames]; next.splice(i, 1); setOptionNames(next);
    setNumOptions(Math.max(2, numOptions - 1));
  };
  const displayCount = Math.max(filledCount + 1, 2);
  return (
    <div className="max-w-lg mx-auto">
      <h2 className="text-xl font-bold text-slate-800 mb-1">What are you comparing?</h2>
      <p className="text-sm text-slate-400 mb-6">You can always add or remove rows in the spreadsheet later.</p>

      {knowOptions === null && (
        <div className="space-y-3">
          <button onClick={() => setKnowOptions("yes")}
            className="w-full text-left p-4 bg-white rounded-xl border border-slate-200 hover:border-slate-300 hover:shadow-sm transition-all">
            <p className="text-sm font-medium text-slate-700">I know my options</p>
            <p className="text-xs text-slate-400">I can name them now</p>
          </button>
          <button onClick={() => setKnowOptions("no")}
            className="w-full text-left p-4 bg-white rounded-xl border border-slate-200 hover:border-slate-300 hover:shadow-sm transition-all">
            <p className="text-sm font-medium text-slate-700">I'm not sure yet</p>
            <p className="text-xs text-slate-400">I'll just set a rough number and fill in names later</p>
          </button>
        </div>
      )}

      {knowOptions === "yes" && (
        <div>
          <label className="text-xs font-medium text-slate-500 mb-3 block">Name your options. A new row appears as you type.</label>
          <div className="space-y-2">
            {Array.from({ length: displayCount }, (_, i) => {
              const value = optionNames[i] || "";
              const isFilled = value.trim().length > 0;
              const isGhost = i === displayCount - 1 && !isFilled;
              return (
                <div key={i} className="flex items-center gap-2">
                  <span className={`text-xs w-6 text-right shrink-0 ${isGhost ? "text-slate-200" : "text-slate-400"}`}>{i + 1}.</span>
                  <input type="text" value={value}
                    placeholder={isGhost ? "Add another option..." : `Option ${i + 1}`}
                    onChange={e => handleChange(i, e.target.value)}
                    className={`flex-1 text-sm rounded-lg px-3 py-2 focus:outline-none focus:border-slate-400 focus:ring-1 focus:ring-slate-200 transition-colors ${isGhost ? "bg-slate-50 border border-dashed border-slate-200 text-slate-400 placeholder-slate-300" : "bg-white border border-slate-200 text-slate-700 font-medium"}`} />
                  {isFilled && displayCount > 2 && (
                    <button onClick={() => handleRemove(i)} className="text-slate-300 hover:text-red-400 p-1">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M18 6L6 18M6 6l12 12"/></svg>
                    </button>
                  )}
                </div>
              );
            })}
          </div>
          <p className="text-xs text-slate-400 mt-3 italic">{filledCount} option{filledCount === 1 ? "" : "s"} named.</p>
        </div>
      )}

      {knowOptions === "no" && (
        <div>
          <div className="flex items-center gap-4 mb-3">
            <label className="text-xs font-medium text-slate-500">Roughly how many options?</label>
            <input type="number" min="2" max="30" value={numOptions}
              onChange={e => setNumOptions(Math.max(2, Math.min(30, parseInt(e.target.value) || 4)))}
              className="w-20 text-sm text-center border border-slate-200 rounded-lg px-2 py-2 bg-white focus:outline-none focus:border-slate-400" />
          </div>
          <p className="text-xs text-slate-400 italic">Tip: choose more than you think you need. It's easier to delete unused rows than to add new ones.</p>
        </div>
      )}

      {knowOptions !== null && (
        <div className="mt-4">
          <button onClick={() => setKnowOptions(null)} className="text-xs text-slate-400 hover:text-slate-600 underline decoration-dotted">
            Change approach
          </button>
        </div>
      )}

      <NavButtons onBack={onBack} onNext={onNext} />
    </div>
  );
}

function CriteriaWeightVisual({ criteria }) {
  const total = criteria.reduce((s, c) => s + c.weight, 0);
  const barColor = total === 100 ? "#10b981" : total > 100 ? "#ef4444" : "#f59e0b";
  if (criteria.length === 0 || total === 0) return null;
  return (
    <div className="flex gap-0.5 mt-2 mb-1">
      {criteria.map(crit => {
        const pct = crit.weight;
        const lbl = crit.name || "Untitled";
        return pct > 0 ? (
          <div key={crit.id} className="rounded-sm transition-all duration-300 overflow-hidden flex items-center justify-center"
            style={{ width: `${pct}%`, backgroundColor: barColor, opacity: 0.5 + (pct / 200), minHeight: "24px" }}>
            {pct >= 12 && <span className="text-[9px] text-white font-semibold truncate px-1">{lbl} ({pct}%)</span>}
            {pct >= 6 && pct < 12 && <span className="text-[8px] text-white font-bold">{pct}%</span>}
          </div>
        ) : null;
      })}
    </div>
  );
}

function SubcriterionRow({ criterion, onChange, onRemove }) {
  return (
    <div className="flex items-start gap-3 py-2.5 group">
      <div className="flex-1 min-w-0">
        <input type="text" value={criterion.name} placeholder="Sub-criterion name"
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

function CriterionCard({ criterion, onChange, onRemove, index }) {
  const subTotal = criterion.criteria.reduce((s, c) => s + c.weight, 0);
  const subBalanced = subTotal === 100;
  const subColor = subBalanced ? "emerald" : subTotal > 100 ? "red" : "amber";
  const toggleSub = () => {
    if (criterion.hasSubcriteria) onChange({ ...criterion, hasSubcriteria: false, criteria: [] });
    else onChange({ ...criterion, hasSubcriteria: true, criteria: [makeSubcriterion("")] });
  };
  const updateSub = (id, updated) => onChange({ ...criterion, criteria: criterion.criteria.map(c => c.id === id ? updated : c) });
  const removeSub = (id) => onChange({ ...criterion, criteria: criterion.criteria.filter(c => c.id !== id) });
  const addSub = () => onChange({ ...criterion, criteria: [...criterion.criteria, makeSubcriterion()] });
  const autoBalance = () => {
    const count = criterion.criteria.length; if (count === 0) return;
    const base = Math.floor(100 / count); const remainder = 100 - base * count;
    onChange({ ...criterion, criteria: criterion.criteria.map((c, i) => ({ ...c, weight: base + (i < remainder ? 1 : 0) })) });
  };
  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
      <div className="px-5 py-4 bg-slate-50 border-b border-slate-200">
        <div className="flex items-center gap-3">
          <span className="text-xs font-bold text-slate-400 bg-slate-200 w-6 h-6 rounded-full flex items-center justify-center">{index + 1}</span>
          <input type="text" value={criterion.name} placeholder="e.g. Cost, Quality, Risk..."
            onChange={e => onChange({ ...criterion, name: e.target.value })}
            className="flex-1 text-base font-semibold text-slate-800 bg-transparent border-0 border-b-2 border-transparent hover:border-slate-300 focus:border-slate-500 focus:outline-none px-0 py-0.5 transition-colors" />
          <div className="flex items-center gap-2 shrink-0">
            <span className="text-xs text-slate-500">Weight:</span>
            <input type="number" min="0" max="100" value={criterion.weight}
              onChange={e => onChange({ ...criterion, weight: Math.max(0, Math.min(100, parseInt(e.target.value) || 0)) })}
              className="w-16 text-sm text-right font-bold text-slate-700 border border-slate-200 rounded-lg px-2 py-1.5 focus:outline-none focus:border-slate-400 focus:ring-1 focus:ring-slate-200" />
            <span className="text-xs text-slate-400">%</span>
          </div>
          <button onClick={onRemove} className="text-slate-300 hover:text-red-500 transition-colors p-1 ml-1">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M18 6L6 18M6 6l12 12"/></svg>
          </button>
        </div>
        <div className="mt-2">
          <label className="flex items-center gap-2 cursor-pointer">
            <input type="checkbox" checked={criterion.hasSubcriteria} onChange={toggleSub} className="rounded border-slate-300 text-slate-600 focus:ring-slate-400" />
            <span className="text-xs text-slate-500">Break down into sub-criteria</span>
          </label>
        </div>
      </div>
      {criterion.hasSubcriteria && (
        <div className="px-5 py-3">
          <div className="flex items-center justify-between mb-1">
            <span className="text-xs text-slate-500">Sub-criteria weights (must total 100%)</span>
            <div className="flex items-center gap-3">
              <button onClick={autoBalance} className="text-xs text-slate-500 hover:text-slate-700 underline decoration-dotted">Auto-balance</button>
              <span className={`text-xs font-bold ${subBalanced ? "text-emerald-600" : subTotal > 100 ? "text-red-600" : "text-amber-600"}`}>{subTotal}%</span>
            </div>
          </div>
          <WeightBar value={subTotal} total={100} color={subColor} />
          <CriteriaWeightVisual criteria={criterion.criteria} />
          <div className="mt-2 divide-y divide-slate-100">
            {criterion.criteria.map(sub => (
              <SubcriterionRow key={sub.id} criterion={sub} onChange={updated => updateSub(sub.id, updated)} onRemove={() => removeSub(sub.id)} />
            ))}
          </div>
          <button onClick={addSub} className="mt-3 text-xs text-slate-500 hover:text-slate-700 flex items-center gap-1 py-1.5">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 5v14M5 12h14"/></svg>
            Add sub-criterion
          </button>
        </div>
      )}
      {!criterion.hasSubcriteria && (
        <div className="px-5 py-3"><p className="text-xs text-slate-400 italic">No sub-criteria. This will appear as a single scored column in the spreadsheet.</p></div>
      )}
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
        <span className="text-sm font-semibold text-slate-700">Criteria Weights</span>
        <div className="flex items-center gap-2">
          <span className={`text-2xl font-bold ${balanced ? "text-emerald-700" : total > 100 ? "text-red-700" : "text-amber-700"}`}>{total}%</span>
          <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${balanced ? "bg-emerald-200 text-emerald-800" : total > 100 ? "bg-red-200 text-red-800" : "bg-amber-200 text-amber-800"}`}>{label}</span>
        </div>
      </div>
      <WeightBar value={total} total={100} color={color} />
      <div className="flex gap-1 mt-3">
        {categories.map(cat => {
          const pct = cat.weight; const lbl = cat.name || "Untitled";
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

function EffectiveWeightPreview({ categories }) {
  const totalCatWeight = categories.reduce((s, c) => s + c.weight, 0);
  const allCriteria = [];
  categories.forEach(cat => {
    if (cat.hasSubcriteria && cat.criteria.length > 0) {
      const critTotal = cat.criteria.reduce((s, c) => s + c.weight, 0);
      cat.criteria.forEach(crit => {
        const ew = critTotal > 0 ? (cat.weight / (totalCatWeight || 100)) * (crit.weight / critTotal) * 100 : 0;
        allCriteria.push({ criterion: cat.name || "Untitled", name: crit.name || "Untitled", effectiveWeight: Math.round(ew * 10) / 10 });
      });
    } else {
      const ew = totalCatWeight > 0 ? (cat.weight / totalCatWeight) * 100 : 0;
      allCriteria.push({ criterion: cat.name || "Untitled", name: cat.name || "Untitled", effectiveWeight: Math.round(ew * 10) / 10 });
    }
  });
  allCriteria.sort((a, b) => b.effectiveWeight - a.effectiveWeight);
  const maxWeight = Math.max(...allCriteria.map(c => c.effectiveWeight), 1);
  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-5">
      <h3 className="text-sm font-semibold text-slate-700 mb-1">Effective Weights</h3>
      <p className="text-xs text-slate-400 mb-4">The actual contribution of each factor to the final score.</p>
      <div className="space-y-2">
        {allCriteria.map((c, i) => (
          <div key={i} className="flex items-center gap-3">
            <span className="text-xs text-slate-500 w-28 shrink-0 truncate" title={c.criterion}>{c.criterion}</span>
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
      {allCriteria.length === 0 && <p className="text-xs text-slate-400 italic">Add criteria to see effective weights</p>}
    </div>
  );
}

function BuilderStep({ categories, setCategories, onBack, onExport, allBalanced, hasContent }) {
  const [showPreview, setShowPreview] = useState(true);
  const updateCriterion = useCallback((id, updated) => setCategories(cats => cats.map(c => c.id === id ? updated : c)), []);
  const removeCriterion = useCallback((id) => setCategories(cats => cats.filter(c => c.id !== id)), []);
  const addCriterion = useCallback(() => setCategories(cats => [...cats, { ...makeCriterion(""), criteria: [makeSubcriterion("")] }]), []);
  const autoBalanceAll = useCallback(() => {
    setCategories(cats => {
      const count = cats.length; if (count === 0) return cats;
      const base = Math.floor(100 / count); const remainder = 100 - base * count;
      return cats.map((c, i) => ({ ...c, weight: base + (i < remainder ? 1 : 0) }));
    });
  }, []);
  const totalWeight = categories.reduce((s, c) => s + c.weight, 0);
  return (
    <div>
      <div className="flex items-start justify-between mb-6">
        <div>
          <h2 className="text-xl font-bold text-slate-800 mb-1">Define your criteria</h2>
          <p className="text-sm text-slate-400">What factors matter for this decision? Set weights to reflect their relative importance.</p>
        </div>
      </div>
      <div className="mb-4"><WeightSummary categories={categories} /></div>
      <div className="space-y-4 mb-4">
        {categories.map((cat, i) => (
          <CriterionCard key={cat.id} criterion={cat} index={i}
            onChange={updated => updateCriterion(cat.id, updated)} onRemove={() => removeCriterion(cat.id)} />
        ))}
      </div>
      <div className="flex items-center gap-4 mb-8">
        <button onClick={addCriterion}
          className="flex items-center gap-2 px-4 py-2.5 rounded-xl border-2 border-dashed border-slate-300 text-sm font-medium text-slate-500 hover:border-slate-400 hover:text-slate-700 transition-colors">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M12 5v14M5 12h14"/></svg>
          Add criterion
        </button>
        <button onClick={autoBalanceAll} className="text-xs text-slate-500 hover:text-slate-700 underline decoration-dotted">
          Auto-balance equally
        </button>
      </div>
      {showPreview && <div className="mb-4"><EffectiveWeightPreview categories={categories} /></div>}
      <div className="flex justify-center mb-6">
        <button onClick={() => setShowPreview(!showPreview)} className="text-xs text-slate-500 hover:text-slate-700 underline decoration-dotted">
          {showPreview ? "Hide" : "Show"} effective weights
        </button>
      </div>
      {!allBalanced && hasContent && (
        <div className="mb-6 p-4 bg-amber-50 border border-amber-200 rounded-xl text-sm text-amber-800">
          <p className="font-semibold mb-1">Fix weightings before exporting:</p>
          <ul className="list-disc list-inside space-y-0.5 text-xs">
            {totalWeight !== 100 && <li>Criteria weights sum to {totalWeight}% (need 100%)</li>}
            {categories.filter(c => { if (!c.hasSubcriteria) return false; const t = c.criteria.reduce((s, cr) => s + cr.weight, 0); return c.criteria.length > 0 && t !== 100; }).map(c => (
              <li key={c.id}>&ldquo;{c.name || "Untitled"}&rdquo; sub-criteria sum to {c.criteria.reduce((s, cr) => s + cr.weight, 0)}% (need 100%)</li>
            ))}
          </ul>
        </div>
      )}
      <div className="flex items-center justify-between">
        <button onClick={onBack} className="text-sm text-slate-500 hover:text-slate-700 flex items-center gap-1">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>
          Back
        </button>
        <div className="flex items-center gap-3">
          {(!allBalanced || !hasContent) && (
            <span className="text-xs text-slate-400 italic">{!hasContent ? "Add at least one named criterion" : "Balance all weights to enable export"}</span>
          )}
          <button onClick={onExport} disabled={!allBalanced || !hasContent}
            className={`px-6 py-3 rounded-xl text-sm font-semibold transition-all shadow-sm ${allBalanced && hasContent ? "bg-emerald-600 text-white hover:bg-emerald-700 hover:shadow-md" : "bg-slate-200 text-slate-400 cursor-not-allowed"}`}>
            <span className="flex items-center gap-2">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Export to Excel
            </span>
          </button>
        </div>
      </div>
      <p className="text-xs text-slate-400 text-center mt-4">The exported .xlsx works in both Excel and Google Sheets. For Sheets: upload to Drive, then open with Google Sheets.</p>
    </div>
  );
}

const DEFAULT_CRITERIA = [
  { ...makeCriterion(""), criteria: [makeSubcriterion("")] },
  { ...makeCriterion(""), criteria: [makeSubcriterion("")] },
];

export default function App() {
  const [step, setStep] = useState(0);
  const [modelName, setModelName] = useState("");
  const [modelDescription, setModelDescription] = useState("");
  const [numOptions, setNumOptions] = useState(4);
  const [optionNames, setOptionNames] = useState([]);
  const [categories, setCategories] = useState(DEFAULT_CRITERIA);
  const totalWeight = categories.reduce((s, c) => s + c.weight, 0);
  const allBalanced = totalWeight === 100 && categories.every(cat => {
    if (!cat.hasSubcriteria) return true;
    const t = cat.criteria.reduce((s, c) => s + c.weight, 0);
    return t === 100 || cat.criteria.length === 0;
  });
  const hasContent = categories.some(c => c.name.trim());
  const filledOptionCount = optionNames.filter(n => n && n.trim()).length;
  const effectiveNumOptions = filledOptionCount > 0 ? filledOptionCount : numOptions;
  const handleExport = async () => {
    await exportToExcel(categories, modelName || "Weighted Factor Model", modelDescription, effectiveNumOptions, optionNames);
  };
  return (
    <div style={{ minHeight: "100vh" }}>
      <div style={{ maxWidth: step <= 2 ? "640px" : "900px", margin: "0 auto", padding: "40px 16px", transition: "max-width 0.3s" }}>
        {step > 0 && step <= 3 && <StepDots current={step - 1} total={3} />}
        {step === 0 && <WelcomeStep onNext={() => setStep(1)} />}
        {step === 1 && <NameStep modelName={modelName} setModelName={setModelName} modelDescription={modelDescription} setModelDescription={setModelDescription} onBack={() => setStep(0)} onNext={() => setStep(2)} />}
        {step === 2 && <OptionsStep numOptions={numOptions} setNumOptions={setNumOptions} optionNames={optionNames} setOptionNames={setOptionNames} onBack={() => setStep(1)} onNext={() => setStep(3)} />}
        {step === 3 && <BuilderStep categories={categories} setCategories={setCategories} onBack={() => setStep(2)} onExport={handleExport} allBalanced={allBalanced} hasContent={hasContent} />}
        <div className="mt-12 text-center text-xs text-slate-300">&copy; Jacco Rubens</div>
      </div>
    </div>
  );
}
