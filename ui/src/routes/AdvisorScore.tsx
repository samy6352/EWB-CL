import { useQuery } from "@tanstack/react-query";
import { useSearch } from "@tanstack/react-router";
import { useState } from "react";
import { AdvisorScorePanel } from "@/components/AdvisorScore";
import { PageHeader } from "@/components/blocks";
import { OptimizeNotices, OptimizeQueryState, OptimizeSelect, OptimizeTable, useOptimizeScope } from "@/components/OptimizeShared";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { ADVISOR_PILLARS, advisorScore, advisorScoreDate, advisorScoreQuery, type AdvisorCategory, type AdvisorScoreReport } from "@/lib/advisor-score";
import { nativeMoney, optimizeQuery, type AdvisorReport, type AdvisorRow } from "@/lib/optimize-queries";
import { usePrefs } from "@/lib/prefs";

export function AdvisorScoreTable({ data }: { data: AdvisorScoreReport }) {
  return <Card className="overflow-hidden">
    <div className="border-b border-border px-3 py-2">
      <h3 className="text-sm font-semibold">Per-subscription scores</h3>
      <p className="mt-1 text-[11px] text-muted">Percentages from Azure, 0–100. Each cell shows its source timestamp (UTC), not the browser fetch time.</p>
    </div>
    <div className="overflow-x-auto">
      <table className="w-full text-left text-xs">
        <caption className="sr-only">Azure Advisor scores and source timestamps by subscription and pillar</caption>
        <thead className="bg-surface-2 text-muted"><tr>
          <th scope="col" className="px-3 py-2 font-medium">Subscription / coverage</th>
          {ADVISOR_PILLARS.map(p => <th scope="col" className="px-3 py-2 font-medium" key={p.category}>{p.label}</th>)}
        </tr></thead>
        <tbody className="divide-y divide-border">{data.subscriptions.map(subscription => {
          const available = subscription.status === "error" ? 0 : ADVISOR_PILLARS.filter(({ category }) =>
            advisorScore(subscription.pillars.find(p => p.category === category)?.score) !== "Unavailable").length;
          const partial = subscription.status !== "ok" || available < ADVISOR_PILLARS.length;
          return <tr key={subscription.subscription_id}>
          <th scope="row" className="max-w-64 px-3 py-2 align-top font-normal">
            <span className="block break-words font-medium">{subscription.subscription_id}</span>
            <span className={partial ? "text-warn" : "text-muted"}>{subscription.status === "error" ? "Error · scores unavailable" : partial ? "Partial coverage" : "Available"} · {available}/5 pillars</span>
            {subscription.error && <span className="mt-1 block break-words text-warn">{subscription.error}</span>}
          </th>
          {ADVISOR_PILLARS.map(({ category }) => {
            const pillar = subscription.pillars.find(p => p.category === category);
            const score = subscription.status === "error" ? "Unavailable" : advisorScore(pillar?.score);
            return <td className="px-3 py-2 align-top" key={category}>
              <span className="num block font-medium">{score}</span>
              <span className="mt-1 block whitespace-nowrap text-[10px] text-muted">{advisorScoreDate(pillar?.date)}</span>
            </td>;
          })}
        </tr>;
        })}</tbody>
      </table>
    </div>
    {!data.subscriptions.length && <p className="px-3 py-3 text-xs text-muted">No subscription scores returned. This is unavailable evidence, not a perfect or zero score.</p>}
  </Card>;
}

const IMPACT_RANK: Record<string, number> = { High: 0, Medium: 1, Low: 2 };

/** One fetch of every pillar; the category and impact filters narrow it in the browser. */
export function AdvisorRecommendations() {
  const scope = useOptimizeScope();
  const { numbers } = usePrefs();
  const [category, setCategory] = useState<"All" | AdvisorCategory>("All");
  const [impact, setImpact] = useState("");
  const [selected, setSelected] = useState<AdvisorRow>();
  const query = useQuery(optimizeQuery<AdvisorReport>("advisor", scope, { category: "All" }));
  const d = query.data;
  const all = d?.recommendations ?? [];
  const counts = new Map<string, number>();
  for (const r of all) counts.set(r.category ?? "", (counts.get(r.category ?? "") ?? 0) + 1);
  const inCategory = all.filter(r => category === "All" || r.category === category);
  const rows = inCategory
    .filter(r => !impact || r.impact === impact)
    .sort((a, b) => (IMPACT_RANK[a.impact ?? ""] ?? 3) - (IMPACT_RANK[b.impact ?? ""] ?? 3) || (b.annual_savings ?? 0) - (a.annual_savings ?? 0));
  const byImpact = (level: string) => inCategory.filter(r => r.impact === level).length;
  const sourceErrors = Array.isArray(d?.errors)
    ? (d.errors as unknown as Array<{ subscription_id: string; error: string }>).map(e => `${e.subscription_id}: ${e.error}`)
    : Object.entries(d?.errors ?? {}).map(([sub, error]) => `${sub}: ${error}`);
  const label = category === "All" ? "All pillars" : ADVISOR_PILLARS.find(p => p.category === category)?.label ?? category;
  return <section aria-label="Advisor recommendations" className="mt-6">
    <div className="mb-3 flex flex-wrap items-end justify-between gap-3">
      <div>
        <h3 className="text-sm font-semibold">Advisor recommendations</h3>
        <p className="mt-1 text-xs text-muted">The recommendations behind each pillar score. Choose a pillar to focus on it.</p>
      </div>
      <OptimizeSelect label="Pillar" value={category} onChange={v => { setCategory(v as "All" | AdvisorCategory); setSelected(undefined); }}
        options={[{ value: "All", label: `All pillars (${all.length})` }, ...ADVISOR_PILLARS.map(p => ({ value: p.category, label: `${p.label} (${counts.get(p.category) ?? 0})` }))]} />
    </div>
    <OptimizeQueryState query={query}>
      <OptimizeNotices notices={[d?.error, d?.note, ...sourceErrors,
        d?.partial ? "Advisor assessment incomplete; unavailable subscriptions are not empty recommendations." : null]} />
      <div className="mb-3 flex flex-wrap items-end gap-3">
        <OptimizeSelect label="Impact" value={impact} onChange={setImpact}
          options={[{ value: "", label: "Any impact" }, ...["High", "Medium", "Low"].map(v => ({ value: v, label: `${v} (${byImpact(v)})` }))]} />
        <span role="status" className="text-xs text-muted">{rows.length} of {all.length} recommendations · {label}</span>
      </div>
      <OptimizeTable<AdvisorRow> title={`Advisor recommendations · ${label}`} rows={rows}
        rowKey={(r, i) => r.id ?? `${r.resource}-${i}`} onSelect={setSelected}
        empty="No Advisor recommendations for this selection. Empty is normal for a healthy or new subscription."
        columns={[
          { id: "problem", label: "Recommendation", value: r => r.problem || r.solution },
          { id: "category", label: "Pillar", value: r => ADVISOR_PILLARS.find(p => p.category === r.category)?.label ?? r.category },
          { id: "impact", label: "Impact", value: r => r.impact },
          { id: "resource", label: "Resource", value: r => r.resource },
          { id: "group", label: "Resource group", value: r => r.resourceGroup },
          { id: "saving", label: "Annual saving", value: r => r.annual_savings,
            cell: r => r.annual_savings == null ? "—" : nativeMoney(r.annual_savings, r.currency || d?.currency, numbers), numeric: true },
        ]}
        note="Advisor's own findings. Savings are provider estimates for Cost recommendations, not realised savings." />
      {selected && <Card className="mt-4 p-5">
        <h3 className="font-medium">{selected.problem}</h3>
        <p className="mt-2 text-sm">{selected.solution}</p>
        <p className="mt-2 break-words text-sm text-muted">{[selected.resource, selected.resourceGroup, selected.subscriptionId, selected.impact && `${selected.impact} impact`].filter(Boolean).join(" · ")}</p>
        <Button className="mt-3" variant="ghost" onClick={() => setSelected(undefined)}>Close details</Button>
      </Card>}
    </OptimizeQueryState>
  </section>;
}

export function AdvisorScorePage() {
  const search = useSearch({ from: "__root__" });
  const query = useQuery(advisorScoreQuery(search.scope));
  return <div className="min-w-0">
    <PageHeader title="Advisor Score" subtitle="Review five pillars and subscription-level evidence in the selected scope." />
    <AdvisorScorePanel details />
    {!query.isError && query.data && <AdvisorScoreTable data={query.data} />}
    <AdvisorRecommendations />
  </div>;
}
