const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

function loadData() {
  const wb = XLSX.readFile(path.join(__dirname, 'data', 'algonox_data.xlsx'));
  const data = {};
  wb.SheetNames.forEach(name => {
    data[name.toLowerCase()] = XLSX.utils.sheet_to_json(wb.Sheets[name]);
  });
  return data;
}

const fmt = (n) => `₹${(n / 100000).toFixed(1)}L`;
const avg = (arr) => arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0;

function queryEngine(question, data) {
  const q = question.toLowerCase();
  const { pipeline, revenue, incidents, team, prospects } = data;
  const wonDeals = pipeline.filter(r => r.Stage === 'Closed Won');
  const lostDeals = pipeline.filter(r => r.Stage === 'Closed Lost');

  // ══════════════════════════════════════════════════════════
  // MEMORA — ORGANIZATIONAL MEMORY QUERIES (checked FIRST)
  // ══════════════════════════════════════════════════════════

  // ── ACE Platform incident history ────────────────────────
  if (q.match(/ace.*(incident|fail|issue|problem|bug|down|histor|occur|before|crash|error)|agentic.*(fail|issue|problem)|workflow.*(fail|stuck|loop|issue)/)) {
    const aceInc = incidents.filter(i => i.System.toLowerCase().includes('ace'));
    const avgRes = avg(aceInc.map(i => i.ResolutionTime_hrs));
    const cats = {};
    aceInc.forEach(i => { cats[i.Category] = (cats[i.Category] || 0) + 1; });
    const topCat = Object.entries(cats).sort((a, b) => b[1] - a[1])[0];
    return {
      type: 'memora',
      title: `ACE Platform — ${aceInc.length} Historical Incidents Retrieved`,
      summary: `MEMORA surfaced **${aceInc.length} past incidents** on ACE Platform spanning **${new Set(aceInc.map(i=>i.Category)).size} categories**. Average resolution time: **${avgRes.toFixed(1)} hrs**. Most frequent failure category: **${topCat[0]}** (${topCat[1]} occurrences).`,
      table: aceInc.map(i => ({ ID: i.IncidentID, Date: i.Date, Category: i.Category, Severity: i.Severity, 'Root Cause': i.RootCause, 'Resolved By': i.ResolvedBy, 'Time (hrs)': i.ResolutionTime_hrs })),
      memory: `Pattern identified: Connector/integration failures (Salesforce OAuth, SAP callbacks) are the #1 recurring category. MEMORA recommends proactive OAuth token refresh monitoring and SAP callback validation before every deployment.`,
      timeline: aceInc.map(i => ({ id: i.IncidentID, date: i.Date, title: i.Category, severity: i.Severity, desc: i.RootCause, resolvedBy: i.ResolvedBy, hours: i.ResolutionTime_hrs })),
      source: 'MEMORA Knowledge Base → ACE Platform Incidents → Pattern Analysis'
    };
  }

  // ── Sweet Hello AI incidents ─────────────────────────────
  if (q.match(/sweet hello.*(incident|fail|issue|histor|problem|before)|voice.*(fail|drop|issue|quality|problem|histor)|call.*(drop|fail|quality|issue|histor)|ivr|nlp.*(issue|fail|accuracy)|intent.*(fail|drop|issue)/)) {
    const shInc = incidents.filter(i => i.System.toLowerCase().includes('sweet hello'));
    return {
      type: 'memora',
      title: `Sweet Hello AI — ${shInc.length} Historical Incidents Retrieved`,
      summary: `MEMORA surfaced **${shInc.length} past incidents** on Sweet Hello AI. Issues span voice infrastructure, NLP models, and compliance. Avg resolution: **${avg(shInc.map(i => i.ResolutionTime_hrs)).toFixed(1)} hrs**. Recurring pattern: voice codec and carrier failover failures.`,
      table: shInc.map(i => ({ ID: i.IncidentID, Date: i.Date, Category: i.Category, Severity: i.Severity, 'Root Cause': i.RootCause, 'Resolved By': i.ResolvedBy, Resolution: i.Resolution })),
      memory: `MEMORA pattern: Voice infrastructure failures (codec mismatch, carrier failover) represent 50% of Sweet Hello AI incidents. Proactive codec standardization and multi-carrier fallback testing is advised before every release.`,
      timeline: shInc.map(i => ({ id: i.IncidentID, date: i.Date, title: i.Category, severity: i.Severity, desc: i.RootCause, resolvedBy: i.ResolvedBy, hours: i.ResolutionTime_hrs })),
      source: 'MEMORA Knowledge Base → Sweet Hello AI Incidents → Root Cause Analysis'
    };
  }

  // ── Expert discovery ──────────────────────────────────────
  if (q.match(/who.*(fix|resolve|handle|expert|contact|know|help|call|reach)|expert.*(for|on|about)|contact.*(for|about|on)|specialist|sme/)) {
    const keyword = q.match(/voice|call|ivr|audio|codec/) ? 'Voice' :
                    q.match(/nlp|intent|model|ai|ml/) ? 'AI/ML' :
                    q.match(/infra|dns|cloud|kafka|server/) ? 'Infra' :
                    q.match(/compliance|trai|dnd|gdpr|regulation/) ? 'Compliance' :
                    q.match(/integration|salesforce|sap|erp|oauth|connector/) ? 'Integrations' :
                    q.match(/api|gateway|backend/) ? 'Backend' : null;

    const expertTeam = keyword === 'Voice' ? 'Infra Team' :
                       keyword === 'AI/ML' ? 'AI/ML Team' :
                       keyword === 'Infra' ? 'Infra Team' :
                       keyword === 'Compliance' ? 'Compliance Team' :
                       keyword === 'Integrations' ? 'Dev Team – Integrations' :
                       keyword === 'Backend' ? 'Dev Team – Backend' : null;

    const relevantInc = expertTeam ? incidents.filter(i => i.ResolvedBy.includes(expertTeam)) : incidents;
    const relevantTeam = team.filter(e => {
      if (!keyword) return true;
      return e.Expert_In.toLowerCase().includes(keyword.toLowerCase()) || e.Team.toLowerCase().includes(keyword.toLowerCase());
    });

    return {
      type: 'memora',
      title: `Expert Discovery${keyword ? ` — ${keyword} Domain` : ' — All Domains'}`,
      summary: `MEMORA identified **${relevantTeam.length} subject-matter expert(s)**${keyword ? ` for ${keyword} issues` : ' across all domains'}. These experts have collectively resolved **${relevantInc.length} incidents** in the knowledge base. Expertise mapped from historical incident resolution records.`,
      table: relevantTeam.map(e => ({ Name: e.Name, Team: e.Team, Role: e.Role, 'Expert In': e.Expert_In, Location: e.Location })),
      memory: `MEMORA expertise mapping: These ${relevantTeam.length} experts have the highest incident resolution rate in the ${keyword || 'relevant'} domain. Contact them directly for fastest resolution. Historical resolution data confirms their specialization.`,
      expertCards: relevantTeam.map(e => ({
        name: e.Name, team: e.Team, role: e.Role, expertIn: e.Expert_In,
        location: e.Location,
        incidentsResolved: incidents.filter(i => i.ResolvedBy && i.ResolvedBy.includes(e.Team)).length
      })),
      source: 'MEMORA → Team Expertise Graph + Incident Resolution History'
    };
  }

  // ── Compliance / TRAI / DND ───────────────────────────────
  if (q.match(/compliance|trai|dnd|gdpr|regulation|privacy|legal/)) {
    const compInc = incidents.filter(i => i.Category.toLowerCase().includes('compliance'));
    const compTeam = team.filter(e => e.Team === 'Compliance');
    return {
      type: 'memora',
      title: 'Compliance — Regulatory Incident Memory',
      summary: `MEMORA found **${compInc.length} compliance-related incident(s)** — all classified Critical severity. Required immediate regulatory action. Compliance team lead: **${compTeam[0]?.Name}**. Resolution involved TRAI DND registry re-sync and automated daily validation enforcement.`,
      table: compInc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy })),
      memory: `MEMORA regulatory insight: TRAI DND registry sync failures are a serious regulatory risk. Automated daily DND validation is now enforced post INC-008. Always verify compliance status before launching new outbound campaigns. Non-compliance can result in license suspension.`,
      timeline: compInc.map(i => ({ id: i.IncidentID, date: i.Date, title: i.Category, severity: i.Severity, desc: i.RootCause, resolvedBy: i.ResolvedBy, hours: i.ResolutionTime_hrs })),
      source: 'MEMORA Knowledge Base → Compliance Category → Regulatory Filter'
    };
  }

  // ── Critical / High severity ──────────────────────────────
  if (q.match(/critical|severity|high.*(incident|issue)|major.*(incident|outage|issue)|p0|p1/)) {
    const critInc = incidents.filter(i => i.Severity === 'Critical' || i.Severity === 'High');
    const critOnly = incidents.filter(i => i.Severity === 'Critical');
    const highOnly = incidents.filter(i => i.Severity === 'High');
    return {
      type: 'memora',
      title: `High & Critical Incidents — ${critInc.length} Total`,
      summary: `MEMORA found **${critOnly.length} Critical** and **${highOnly.length} High** severity incidents across both platforms. All resolved. Avg resolution for Critical: **${avg(critOnly.map(i => i.ResolutionTime_hrs)).toFixed(1)} hrs**. Fastest resolution: **${Math.min(...critOnly.map(i=>i.ResolutionTime_hrs))} hrs**.`,
      table: critInc.sort((a, b) => new Date(b.Date) - new Date(a.Date)).map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, System: i.System, 'Root Cause': i.RootCause, 'Resolved By': i.ResolvedBy, 'Time (hrs)': i.ResolutionTime_hrs })),
      memory: `MEMORA pattern: Critical incidents (API rate limiting, TRAI DND) had fastest resolution due to immediate escalation protocols. High-severity incidents related to NLP accuracy required extended root-cause analysis. Recommend formalizing on-call escalation SLA for all P0 issues within 2 hrs.`,
      timeline: critInc.map(i => ({ id: i.IncidentID, date: i.Date, title: i.Category, severity: i.Severity, desc: i.RootCause, resolvedBy: i.ResolvedBy, hours: i.ResolutionTime_hrs })),
      source: 'MEMORA Knowledge Base → Severity Filter: Critical + High → Sorted by Date'
    };
  }

  // ── Recurring patterns / RCA ──────────────────────────────
  if (q.match(/recurring|pattern|repeat|common.*(issue|problem|fail)|what.*(keep|keeps).*(failing|happening)|root cause analysis|rca|rca|most.*incident|frequent/)) {
    const categories = {};
    incidents.forEach(i => { categories[i.Category] = (categories[i.Category] || 0) + 1; });
    const sorted = Object.entries(categories).sort((a, b) => b[1] - a[1]);
    return {
      type: 'memora',
      title: 'Root Cause Analysis — Recurring Incident Patterns',
      summary: `MEMORA deep analysis across **${incidents.length} incidents**: **${sorted[0][0]}** is the most incident-prone category with **${sorted[0][1]} occurrences**. Integration and infrastructure failures account for 70% of all incidents. This pattern repeats across both ACE Platform and Sweet Hello AI.`,
      table: sorted.map(([cat, count]) => ({
        Category: cat,
        'Incident Count': count,
        'Frequency %': `${Math.round((count / incidents.length) * 100)}%`,
        'Avg Resolution (hrs)': avg(incidents.filter(i => i.Category === cat).map(i => i.ResolutionTime_hrs)).toFixed(1),
        'Systems Affected': [...new Set(incidents.filter(i => i.Category === cat).map(i => i.System))].join(', ')
      })),
      memory: `MEMORA recommendation: Invest in automated connector health checks and proactive carrier failover testing. These two categories alone account for 60% of all incidents. A monthly incident retrospective process would significantly reduce MTTR across both platforms.`,
      source: 'MEMORA Knowledge Base → All Incidents → Category Grouping → Pattern Analysis'
    };
  }

  // ── API Gateway / rate limiting ───────────────────────────
  if (q.match(/api.*(fail|issue|gateway|limit|rate)|rate.*(limit|issue)|client.*(lockout|block|deny)/)) {
    const apiInc = incidents.filter(i => i.Category.toLowerCase().includes('api') || i.RootCause.toLowerCase().includes('rate limit'));
    return {
      type: 'memora',
      title: 'API Gateway — Incident Memory',
      summary: `MEMORA found **${apiInc.length} API-related incident(s)**. Most critical: INC-005 (Feb 2024) — enterprise client lockout due to rate limit misconfiguration. Resolved in **${apiInc[0]?.ResolutionTime_hrs} hrs**. Impact: enterprise client access blocked for ${apiInc[0]?.ResolutionTime_hrs} hours.`,
      table: apiInc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, 'Resolved By': i.ResolvedBy, Resolution: i.Resolution })),
      memory: `MEMORA insight: Rate limit configurations must be validated in staging before production deployment. Whitelisting enterprise IP ranges is a required step for all new enterprise onboardings. INC-005 caused direct client escalation — avoid at all costs.`,
      timeline: apiInc.map(i => ({ id: i.IncidentID, date: i.Date, title: i.Category, severity: i.Severity, desc: i.RootCause, resolvedBy: i.ResolvedBy, hours: i.ResolutionTime_hrs })),
      source: 'MEMORA Knowledge Base → API Gateway / Rate Limit Filter'
    };
  }

  // ── NLP / AI/ML incidents ─────────────────────────────────
  if (q.match(/nlp|intent class|ml|ai.*issue|model.*fail|accuracy|bert|training|ai.*(incident|problem)/)) {
    const nlpInc = incidents.filter(i => i.Category.toLowerCase().includes('nlp') || i.Category.toLowerCase().includes('ai') || i.RootCause.toLowerCase().includes('nlp') || i.RootCause.toLowerCase().includes('model'));
    return {
      type: 'memora',
      title: 'NLP & AI/ML — Incident Memory',
      summary: `MEMORA found **${nlpInc.length} NLP/AI-related incident(s)**. Root causes include outdated intent models and insufficient training data coverage. Avg resolution: **${avg(nlpInc.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs** — longer than avg due to model retraining requirements.`,
      table: nlpInc.map(i => ({ ID: i.IncidentID, Date: i.Date, Category: i.Category, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy })),
      memory: `MEMORA AI insight: NLP incidents take longest to resolve (avg ${avg(nlpInc.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs) because they require model retraining cycles. Recommend establishing monthly model accuracy reviews with automated regression test suites. Intent classification drift is predictable and preventable.`,
      source: 'MEMORA Knowledge Base → NLP/AI/ML Category Filter'
    };
  }

  // ── Knowledge graph issues ────────────────────────────────
  if (q.match(/knowledge graph|graph.*(timeout|issue|fail)|large.*(dataset|data)|redis|cache.*(issue|miss)/)) {
    const kgInc = incidents.filter(i => i.Category.toLowerCase().includes('knowledge') || i.RootCause.toLowerCase().includes('graph') || i.RootCause.toLowerCase().includes('timeout'));
    return {
      type: 'memora',
      title: 'Knowledge Graph — Incident Memory',
      summary: `MEMORA found **${kgInc.length} knowledge graph related incident(s)**. Key issue: Graph traversal timeout on large enterprise datasets (>500k nodes) — resolved via pagination + Redis caching. This is a known scaling boundary.`,
      table: kgInc.map(i => ({ ID: i.IncidentID, Date: i.Date, 'Root Cause': i.RootCause, 'Resolved By': i.ResolvedBy, Resolution: i.Resolution })),
      memory: `MEMORA scaling lesson from INC-007: Always benchmark graph queries against datasets >500k nodes in staging. Redis subgraph caching is now standard for all enterprise deployments. Pagination of graph traversal is non-negotiable above 100k nodes.`,
      source: 'MEMORA Knowledge Base → Knowledge Graph Filter'
    };
  }

  // ── Full incident log ────────────────────────────────────
  if (q.match(/all incident|full history|incident log|show.*incident|list.*incident/)) {
    const sorted = [...incidents].sort((a, b) => new Date(b.Date) - new Date(a.Date));
    return {
      type: 'memora',
      title: `Full Incident Log — ${incidents.length} Records`,
      summary: `MEMORA complete incident history: **${incidents.length} incidents** across ACE Platform (${incidents.filter(i=>i.System.includes('ACE')).length}) and Sweet Hello AI (${incidents.filter(i=>i.System.includes('Sweet Hello')).length}). All resolved. Avg resolution: **${avg(incidents.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs**.`,
      table: sorted.map(i => ({ ID: i.IncidentID, Date: i.Date, System: i.System, Severity: i.Severity, Category: i.Category, 'Resolved By': i.ResolvedBy, 'Time (hrs)': i.ResolutionTime_hrs })),
      memory: `MEMORA full retrieval: Infra Team and Dev Team – Integrations are the top resolving teams. Recommend quarterly incident retrospectives to track pattern closure and verify remediations are holding.`,
      timeline: sorted.map(i => ({ id: i.IncidentID, date: i.Date, title: i.Category, severity: i.Severity, desc: i.RootCause, resolvedBy: i.ResolvedBy, hours: i.ResolutionTime_hrs })),
      source: 'MEMORA Knowledge Base → Full Retrieval → Sorted by Date'
    };
  }

  // ── What is MEMORA ────────────────────────────────────────
  if (q.match(/what is memora|memora.*work|how.*memora|tell.*memora|about memora|explain memora/)) {
    return {
      type: 'memora',
      title: 'MEMORA AI — Enterprise Operational Memory Platform',
      summary: `**MEMORA** is Algonox's Enterprise Operational Memory Platform. It captures, structures, and retrieves organizational intelligence from incidents, workflows, expert knowledge, and historical resolutions — ensuring operational knowledge is never lost when employees leave or systems change.`,
      table: [
        { Capability: 'Incident Memory', Description: 'Retrieves historical incidents, root causes, and resolutions instantly' },
        { Capability: 'Expert Discovery', Description: 'Maps who resolved what and recommends the right SME for each issue' },
        { Capability: 'Pattern Analysis', Description: 'Identifies recurring failure categories and systemic risks' },
        { Capability: 'Compliance Memory', Description: 'Tracks regulatory incidents and enforced remediation actions' },
        { Capability: 'Root Cause Intelligence', Description: 'Infers likely causes using dependency and workflow relationships' },
        { Capability: 'Knowledge Graph', Description: 'Connects employees, systems, incidents, vendors, and resolutions' },
      ],
      memory: `MEMORA solves the #1 enterprise problem: operational intelligence loss. When engineers leave, their troubleshooting knowledge stays in MEMORA forever. This platform currently indexes ${incidents.length} incidents, ${team.length} team experts, and cross-platform resolution patterns.`,
      source: 'MEMORA Platform Intelligence Guide + Algonox Product Architecture'
    };
  }

  // ══════════════════════════════════════════════════════════
  // IRIS — LIVE DATA QUERIES
  // ══════════════════════════════════════════════════════════

  if (q.match(/top.*(exec|account|perform|sales|person|ae|rep)|best.*(exec|perform|deal|rep)|who.*(sold|closed|highest)/)) {
    const totals = {};
    wonDeals.forEach(r => { totals[r.AccountExec] = (totals[r.AccountExec] || 0) + r.DealValue; });
    const sorted = Object.entries(totals).sort((a, b) => b[1] - a[1]);
    return {
      type: 'iris',
      title: 'Account Executive Performance — Closed Won Revenue',
      summary: `**${sorted[0][0]}** leads with ${fmt(sorted[0][1])} in closed revenue across all verticals.`,
      table: sorted.map(([exec, rev]) => ({ 'Account Executive': exec, 'Closed Revenue': fmt(rev), 'Deals Won': wonDeals.filter(r=>r.AccountExec===exec).length, 'Win Rate': `${Math.round((wonDeals.filter(r=>r.AccountExec===exec).length/pipeline.filter(r=>r.AccountExec===exec).length)*100)}%` })),
      chart: { type: 'bar', label: 'Revenue by Account Executive (₹L)', labels: sorted.map(([e]) => e.split(' ')[0]), datasets: [{ label: 'Closed Revenue (₹L)', data: sorted.map(([, v]) => +(v/100000).toFixed(1)), color: '#4F46E5' }] },
      source: 'Pipeline Sheet → Closed Won → Grouped by Account Executive'
    };
  }

  if (q.match(/revenue.*(product|ace|sweet hello|platform)|product.*(revenue|sales|perform|split|breakdown)/)) {
    const totals = {};
    revenue.forEach(r => { totals[r.Product] = (totals[r.Product] || 0) + r.Revenue; });
    const entries = Object.entries(totals).sort((a, b) => b[1] - a[1]);
    const totalRev = entries.reduce((s, [, v]) => s + v, 0);
    return {
      type: 'iris',
      title: 'Revenue by Product — Full Year 2025',
      summary: `**ACE Platform** contributes ${Math.round((totals['ACE Platform']/totalRev)*100)}% of total revenue. Sweet Hello AI is at ${Math.round((totals['Sweet Hello AI']/totalRev)*100)}%.`,
      table: entries.map(([product, rev]) => ({ Product: product, 'Total Revenue': fmt(rev), 'Share': `${Math.round((rev/totalRev)*100)}%`, 'Total Deals': revenue.filter(r=>r.Product===product).reduce((s,r)=>s+r.Deals,0) })),
      chart: { type: 'pie', label: 'Revenue Share by Product', labels: entries.map(([p]) => p), datasets: [{ data: entries.map(([, v]) => +(v/100000).toFixed(1)), colors: ['#4F46E5', '#0284C7'] }] },
      source: 'Revenue Sheet → Grouped by Product → FY2025'
    };
  }

  if (q.match(/quarter|trend|q1|q2|q3|q4|growth|monthly|annual/)) {
    const quarters = [...new Set(revenue.map(r => r.Quarter))];
    const aceData = quarters.map(q => revenue.filter(r=>r.Quarter===q&&r.Product==='ACE Platform').reduce((s,r)=>s+r.Revenue,0));
    const shData = quarters.map(q => revenue.filter(r=>r.Quarter===q&&r.Product==='Sweet Hello AI').reduce((s,r)=>s+r.Revenue,0));
    return {
      type: 'iris',
      title: 'Quarterly Revenue Trend — 2025',
      summary: `Q4 2025 was the strongest quarter at **${fmt(aceData[3]+shData[3])}** combined. ACE Platform grew **24%** QoQ in Q4.`,
      table: quarters.map((q, i) => ({ Quarter: q, 'ACE Platform': fmt(aceData[i]), 'Sweet Hello AI': fmt(shData[i]), 'Combined': fmt(aceData[i]+shData[i]) })),
      chart: { type: 'line', label: 'Quarterly Revenue Trend (₹L)', labels: quarters, datasets: [{ label: 'ACE Platform', data: aceData.map(v => +(v/100000).toFixed(1)), color: '#4F46E5' }, { label: 'Sweet Hello AI', data: shData.map(v => +(v/100000).toFixed(1)), color: '#0284C7' }] },
      source: 'Revenue Sheet → Quarterly breakdown → FY2025'
    };
  }

  if (q.match(/vertical|industry|bfsi|healthtech|insurtech|saas|logistics|sector/)) {
    const vTotals = {};
    wonDeals.forEach(r => { vTotals[r.Vertical] = (vTotals[r.Vertical] || 0) + r.DealValue; });
    const sorted = Object.entries(vTotals).sort((a, b) => b[1] - a[1]);
    return {
      type: 'iris',
      title: 'Revenue by Vertical — Closed Won Deals',
      summary: `**${sorted[0][0]}** is Algonox's strongest vertical at ${fmt(sorted[0][1])}.`,
      table: sorted.map(([v, rev]) => ({ Vertical: v, 'Revenue Won': fmt(rev), 'Deals': wonDeals.filter(r=>r.Vertical===v).length, 'Avg Deal': fmt(rev/wonDeals.filter(r=>r.Vertical===v).length) })),
      chart: { type: 'bar', label: 'Revenue by Vertical (₹L)', labels: sorted.map(([v]) => v), datasets: [{ label: 'Closed Revenue (₹L)', data: sorted.map(([, v]) => +(v/100000).toFixed(1)), color: '#059669' }] },
      source: 'Pipeline Sheet → Closed Won → Grouped by Vertical'
    };
  }

  if (q.match(/us|india|market|region|geography|geo/)) {
    const rTotals = {};
    wonDeals.forEach(r => { rTotals[r.Region] = (rTotals[r.Region] || 0) + r.DealValue; });
    const entries = Object.entries(rTotals);
    const total = entries.reduce((s, [, v]) => s + v, 0);
    return {
      type: 'iris',
      title: 'Revenue Split — US vs India Market',
      summary: `US market contributes **${Math.round((rTotals['US']/total)*100)}%** of Algonox's closed revenue. India at **${Math.round((rTotals['India']/total)*100)}%**.`,
      table: entries.map(([region, rev]) => ({ Market: region, 'Closed Revenue': fmt(rev), 'Share': `${Math.round((rev/total)*100)}%`, 'Deals Won': wonDeals.filter(r=>r.Region===region).length })),
      chart: { type: 'pie', label: 'Revenue by Market', labels: entries.map(([r]) => r + ' Market'), datasets: [{ data: entries.map(([,v])=>+(v/100000).toFixed(1)), colors: ['#4F46E5','#059669'] }] },
      source: 'Pipeline Sheet → Closed Won → Grouped by Region'
    };
  }

  if (q.match(/win rate|conversion|lost|pipeline|close rate|funnel/)) {
    const total = pipeline.length;
    const won = wonDeals.length;
    const winRate = Math.round((won / total) * 100);
    const byExec = [...new Set(pipeline.map(r => r.AccountExec))];
    return {
      type: 'iris',
      title: `Pipeline Win Rate — ${winRate}% Overall`,
      summary: `Algonox has converted **${won} of ${total}** deals. **${lostDeals.length}** deals lost — primarily in Logistics and InsurTech.`,
      table: byExec.map(exec => { const all = pipeline.filter(r=>r.AccountExec===exec); const w = all.filter(r=>r.Stage==='Closed Won').length; return { 'Account Executive': exec, 'Won': w, 'Lost': all.length-w, 'Win Rate': `${Math.round((w/all.length)*100)}%`, 'Rev Won': fmt(wonDeals.filter(r=>r.AccountExec===exec).reduce((s,r)=>s+r.DealValue,0)) }; }),
      chart: { type: 'bar', label: 'Win / Loss by Account Executive', labels: byExec.map(e=>e.split(' ')[0]), datasets: [{ label: 'Won', data: byExec.map(e=>wonDeals.filter(r=>r.AccountExec===e).length), color: '#059669' }, { label: 'Lost', data: byExec.map(e=>lostDeals.filter(r=>r.AccountExec===e).length), color: '#DC2626' }] },
      source: 'Pipeline Sheet → All stages → Win/Loss analysis'
    };
  }

  if (q.match(/lead source|inbound|outbound|referral|apollo|linkedin|channel/)) {
    const srcTotals = {};
    wonDeals.forEach(r => { srcTotals[r.LeadSource] = (srcTotals[r.LeadSource] || 0) + r.DealValue; });
    const entries = Object.entries(srcTotals).sort((a,b)=>b[1]-a[1]);
    return {
      type: 'iris',
      title: 'Revenue by Lead Source',
      summary: `**${entries[0][0]}** is the highest-value lead channel at ${fmt(entries[0][1])}.`,
      table: entries.map(([src, rev]) => ({ 'Lead Source': src, 'Revenue Won': fmt(rev), 'Deals': wonDeals.filter(r=>r.LeadSource===src).length, 'Avg Deal': fmt(rev/wonDeals.filter(r=>r.LeadSource===src).length) })),
      chart: { type: 'pie', label: 'Revenue by Lead Source', labels: entries.map(([s])=>s), datasets: [{ data: entries.map(([,v])=>+(v/100000).toFixed(1)), colors: ['#4F46E5','#0284C7','#059669'] }] },
      source: 'Pipeline Sheet → Closed Won → Grouped by Lead Source'
    };
  }

  if (q.match(/team|employee|staff|headcount|people|department|who work/)) {
    const deptCount = {};
    team.forEach(e => { deptCount[e.Team] = (deptCount[e.Team] || 0) + 1; });
    return {
      type: 'iris',
      title: `Algonox Team Directory — ${team.length} Members`,
      summary: `Algonox has **${team.length} team members** across ${Object.keys(deptCount).length} functions.`,
      table: team.map(e => ({ Name: e.Name, Team: e.Team, Role: e.Role, Location: e.Location })),
      chart: { type: 'pie', label: 'Team Headcount by Function', labels: Object.keys(deptCount), datasets: [{ data: Object.values(deptCount), colors: ['#4F46E5','#0284C7','#059669','#D97706','#DC2626'] }] },
      source: 'Team Sheet → All records'
    };
  }

  if (q.match(/prospect|pipeline status|leads|iris campaign|qualified|score|demo|outreach/)) {
    const stageCounts = {};
    prospects.forEach(p => { stageCounts[p.Stage] = (stageCounts[p.Stage] || 0) + 1; });
    const sorted = prospects.slice().sort((a,b)=>b.Score-a.Score);
    return {
      type: 'iris',
      title: `IRIS Campaign — ${prospects.length} Active Prospects`,
      summary: `**${prospects.filter(p=>p.Score>=7.5).length} high-priority prospects** (score ≥ 7.5). Top prospect: ${sorted[0].Company} (Score: ${sorted[0].Score}).`,
      table: sorted.map(p => ({ Company: p.Company, Vertical: p.Vertical, Country: p.Country, Stage: p.Stage, Score: p.Score, Contact: p.Persona })),
      chart: { type: 'bar', label: 'Prospect Pipeline by Stage', labels: Object.keys(stageCounts), datasets: [{ label: 'Prospects', data: Object.values(stageCounts), color: '#4F46E5' }] },
      source: 'Prospects Sheet → Sorted by Score → All stages'
    };
  }

  // ── Combined: Algonox Overview ────────────────────────────
  if (q.match(/overview|dashboard|summary|overall|how are we|company|algonox/)) {
    const totalRev = revenue.reduce((s,r)=>s+r.Revenue,0);
    const totalDeals = revenue.reduce((s,r)=>s+r.Deals,0);
    const quarters = [...new Set(revenue.map(r=>r.Quarter))];
    const qRev = quarters.map(q => revenue.filter(r=>r.Quarter===q).reduce((s,r)=>s+r.Revenue,0));
    return {
      type: 'combined',
      title: 'Algonox Intelligence Overview — FY2025',
      iris: {
        label: 'Live Business Metrics (IRIS)',
        summary: `Total FY2025 revenue: **${fmt(totalRev)}** across **${totalDeals} deals closed**. Win rate: **${Math.round((wonDeals.length/pipeline.length)*100)}%**. Active prospects: **${prospects.length}**.`,
        table: [{ Metric: 'FY2025 Total Revenue', Value: fmt(totalRev) }, { Metric: 'Total Deals Closed', Value: totalDeals }, { Metric: 'Pipeline Win Rate', Value: `${Math.round((wonDeals.length/pipeline.length)*100)}%` }, { Metric: 'Best Quarter', Value: `Q4 2025 — ${fmt(qRev[3])}` }, { Metric: 'Active Prospects', Value: prospects.length }],
        chart: { type: 'line', label: 'Quarterly Revenue (₹L)', labels: quarters, datasets: [{ label: 'Combined Revenue', data: qRev.map(v=>+(v/100000).toFixed(1)), color: '#4F46E5' }] }
      },
      memora: {
        label: 'Operational Memory (MEMORA)',
        summary: `**${incidents.length} incidents** indexed. Avg resolution: **${avg(incidents.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs**. **${incidents.filter(i=>i.Severity==='Critical').length} Critical** incidents — all resolved. ${team.length} experts mapped across ${[...new Set(team.map(t=>t.Team))].length} teams.`,
        table: [{ Metric: 'Total Incidents Indexed', Value: incidents.length }, { Metric: 'Avg Resolution Time', Value: `${avg(incidents.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs` }, { Metric: 'Critical Incidents', Value: incidents.filter(i=>i.Severity==='Critical').length }, { Metric: 'Experts Mapped', Value: team.length }, { Metric: 'Recurring Patterns Found', Value: [...new Set(incidents.map(i=>i.Category))].length }]
      },
      source: 'All Sheets → IRIS + MEMORA Combined Intelligence'
    };
  }

  // ── Sweet Hello AI combined ───────────────────────────────
  if (q.match(/sweet hello.*(perform|revenue|sales|overview)|voice agent.*(perform|overview)/)) {
    const shWon = wonDeals.filter(r=>r.Product==='Sweet Hello AI');
    const shRev = shWon.reduce((s,r)=>s+r.DealValue,0);
    const shInc = incidents.filter(i=>i.System.includes('Sweet Hello'));
    const shRevQ = [...new Set(revenue.map(r=>r.Quarter))].map(q=>({ q, v: revenue.find(r=>r.Quarter===q&&r.Product==='Sweet Hello AI')?.Revenue||0 }));
    return {
      type: 'combined',
      title: 'Sweet Hello AI — Full Platform Intelligence',
      iris: { label: 'Live Sales Data', summary: `Sweet Hello AI closed **${fmt(shRev)}** across **${shWon.length} won deals**. Top verticals: BFSI and HealthTech.`, table: shWon.map(r=>({ Month: r.Month, Vertical: r.Vertical, Deal: fmt(r.DealValue), AE: r.AccountExec })), chart: { type: 'bar', label: 'Sweet Hello AI — Quarterly Revenue (₹L)', labels: shRevQ.map(x=>x.q), datasets: [{ label: 'Revenue (₹L)', data: shRevQ.map(x=>+(x.v/100000).toFixed(1)), color: '#0284C7' }] } },
      memora: { label: 'Incident & Operational Memory', summary: `**${shInc.length} past incidents** on Sweet Hello AI. Main issues: voice quality, NLP accuracy, compliance. Avg resolution: **${avg(shInc.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs**.`, table: shInc.map(i=>({ ID: i.IncidentID, Category: i.Category, 'Root Cause': i.RootCause, Resolution: i.Resolution })) },
      source: 'Pipeline + Revenue Sheets (IRIS) + Incidents Sheet (MEMORA) → Sweet Hello AI'
    };
  }

  // ── Fallback ──────────────────────────────────────────────
  return {
    type: 'fallback',
    title: 'Try one of these queries',
    summary: 'MEMORA and IRIS are ready. Here are questions you can ask:',
    suggestions: [
      'Has the ACE Platform had incidents before?',
      'What issues have occurred with Sweet Hello AI voice calls?',
      'Who should I contact for voice call quality issues?',
      'Show all critical and high severity incidents',
      'What are the most recurring incident patterns?',
      'Show compliance incident history',
      'Who is the expert for compliance and TRAI regulations?',
      'What is MEMORA and how does it work?',
      'Show revenue by product',
      'Who is the top account executive?',
      'Show quarterly revenue trend',
      'Give me an Algonox company overview',
    ]
  };
}

app.post('/api/query', (req, res) => {
  const { question } = req.body;
  if (!question || question.trim().length < 3) return res.json({ error: 'Please type a question.' });
  try {
    const data = loadData();
    const result = queryEngine(question, data);
    setTimeout(() => res.json(result), 600 + Math.random() * 400);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Error reading data: ' + err.message });
  }
});

app.get('/api/data', (req, res) => {
  try { res.json(loadData()); } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/memora-questions', (req, res) => {
  res.json([
    { q: 'Has the ACE Platform had incidents before?', tag: 'ACE History', engine: 'memora' },
    { q: 'What issues have occurred with Sweet Hello AI voice calls?', tag: 'Voice Incidents', engine: 'memora' },
    { q: 'Show all critical and high severity incidents', tag: 'Critical Severity', engine: 'memora' },
    { q: 'What are the most recurring incident patterns?', tag: 'Root Cause Analysis', engine: 'memora' },
    { q: 'Who should I contact for voice call quality issues?', tag: 'Expert Discovery', engine: 'memora' },
    { q: 'Who is the expert for compliance and TRAI regulations?', tag: 'Expert Discovery', engine: 'memora' },
    { q: 'Show compliance incident history', tag: 'Compliance', engine: 'memora' },
    { q: 'What happened with the API gateway rate limiting issue?', tag: 'API History', engine: 'memora' },
    { q: 'Show all incidents related to NLP and intent classification', tag: 'AI/ML Incidents', engine: 'memora' },
    { q: 'What recurring issues should we fix in the ACE Platform?', tag: 'Pattern Analysis', engine: 'memora' },
    { q: 'Who resolved the TRAI DND compliance issue?', tag: 'Expert Discovery', engine: 'memora' },
    { q: 'Show the full Algonox incident log', tag: 'Full Log', engine: 'memora' },
    { q: 'What issues has the knowledge graph had?', tag: 'AI/ML Incidents', engine: 'memora' },
    { q: 'Who handles Salesforce and SAP integration issues?', tag: 'Expert Discovery', engine: 'memora' },
    { q: 'What is MEMORA and how does it work?', tag: 'About MEMORA', engine: 'memora' },
  ]);
});

const PORT = 3000;
app.listen(PORT, () => console.log(`Algonox Intelligence Platform → http://localhost:${PORT}`));