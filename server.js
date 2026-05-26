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

function enrichTeam(member) {
  const isLeft = member.Status === 'Left' || !!member.LeftYear;
  return {
    ...member,
    email: member.Email || null,
    leftYear: isLeft ? (member.LeftYear || 'Unknown') : null,
    isLeft
  };
}

function queryEngine(question, data) {
  const q = question.toLowerCase();
  const { pipeline, revenue, incidents, team, prospects } = data;
  const wonDeals = pipeline.filter(r => r.Stage === 'Closed Won');
  const lostDeals = pipeline.filter(r => r.Stage === 'Closed Lost');

  // ═══════════════════════════════════════════════════════
  // MEMORA — incident-specific queries (checked first)
  // ═══════════════════════════════════════════════════════

  // ALG-INC-001 — Agent task queue deadlock / workflow engine
  if (q.match(/task queue|deadlock|parallel.*exec|agent.*hang|workflow.*stuck|thread.*pool|agent task|approval.*loop|multi.?step.*loop|workflow.*loop/)) {
    const inc = incidents.filter(i => ['ALG-INC-001','ALG-INC-009'].includes(i.IncidentID));
    const experts = team.filter(e => e.Team === 'Engineering' || (e.Expert_In||'').includes('ACE')).map(enrichTeam);
    return {
      type: 'memora', title: 'Workflow Engine — Task Queue & Loop Issues',
      summary: `Found **${inc.length} incident(s)** related to workflow engine failures on ACE Platform. Root causes: agent task queue deadlock during parallel execution, and multi-step approval loop on SAP integration.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Identify the deadlocked agent tasks using the ACE Platform task manager dashboard.','Increase the thread pool size in the worker config from default 10 to 25.','Add a 30-second timeout guard on all parallel agent task executions.','Deploy the patched config and restart the workflow engine.','Monitor task queue depth for 15 minutes to confirm resolution.'],
      memory: `Fix: Increase thread pool size and add timeout guards on agent tasks. For SAP loops — patch the connector and add idempotency keys on approval callbacks to prevent re-triggering.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-001, ALG-INC-009 → Workflow Engine'
    };
  }

  // ALG-INC-002 — Call drops / outbound drop rate
  if (q.match(/call.?drop|drop.?rate|outbound.*call|toll.?free|carrier|call.*fail|calls.*drop|dropping/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-002');
    const experts = team.filter(e => e.Team === 'Infrastructure').map(enrichTeam);
    return {
      type: 'memora', title: 'Outbound Call Drops — Sweet Hello AI',
      summary: `**ALG-INC-002** (Jul 2023): Outbound call drop rate spiked to **18%** on US toll-free lines. Severity: **High**. Resolved in **4 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Check carrier health dashboard — identify spike in drop rate on primary US toll-free route.','Switch outbound routing to secondary carrier failover (pre-configured in Sweet Hello AI routing config).','Validate drop rate on secondary carrier — confirm below 1%.','Contact primary carrier to report and log the route issue.','Once primary carrier confirms fix, retest and restore as default route.','Configure automated carrier health checks every 15 minutes going forward.'],
      memory: `Resolution: Carrier failover to secondary provider reduced drop rate from 18% to 0.4%. Always maintain a tested secondary carrier failover config. Validate carrier health before high-volume outbound campaigns.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-002 → Sweet Hello AI Voice Infra'
    };
  }

  // ALG-INC-003 — Salesforce OAuth / connector failure
  if (q.match(/salesforce|oauth|connector.*fail|pipeline.*stall|token.*refresh|token.*fail|crm.*connect|integration.*fail|salesforce.*not work/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-003');
    const experts = team.filter(e => (e.Expert_In||'').toLowerCase().includes('salesforce') || e.Team === 'Engineering').map(enrichTeam);
    return {
      type: 'memora', title: 'Salesforce OAuth Connector Failure — ACE Platform',
      summary: `**ALG-INC-003** (Sep 2023): Salesforce OAuth token refresh failed, causing ACE Platform pipeline stall. Severity: **Medium**. Resolved in **3 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Identify the stalled pipeline — check ACE Platform connector logs for OAuth 401 errors.','Manually re-authenticate the Salesforce OAuth connection from the connector settings panel.','Refresh the OAuth scope to include all required Salesforce objects.','Add exponential backoff retry logic to the connector (max 3 retries, 2s/4s/8s intervals).','Implement proactive token refresh 5 minutes before token expiry using a scheduled job.','Test the connector end-to-end with a sample Salesforce record sync.'],
      memory: `Fix: Refresh OAuth scope and add auto-retry with exponential backoff. Root cause was token expiry without silent renewal. Implement proactive token refresh 5 mins before expiry to prevent pipeline stalls.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-003 → ACE Platform Data Connectors'
    };
  }

  // ALG-INC-004 — NLP accuracy drop / intent classification
  if (q.match(/nlp|intent.*class|accuracy.*drop|model.*update|classification.*fail|ai.*accuracy|model.*rollback|intent.*wrong|voice.*understand|speech.*understand/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-004');
    const experts = team.filter(e => e.Team === 'AI/ML').map(enrichTeam);
    return {
      type: 'memora', title: 'NLP Intent Classification Accuracy Drop — Sweet Hello AI',
      summary: `**ALG-INC-004** (Nov 2023): Intent classification accuracy dropped to **61%** after model update. Severity: **High**. Required rollback and retraining. Resolved in **8 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Confirm accuracy drop by running the full intent test suite against the new model — record score (61%).','Immediately roll back the NLP model to the last stable version (v2.3.1) via the AI model registry.','Restart the Sweet Hello AI voice processing service to load the rolled-back model.','Validate accuracy is restored — re-run intent test suite, confirm score above 85%.','Initiate retraining pipeline on the new model with corrected training data.','Set 85% accuracy as a hard deployment gate — block model updates that fall below this threshold.'],
      memory: `Fix: Rolled back to v2.3.1 model; retraining pipeline initiated. Before any model update, run full regression on intent test suite. Maintain minimum 85% accuracy threshold as deployment gate.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-004 → Sweet Hello AI NLP Engine'
    };
  }

  // ALG-INC-005 — API Gateway / rate limiting / client lockout
  if (q.match(/rate.?limit|api.*gateway|client.*lockout|lockout|ip.*whitelist|enterprise.*block|api.*block|client.*block|api.*access.*fail/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-005');
    const experts = team.filter(e => (e.Expert_In||'').includes('API') || e.Team === 'Infrastructure').map(enrichTeam);
    return {
      type: 'memora', title: 'API Gateway Rate Limit Misconfiguration — ACE Platform',
      summary: `**ALG-INC-005** (Feb 2024): Rate limiting misconfiguration locked out an enterprise client from ACE Platform APIs. Severity: **Critical**. Resolved in **2 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Identify the locked-out client — check API gateway logs for 429 Too Many Requests on their IP range.','Temporarily increase rate limit threshold for the affected enterprise client as an emergency measure.','Correct the rate limit rule in the API gateway config — set enterprise tier to 10,000 req/min.','Whitelist the enterprise client IP range in the gateway allowlist.','Deploy the corrected config and confirm client access is restored.','Add IP whitelisting as a mandatory checklist item for all new enterprise onboardings.'],
      memory: `Fix: Corrected rate limit rules and whitelisted enterprise IP ranges. Mandatory step for all new enterprise onboardings: whitelist client IP ranges in the API gateway before go-live. Always validate rate limit config in staging first.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-005 → ACE Platform API Gateway'
    };
  }

  // ALG-INC-006 — IVR loop / Hindi language fallback
  if (q.match(/ivr|language.*loop|hindi|fallback.*loop|call.*routing|routing.*loop|ivr.*loop|language.*fallback|stuck.*ivr|ivr.*stuck/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-006');
    const experts = team.filter(e => (e.Team||'').toLowerCase().includes('voice') || (e.Expert_In||'').toLowerCase().includes('voice')).map(enrichTeam);
    return {
      type: 'memora', title: 'IVR Logic Loop on Language Fallback — Sweet Hello AI',
      summary: `**ALG-INC-006** (May 2024): IVR logic entered an infinite loop on Hindi language fallback path, causing callers to be stuck. Severity: **Medium**. Resolved in **5 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Reproduce the IVR loop by triggering the Hindi language path in the test environment.','Map the IVR decision tree — identify the missing exit node on the Hindi fallback path.','Add an explicit fallback exit node that routes to the default English IVR after 2 failed language detections.','Test all language paths (English, Hindi, Tamil) in staging to confirm no loops.','Deploy the patched IVR routing config.','Add non-English IVR path testing to the pre-deployment checklist.'],
      memory: `Fix: Fixed language detection tree and added an explicit fallback exit node. Every IVR language path must have a tested exit condition. Test all non-English fallback paths before deployment.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-006 → Sweet Hello AI Call Routing'
    };
  }

  // ALG-INC-007 — Knowledge graph timeout / large dataset
  if (q.match(/knowledge.*graph|graph.*timeout|graph.*slow|large.*dataset|500k|node.*timeout|query.*timeout|redis.*cache|graph.*perform/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-007');
    const experts = team.filter(e => (e.Team||'').includes('AI') || (e.Expert_In||'').toLowerCase().includes('ace')).map(enrichTeam);
    return {
      type: 'memora', title: 'Knowledge Graph Query Timeout — ACE Platform',
      summary: `**ALG-INC-007** (Aug 2024): Graph traversal timed out on enterprise dataset exceeding **500k nodes**. Severity: **High**. Resolved in **7 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Identify the timing-out queries in the knowledge graph service logs — note the dataset size (>500k nodes).','Add query pagination to limit graph traversal to 10,000 nodes per page with cursor-based iteration.','Identify the most-queried subgraphs (hot paths) and cache them in Redis with a 10-minute TTL.','Test the paginated queries against the full 500k node dataset — confirm response time under 3 seconds.','Deploy the pagination and caching changes to production.','Set Redis subgraph caching as mandatory for all enterprise tenants with datasets above 100k nodes.'],
      memory: `Fix: Implemented query pagination and added Redis cache for hot subgraphs. Always benchmark graph queries against datasets >500k nodes before enterprise deployment. Redis caching is now mandatory for all large-graph tenants.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-007 → ACE Platform Knowledge Graph'
    };
  }

  // ALG-INC-008 — TRAI DND compliance / non-compliant calls
  if (q.match(/trai|dnd|compliance.*breach|non.?compliant|dnd.*sync|compliance.*fail|outbound.*complian|regulatory.*breach|dnd.*fail/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-008');
    const experts = team.filter(e => e.Team === 'Compliance').map(enrichTeam);
    return {
      type: 'memora', title: 'TRAI DND Registry Sync Failure — Sweet Hello AI',
      summary: `**ALG-INC-008** (Nov 2024): TRAI DND registry sync failure caused non-compliant outbound calls. Severity: **Critical**. Resolved in **3 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Immediately halt all active outbound campaigns on Sweet Hello AI.','Trigger an emergency manual DND registry sync via the compliance dashboard.','Audit all calls made during the sync failure window — identify and log non-compliant calls.','Report the incident to the compliance team and file a voluntary disclosure with TRAI if required.','Set up automated daily DND registry sync at 12:00 AM IST.','Add a DND sync status check as a mandatory pre-launch gate for every outbound campaign.'],
      memory: `Fix: Emergency DND sync triggered; daily automated DND validation now enforced. Non-compliance with TRAI can result in license suspension. Always verify DND registry sync status before any outbound campaign launch.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-008 → Sweet Hello AI Compliance Module'
    };
  }

  // ALG-INC-009 — SAP integration / approval loop (also caught by workflow above)
  if (q.match(/sap|sap.*integrat|erp.*connect|sap.*loop|sap.*callback|idempoten|approval.*callback|sap.*fail/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-009');
    const experts = team.filter(e => (e.Expert_In||'').toLowerCase().includes('sap') || (e.Expert_In||'').toLowerCase().includes('erp')).map(enrichTeam);
    return {
      type: 'memora', title: 'SAP Integration Approval Loop — ACE Platform',
      summary: `**ALG-INC-009** (Jan 2025): Multi-step approval workflow entered a loop on SAP integration node. Severity: **Medium**. Resolved in **4 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Identify the looping workflow in the ACE Platform workflow monitor — note the SAP approval node.','Pause the affected workflow instances to stop the loop.','Review the SAP connector callback handler — confirm missing idempotency key.','Add a unique idempotency key (workflow_id + step_id) to every SAP approval callback request.','Patch and deploy the updated SAP connector.','Resume the paused workflow instances and monitor for one complete cycle to confirm no re-looping.'],
      memory: `Fix: Patched SAP connector; added idempotency key on approval callbacks. Root cause was missing idempotency on the callback — SAP was re-triggering the same approval step. Always test multi-step workflows with SAP in a full end-to-end staging environment.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-009 → ACE Platform Agentic Workflow'
    };
  }

  // ALG-INC-010 — Audio codec / garbled speech / BFSI
  if (q.match(/codec|garbled|audio.*quality|speech.*quality|g\.711|voice.*codec|audio.*mismatch|codec.*mismatch|voice.*garbled|distorted.*voice|speech.*garbled/)) {
    const inc = incidents.filter(i => i.IncidentID === 'ALG-INC-010');
    const experts = team.filter(e => e.Team === 'Infrastructure').map(enrichTeam);
    return {
      type: 'memora', title: 'Audio Codec Mismatch — Garbled Speech on BFSI Calls',
      summary: `**ALG-INC-010** (Mar 2025): Audio codec mismatch caused garbled speech during BFSI client calls on Sweet Hello AI. Severity: **Medium**. Resolved in **2 hrs**.`,
      table: inc.map(i => ({ ID: i.IncidentID, Date: i.Date, Severity: i.Severity, 'Root Cause': i.RootCause, Resolution: i.Resolution, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      steps: ['Reproduce the garbled audio by placing a test call through the BFSI carrier endpoint.','Capture the SIP/RTP negotiation logs — identify codec mismatch (G.729 vs G.711 u-law).','Update the Sweet Hello AI media server config to explicitly set G.711 u-law as the preferred codec for all carrier endpoints.','Disable codec auto-negotiation for BFSI carrier routes.','Re-test audio quality on BFSI client calls — confirm clear audio.','Update the carrier integration checklist: always explicitly set codec, never rely on auto-negotiation.'],
      memory: `Fix: Standardised G.711 codec across all carrier endpoints. Codec negotiation must be explicitly set — do not rely on auto-negotiation with BFSI carrier endpoints. G.711 is now the required standard.`,
      expertCards: experts,
      source: 'MEMORA → ALG-INC-010 → Sweet Hello AI Voice Quality'
    };
  }

  // ── Expert discovery ────────────────────────────────────
  if (q.match(/who.*(fix|resolve|handle|expert|contact|know|help|call|reach)|expert.*(for|on|about)|contact.*(for|about)|specialist|sme|point.*contact/)) {
    const keyword = q.match(/voice|call|ivr|audio|codec|carrier/) ? 'Voice/Infra' :
                    q.match(/nlp|intent|model|ai|ml/) ? 'AI/ML' :
                    q.match(/compliance|trai|dnd|gdpr/) ? 'Compliance' :
                    q.match(/salesforce|sap|erp|oauth|connector|integrat/) ? 'Integrations' :
                    q.match(/api|gateway|backend/) ? 'Backend/API' : null;
    const relevantTeam = team.filter(e => {
      if (!keyword) return true;
      const ei = (e.Expert_In||'').toLowerCase();
      if (keyword === 'Voice/Infra') return ei.includes('voice') || ei.includes('infra') || e.Team === 'Infrastructure';
      if (keyword === 'AI/ML') return e.Team === 'AI/ML' || ei.includes('nlp') || ei.includes('ai');
      if (keyword === 'Compliance') return e.Team === 'Compliance' || ei.includes('trai');
      if (keyword === 'Integrations') return ei.includes('salesforce') || ei.includes('sap') || ei.includes('oauth');
      if (keyword === 'Backend/API') return ei.includes('api') || ei.includes('backend') || e.Team === 'Engineering';
      return true;
    }).map(enrichTeam);
    const relevantInc = incidents.filter(i => {
      if (!keyword) return true;
      if (keyword === 'Voice/Infra') return i.System.includes('Sweet Hello') && (i.Category.includes('Voice') || i.Category.includes('Call'));
      if (keyword === 'AI/ML') return i.Category.includes('NLP') || i.Category.includes('Knowledge');
      if (keyword === 'Compliance') return i.Category.includes('Compliance');
      if (keyword === 'Integrations') return i.Category.includes('Connector') || i.Category.includes('Workflow');
      return true;
    });
    return {
      type: 'memora', title: `Expert Discovery${keyword ? ` — ${keyword}` : ''}`,
      summary: `Found **${relevantTeam.length} expert(s)**${keyword ? ` for ${keyword}` : ''}. They have resolved **${relevantInc.length} related incidents**.`,
      expertCards: relevantTeam,
      memory: `Contact these experts directly for fastest resolution. Expertise confirmed via historical incident resolution records.`,
      source: 'MEMORA → Team expertise graph + incident resolution history'
    };
  }

  // ── All incidents / full log ────────────────────────────
  if (q.match(/all incident|full.*log|incident log|show.*all|list.*all.*incident|every incident/)) {
    const sorted = [...incidents].sort((a,b)=>new Date(b.Date)-new Date(a.Date));
    return {
      type: 'memora', title: `Full Incident Log — ${incidents.length} Records`,
      summary: `Complete MEMORA history: **${incidents.length} incidents** across ACE Platform (${incidents.filter(i=>i.System.includes('ACE')).length}) and Sweet Hello AI (${incidents.filter(i=>i.System.includes('Sweet Hello')).length}). All resolved. Avg resolution: **${avg(incidents.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs**.`,
      table: sorted.map(i => ({ ID: i.IncidentID, Date: i.Date, System: i.System, Category: i.Category, Severity: i.Severity, 'Resolved By': i.ResolvedBy, 'Hrs': i.ResolutionTime_hrs })),
      timeline: sorted.map(i => ({ id: i.IncidentID, date: i.Date, title: i.Category, severity: i.Severity, desc: i.RootCause, resolvedBy: i.ResolvedBy, hours: i.ResolutionTime_hrs })),
      memory: `All incidents resolved. Infra Team and Dev Team – Integrations are the top resolving teams. Recommend quarterly incident retrospectives.`,
      source: 'MEMORA → Full retrieval → All incidents'
    };
  }

  // ── Recurring patterns ──────────────────────────────────
  if (q.match(/recurring|pattern|repeat|most.*common|what.*keep.*fail|rca|root cause analysis|frequent/)) {
    const cats = {};
    incidents.forEach(i => { cats[i.Category] = (cats[i.Category]||0)+1; });
    const sorted = Object.entries(cats).sort((a,b)=>b[1]-a[1]);
    return {
      type: 'memora', title: 'Recurring Incident Patterns — Root Cause Analysis',
      summary: `Analysis of **${incidents.length} incidents**: **${sorted[0][0]}** is the most common failure type. Integration and infrastructure failures account for 70% of all incidents.`,
      table: sorted.map(([cat,count]) => ({
        Category: cat, Count: count,
        'Share': `${Math.round((count/incidents.length)*100)}%`,
        'Avg Hrs': avg(incidents.filter(i=>i.Category===cat).map(i=>i.ResolutionTime_hrs)).toFixed(1),
        'System(s)': [...new Set(incidents.filter(i=>i.Category===cat).map(i=>i.System))].join(', ')
      })),
      memory: `Invest in automated connector health checks and proactive carrier failover testing. These two categories alone account for 60% of all incidents.`,
      source: 'MEMORA → All incidents → Category grouping → Pattern analysis'
    };
  }

  // ═══════════════════════════════════════════════════════
  // IRIS — Live data queries
  // ═══════════════════════════════════════════════════════

  if (q.match(/top.*(exec|account|perform|sales|ae|rep)|best.*(exec|perform|rep)|who.*(sold|closed|highest)/)) {
    const totals = {};
    wonDeals.forEach(r => { totals[r.AccountExec]=(totals[r.AccountExec]||0)+r.DealValue; });
    const sorted = Object.entries(totals).sort((a,b)=>b[1]-a[1]);
    return {
      type: 'iris', title: 'Account Executive Performance',
      summary: `**${sorted[0][0]}** leads with ${fmt(sorted[0][1])} in closed revenue.`,
      table: sorted.map(([exec,rev]) => ({ 'Account Executive': exec, 'Revenue': fmt(rev), 'Deals Won': wonDeals.filter(r=>r.AccountExec===exec).length, 'Win Rate': `${Math.round((wonDeals.filter(r=>r.AccountExec===exec).length/pipeline.filter(r=>r.AccountExec===exec).length)*100)}%` })),
      chart: { type: 'bar', label: 'Revenue by AE (₹L)', labels: sorted.map(([e])=>e.split(' ')[0]), datasets: [{ label: 'Revenue (₹L)', data: sorted.map(([,v])=>+(v/100000).toFixed(1)), color: '#4F46E5' }] },
      source: 'Pipeline → Closed Won → By Account Executive'
    };
  }

  if (q.match(/revenue.*(product|ace|sweet hello|platform)|product.*(revenue|split|breakdown)/)) {
    const totals = {};
    revenue.forEach(r => { totals[r.Product]=(totals[r.Product]||0)+r.Revenue; });
    const entries = Object.entries(totals).sort((a,b)=>b[1]-a[1]);
    const totalRev = entries.reduce((s,[,v])=>s+v,0);
    return {
      type: 'iris', title: 'Revenue by Product — FY2025',
      summary: `**ACE Platform** contributes ${Math.round((totals['ACE Platform']/totalRev)*100)}% of total revenue. Sweet Hello AI at ${Math.round((totals['Sweet Hello AI']/totalRev)*100)}%.`,
      table: entries.map(([p,rev]) => ({ Product: p, Revenue: fmt(rev), Share: `${Math.round((rev/totalRev)*100)}%`, Deals: revenue.filter(r=>r.Product===p).reduce((s,r)=>s+r.Deals,0) })),
      chart: { type: 'pie', label: 'Revenue Share by Product', labels: entries.map(([p])=>p), datasets: [{ data: entries.map(([,v])=>+(v/100000).toFixed(1)), colors: ['#4F46E5','#1D6FA4'] }] },
      source: 'Revenue Sheet → By Product → FY2025'
    };
  }

  if (q.match(/quarter|trend|q1|q2|q3|q4|growth|monthly|annual/)) {
    const quarters = [...new Set(revenue.map(r=>r.Quarter))];
    const aceD = quarters.map(q=>revenue.filter(r=>r.Quarter===q&&r.Product==='ACE Platform').reduce((s,r)=>s+r.Revenue,0));
    const shD = quarters.map(q=>revenue.filter(r=>r.Quarter===q&&r.Product==='Sweet Hello AI').reduce((s,r)=>s+r.Revenue,0));
    return {
      type: 'iris', title: 'Quarterly Revenue Trend — 2025',
      summary: `Q4 2025 was the strongest at **${fmt(aceD[3]+shD[3])}** combined. ACE Platform grew **24%** QoQ in Q4.`,
      table: quarters.map((q,i) => ({ Quarter: q, 'ACE Platform': fmt(aceD[i]), 'Sweet Hello AI': fmt(shD[i]), Combined: fmt(aceD[i]+shD[i]) })),
      chart: { type: 'line', label: 'Quarterly Revenue (₹L)', labels: quarters, datasets: [{ label: 'ACE Platform', data: aceD.map(v=>+(v/100000).toFixed(1)), color: '#4F46E5' },{ label: 'Sweet Hello AI', data: shD.map(v=>+(v/100000).toFixed(1)), color: '#1D6FA4' }] },
      source: 'Revenue Sheet → Quarterly breakdown'
    };
  }

  if (q.match(/vertical|industry|bfsi|healthtech|insurtech|saas|logistics|sector/)) {
    const vTotals = {};
    wonDeals.forEach(r => { vTotals[r.Vertical]=(vTotals[r.Vertical]||0)+r.DealValue; });
    const sorted = Object.entries(vTotals).sort((a,b)=>b[1]-a[1]);
    return {
      type: 'iris', title: 'Revenue by Vertical',
      summary: `**${sorted[0][0]}** is the strongest vertical at ${fmt(sorted[0][1])}.`,
      table: sorted.map(([v,rev]) => ({ Vertical: v, Revenue: fmt(rev), Deals: wonDeals.filter(r=>r.Vertical===v).length, 'Avg Deal': fmt(rev/wonDeals.filter(r=>r.Vertical===v).length) })),
      chart: { type: 'bar', label: 'Revenue by Vertical (₹L)', labels: sorted.map(([v])=>v), datasets: [{ label: 'Revenue', data: sorted.map(([,v])=>+(v/100000).toFixed(1)), color: '#059669' }] },
      source: 'Pipeline → Closed Won → By Vertical'
    };
  }

  if (q.match(/win rate|conversion|lost|pipeline.*funnel|close rate/)) {
    const byExec = [...new Set(pipeline.map(r=>r.AccountExec))];
    return {
      type: 'iris', title: `Pipeline Win Rate — ${Math.round((wonDeals.length/pipeline.length)*100)}% Overall`,
      summary: `Converted **${wonDeals.length} of ${pipeline.length}** deals. **${lostDeals.length}** lost — mainly Logistics and InsurTech.`,
      table: byExec.map(exec => { const all=pipeline.filter(r=>r.AccountExec===exec); const w=all.filter(r=>r.Stage==='Closed Won').length; return { AE: exec, Won: w, Lost: all.length-w, 'Win Rate': `${Math.round((w/all.length)*100)}%`, Revenue: fmt(wonDeals.filter(r=>r.AccountExec===exec).reduce((s,r)=>s+r.DealValue,0)) }; }),
      chart: { type: 'bar', label: 'Win / Loss by AE', labels: byExec.map(e=>e.split(' ')[0]), datasets: [{ label: 'Won', data: byExec.map(e=>wonDeals.filter(r=>r.AccountExec===e).length), color: '#059669' },{ label: 'Lost', data: byExec.map(e=>lostDeals.filter(r=>r.AccountExec===e).length), color: '#DC2626' }] },
      source: 'Pipeline → All stages → Win/Loss'
    };
  }

  if (q.match(/us|india|market|region|geo/)) {
    const rTotals = {};
    wonDeals.forEach(r => { rTotals[r.Region]=(rTotals[r.Region]||0)+r.DealValue; });
    const entries = Object.entries(rTotals);
    const total = entries.reduce((s,[,v])=>s+v,0);
    return {
      type: 'iris', title: 'Revenue Split — US vs India',
      summary: `US: **${Math.round((rTotals['US']/total)*100)}%** of revenue. India: **${Math.round((rTotals['India']/total)*100)}%**.`,
      table: entries.map(([region,rev]) => ({ Market: region, Revenue: fmt(rev), Share: `${Math.round((rev/total)*100)}%`, Deals: wonDeals.filter(r=>r.Region===region).length })),
      chart: { type: 'pie', label: 'Revenue by Market', labels: entries.map(([r])=>r+' Market'), datasets: [{ data: entries.map(([,v])=>+(v/100000).toFixed(1)), colors: ['#4F46E5','#059669'] }] },
      source: 'Pipeline → Closed Won → By Region'
    };
  }

  if (q.match(/lead source|inbound|outbound|referral|apollo|linkedin|channel/)) {
    const srcT = {};
    wonDeals.forEach(r => { srcT[r.LeadSource]=(srcT[r.LeadSource]||0)+r.DealValue; });
    const entries = Object.entries(srcT).sort((a,b)=>b[1]-a[1]);
    return {
      type: 'iris', title: 'Revenue by Lead Source',
      summary: `**${entries[0][0]}** is the highest-value channel at ${fmt(entries[0][1])}.`,
      table: entries.map(([src,rev]) => ({ 'Lead Source': src, Revenue: fmt(rev), Deals: wonDeals.filter(r=>r.LeadSource===src).length, 'Avg Deal': fmt(rev/wonDeals.filter(r=>r.LeadSource===src).length) })),
      chart: { type: 'pie', label: 'Revenue by Lead Source', labels: entries.map(([s])=>s), datasets: [{ data: entries.map(([,v])=>+(v/100000).toFixed(1)), colors: ['#4F46E5','#1D6FA4','#059669'] }] },
      source: 'Pipeline → Closed Won → By Lead Source'
    };
  }

  if (q.match(/prospect|lead|iris.*campaign|qualified|score|demo|outreach|pipeline.*status/)) {
    const stageCounts = {};
    prospects.forEach(p => { stageCounts[p.Stage]=(stageCounts[p.Stage]||0)+1; });
    const sorted = [...prospects].sort((a,b)=>b.Score-a.Score);
    return {
      type: 'iris', title: `IRIS Campaign — ${prospects.length} Active Prospects`,
      summary: `**${prospects.filter(p=>p.Score>=7.5).length} high-priority leads** (score ≥ 7.5). Top: **${sorted[0].Company}** (Score: ${sorted[0].Score}).`,
      table: sorted.map(p => ({ Company: p.Company, Vertical: p.Vertical, Country: p.Country, Stage: p.Stage, Score: p.Score })),
      chart: { type: 'bar', label: 'Prospects by Stage', labels: Object.keys(stageCounts), datasets: [{ label: 'Prospects', data: Object.values(stageCounts), color: '#4F46E5' }] },
      source: 'Prospects Sheet → Sorted by Score'
    };
  }

  if (q.match(/overview|dashboard|company|algonox|summary|how are we/)) {
    const totalRev = revenue.reduce((s,r)=>s+r.Revenue,0);
    const totalDeals = revenue.reduce((s,r)=>s+r.Deals,0);
    const quarters = [...new Set(revenue.map(r=>r.Quarter))];
    const qRev = quarters.map(q=>revenue.filter(r=>r.Quarter===q).reduce((s,r)=>s+r.Revenue,0));
    return {
      type: 'combined', title: 'Algonox Intelligence Overview — FY2025',
      iris: {
        label: 'Live Business Metrics',
        summary: `Revenue: **${fmt(totalRev)}** · **${totalDeals} deals** · Win rate: **${Math.round((wonDeals.length/pipeline.length)*100)}%** · **${prospects.length}** active prospects.`,
        table: [{ Metric: 'FY2025 Revenue', Value: fmt(totalRev) },{ Metric: 'Deals Closed', Value: totalDeals },{ Metric: 'Win Rate', Value: `${Math.round((wonDeals.length/pipeline.length)*100)}%` },{ Metric: 'Best Quarter', Value: `Q4 — ${fmt(qRev[3])}` },{ Metric: 'Active Prospects', Value: prospects.length }],
        chart: { type: 'line', label: 'Quarterly Revenue (₹L)', labels: quarters, datasets: [{ label: 'Revenue', data: qRev.map(v=>+(v/100000).toFixed(1)), color: '#4F46E5' }] }
      },
      memora: {
        label: 'Operational Memory',
        summary: `**${incidents.length} incidents** indexed · Avg: **${avg(incidents.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs** · **${incidents.filter(i=>i.Severity==='Critical').length} Critical** resolved · **${team.length}** experts.`,
        table: [{ Metric: 'Incidents Indexed', Value: incidents.length },{ Metric: 'Avg Resolution', Value: `${avg(incidents.map(i=>i.ResolutionTime_hrs)).toFixed(1)} hrs` },{ Metric: 'Critical Incidents', Value: incidents.filter(i=>i.Severity==='Critical').length },{ Metric: 'Experts Mapped', Value: team.length },{ Metric: 'Failure Patterns', Value: [...new Set(incidents.map(i=>i.Category))].length }]
      },
      source: 'All Sheets → IRIS + MEMORA combined'
    };
  }

  return {
    type: 'fallback', title: 'Try one of these',
    summary: 'MEMORA and IRIS are ready. Here are some questions to get started:',
    suggestions: [
      'Agent task queue deadlock during parallel execution',
      'Outbound call drops on US toll-free lines',
      'Salesforce OAuth connector not working',
      'NLP intent classification accuracy drop after model update',
      'Enterprise client locked out of API',
      'IVR stuck in Hindi language fallback loop',
      'Knowledge graph query timing out on large dataset',
      'TRAI DND registry sync failure',
      'SAP approval callback looping',
      'Audio garbled on BFSI client calls',
      'Show revenue by product',
      'Who is the top account executive?',
    ]
  };
}

app.post('/api/query', (req, res) => {
  const { question } = req.body;
  if (!question || question.trim().length < 3) return res.json({ error: 'Please type a question.' });
  try {
    const data = loadData();
    const result = queryEngine(question, data);
    setTimeout(() => res.json(result), 400 + Math.random() * 300);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Error: ' + err.message });
  }
});

app.get('/api/data', (req, res) => {
  try { res.json(loadData()); } catch(e) { res.status(500).json({ error: e.message }); }
});

const PORT = 3000;
app.listen(PORT, () => console.log(`Algonox Intelligence Platform → http://localhost:${PORT}`));