const express = require('express');
const cors = require('cors');
const path = require('path');
const dotenv = require('dotenv');
const { ComprehendClient, DetectKeyPhrasesCommand } = require('@aws-sdk/client-comprehend');

dotenv.config();

const app = express();
const port = Number(process.env.PORT || 3001);
const awsRegion = process.env.AWS_REGION || 'us-east-1';

const comprehend = new ComprehendClient({ region: awsRegion });

app.use(cors());
app.use(express.json({ limit: '1mb' }));
app.use(express.static(__dirname));

app.get('/health', (_req, res) => {
  res.json({ ok: true, service: 'rework-dashboard-ai', region: awsRegion });
});

app.post('/api/ai-insights', async (req, res) => {
  try {
    const body = req.body || {};
    const summaryText = String(body.summaryText || '').trim();
    const metrics = body.metrics || {};

    if (!summaryText) {
      return res.status(400).json({ error: 'summaryText is required' });
    }

    const clippedSummary = summaryText.slice(0, 4500);

    let keyPhrases = [];
    try {
      const command = new DetectKeyPhrasesCommand({
        LanguageCode: 'en',
        Text: clippedSummary
      });
      const response = await comprehend.send(command);
      keyPhrases = (response.KeyPhrases || [])
        .filter((item) => Number(item.Score || 0) >= 0.8)
        .sort((a, b) => Number(b.Score || 0) - Number(a.Score || 0))
        .slice(0, 5)
        .map((item) => item.Text)
        .filter(Boolean);
    } catch (err) {
      console.warn('Comprehend detect key phrases failed:', err.message || err);
    }

    const insights = buildActionableInsights(metrics, keyPhrases);

    res.json({
      generatedAt: new Date().toISOString(),
      keyPhrases,
      insights
    });
  } catch (error) {
    console.error('AI insights error:', error);
    res.status(500).json({ error: 'Unable to generate insights right now.' });
  }
});

app.get('/', (_req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(port, () => {
  console.log(`AI backend running at http://localhost:${port}`);
});

function buildActionableInsights(metrics, keyPhrases) {
  const insights = [];
  const totalHoldUnits = Number(metrics.totalHoldUnits || 0);
  const totalItemsReworked = Number(metrics.totalItemsReworked || 0);
  const reworkPercent = Number(metrics.reworkPercent || 0);
  const percentScrapped = Number(metrics.percentScrapped || 0);
  const topRootCause = String(metrics.topRootCause || 'Unknown');
  const topLocation = String(metrics.topLocation || 'Unknown');
  const topDisposition = String(metrics.topDisposition || 'Unknown');

  // Critical threshold actions
  if (reworkPercent >= 5) {
    insights.push(`🚨 STOP PRODUCTION: Rework at ${reworkPercent.toFixed(1)}% exceeds critical limit. PLANT MANAGER must halt ${topLocation} line immediately. Convene emergency response team within 1 hour. Root cause: ${topRootCause}. Do not resume until containment verified by Quality Director.`);
    insights.push(`📋 IMMEDIATE ACTIONS (Next 4 Hours): (1) Quality Engineer completes failure mode analysis on last 50 units, (2) Production Supervisor segregates all WIP for inspection, (3) Maintenance verifies equipment calibration, (4) Document findings in incident report #[DATE]-${topLocation}.`);
  } else if (reworkPercent >= 3) {
    insights.push(`⚠️ ESCALATE NOW: Rework at ${reworkPercent.toFixed(1)}% requires immediate intervention. OPERATIONS MANAGER assigns dedicated Quality Lead to ${topLocation} for next 48 hours. Implement 100% inspection on ${topRootCause} until 3 consecutive clean batches. Report hourly to Plant Manager.`);
    insights.push(`🔍 CONTAINMENT PLAN (24 Hours): Deploy 3-person audit team to ${topLocation}. Check: (1) Operator training records, (2) Equipment maintenance logs, (3) Material lot traceability. Quality Supervisor presents findings at tomorrow's 7am production meeting with corrective actions.`);
  } else if (reworkPercent >= 1.5) {
    insights.push(`⚡ PREVENTIVE ACTION REQUIRED: Rework at ${reworkPercent.toFixed(1)}% trending toward threshold. LINE SUPERVISOR at ${topLocation} must implement layered process audits every 2 hours focusing on ${topRootCause}. Quality Technician documents all checks in LPA tracker. Escalate if any audit fails.`);
  } else {
    insights.push(`✓ SUSTAIN PERFORMANCE: Rework controlled at ${reworkPercent.toFixed(1)}%. SHIFT LEAD continues standard work audits at ${topLocation}. Monitor ${topRootCause} daily. If any single shift exceeds 1%, trigger immediate supervisor review and document in shift handoff log.`);
  }

  // Dumped-specific actions
  if (percentScrapped >= 2) {
    insights.push(`💰 COST RECOVERY MANDATE: Dumped at ${percentScrapped.toFixed(1)}% = irreversible loss. QUALITY MANAGER implements mandatory pre-disposition review gate at ${topLocation} starting next shift. No material dumped without dual sign-off (Quality + Production). Target: <1% dumped within 14 days. Track daily.`);
  } else if (percentScrapped >= 1) {
    insights.push(`💡 DUMP REDUCTION: At ${percentScrapped.toFixed(1)}%, implement enhanced verification. QUALITY TECHNICIAN adds final checkpoint before any dump decision. Use decision tree posted at quality station. Supervisor reviews all dump decisions weekly.`);
  }

  // Root cause ownership
  if (totalItemsReworked > 100) {
    insights.push(`🎯 ROOT CAUSE CLOSURE: ${topRootCause} drove ${totalItemsReworked.toLocaleString()} rework cases. QUALITY ENGINEER assigned as owner - complete 5-Why analysis by Friday 5pm, present to leadership Monday 8am. Include: (1) Failure timeline, (2) Contributing factors, (3) Permanent corrective actions with owners/dates, (4) Verification plan.`);
  } else {
    insights.push(`🎯 ROOT CAUSE MONITORING: ${topRootCause} is primary driver (${totalItemsReworked.toLocaleString()} cases). QUALITY TECHNICIAN tracks daily occurrences. If count increases 20% week-over-week, escalate to Quality Engineer for formal investigation.`);
  }

  // Resource and capacity planning
  if (totalHoldUnits > 0 && totalItemsReworked > 0) {
    const ratio = (totalItemsReworked / totalHoldUnits) * 100;
    if (ratio > 10) {
      insights.push(`📊 CAPACITY CRISIS: ${ratio.toFixed(1)}% rework rate (${totalItemsReworked.toLocaleString()}/${totalHoldUnits.toLocaleString()} units) requires immediate staffing. PRODUCTION SCHEDULER allocates 4-6 FTEs to dedicated rework cell by tomorrow. Prioritize by age (oldest holds first). Update capacity plan daily at 2pm production meeting.`);
    } else {
      insights.push(`📊 RESOURCE ALLOCATION: ${totalItemsReworked.toLocaleString()} of ${totalHoldUnits.toLocaleString()} units need rework (${ratio.toFixed(1)}%). PRODUCTION SCHEDULER assigns 2-3 FTEs to rework during peak hours (6am-2pm). Balance with production schedule. Review staffing needs every Monday.`);
    }
  }

  // Process standardization
  if (topDisposition && topDisposition !== 'Unknown') {
    insights.push(`📋 STANDARDIZE DECISIONS: ${topDisposition} is most common disposition but variation exists. QUALITY MANAGER creates visual decision matrix (flowchart) by Wednesday. Train all supervisors Thursday-Friday. Post at each quality station. Audit compliance in 2 weeks - target 95% adherence.`);
  }

  return insights.slice(0, 6);
}
