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

  if (reworkPercent >= 5) {
    insights.push(`Rework is elevated at ${reworkPercent.toFixed(1)}%. Launch a 48-hour containment plan at ${topLocation} focused on ${topRootCause}.`);
  } else if (reworkPercent >= 3) {
    insights.push(`Rework is trending above target at ${reworkPercent.toFixed(1)}%. Run layered process checks on shifts with the highest ${topRootCause} volume.`);
  } else {
    insights.push(`Rework is currently controlled at ${reworkPercent.toFixed(1)}%. Maintain standard checks and monitor for drift in ${topRootCause}.`);
  }

  if (percentScrapped >= 2) {
    insights.push(`Scrap is ${percentScrapped.toFixed(1)}%. Prioritize first-pass quality checks before disposition to reduce irreversible loss.`);
  }

  if (totalHoldUnits > 0 && totalItemsReworked > 0) {
    const ratio = (totalItemsReworked / totalHoldUnits) * 100;
    insights.push(`Volume context: ${totalItemsReworked.toLocaleString()} reworked out of ${totalHoldUnits.toLocaleString()} hold units (${ratio.toFixed(1)}%). Allocate staffing to high-risk windows.`);
  }

  insights.push(`Primary driver is ${topRootCause}. Assign one owner to complete root-cause verification and corrective action closure this week.`);

  if (topDisposition && topDisposition !== 'Unknown') {
    insights.push(`Most common disposition is ${topDisposition}. Standardize decision criteria to reduce variation between teams.`);
  }

  if (keyPhrases.length > 0) {
    insights.push(`Comprehend key themes: ${keyPhrases.join(', ')}. Use these to refine work-center huddles and daily action boards.`);
  }

  return insights.slice(0, 6);
}
