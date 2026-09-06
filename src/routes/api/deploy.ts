'use strict';

const express = require('express');
const router = express.Router();
const { requireAuth } = require('../../middleware/auth.ts');
const g = require('../../services/graphService.ts');

type Request = import('express').Request
type Response = import('express').Response

const token = (req: Request) => req.session.auth?.accessToken || '';

// GET /api/deploy/status
router.get('/status', requireAuth, async (req: Request, res: Response) => {
  try {
    const deployments = await g.getOfficeDeployments(token(req));
    if (deployments === null) {
      return res.json({ supported: false, deployments: [], message: 'Office add-in deployment API not available for this tenant.' });
    }
    res.json({ supported: true, deployments });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// POST /api/deploy/prepare
router.post('/prepare', requireAuth, async (req: Request, res: Response) => {
  const { product, source, sourceValue, targets, method } = req.body;
  if ((method || 'prepare') !== 'integrated' && !targets?.length) {
    return res.status(400).json({ error: 'At least one target is required' });
  }
  if (!sourceValue) return res.status(400).json({ error: 'Source URL or manifest path is required' });

  // Log the deployment request (actual Graph deployment API requires specific admin roles)
  const summary = {
    product: product || 'Outlook',
    source: source || 'appsource',
    sourceValue,
    method: method || 'prepare',
    targetCount: (targets || []).length,
    targets: (targets || []).slice(0, 20),
    preparedAt: new Date().toISOString(),
  };

  res.json({
    success: true,
    message: `Deployment prepared for ${targets.length} target(s). Use M365 Admin Center Integrated Apps to complete deployment.`,
    summary,
    m365Url: 'https://admin.microsoft.com/adminportal/home#/Settings/IntegratedApps',
  });
});

// POST /api/deploy/export-csv — export targets list as CSV
router.post('/export-csv', requireAuth, (req: Request, res: Response) => {
  const { targets = [] } = req.body;
  if (!targets.length) return res.status(400).json({ error: 'No targets to export' });
  const lines = ['Name,Email/ID,Type', ...targets.map(t =>
    `"${(t.name || '').replace(/"/g, '""')}","${(t.emailOrId || t.id || '').replace(/"/g, '""')}","${(t.type || '').replace(/"/g, '""')}"`
  )];
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', 'attachment; filename="deploy_targets.csv"');
  res.send(lines.join('\n'));
});

module.exports = router;
