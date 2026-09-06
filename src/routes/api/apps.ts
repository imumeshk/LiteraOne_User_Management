'use strict';

const express = require('express');
const router = express.Router();
const { requireAuth } = require('../../middleware/auth.ts');
const g = require('../../services/graphService.ts');

type Request = import('express').Request
type Response = import('express').Response

const token = (req: Request) => req.session.auth?.accessToken || '';

// GET /api/apps/search?q=Litera
router.get('/search', requireAuth, async (req: Request, res: Response) => {
  try {
    const result = await g.searchServicePrincipals(token(req), req.query.q || 'Litera');
    res.json({ apps: result?.value || [] });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// GET /api/apps/:spId
router.get('/:spId', requireAuth, async (req: Request, res: Response) => {
  try {
    const [sp, grants, principalRoleAssignments] = await Promise.allSettled([
      g.getServicePrincipal(token(req), req.params.spId),
      g.getOAuth2Grants(token(req), req.params.spId),
      g.getPrincipalAppRoleAssignments(token(req), req.params.spId),
    ]);

    const spData = sp.status === 'fulfilled' ? sp.value : null;
    if (!spData) return res.status(404).json({ error: 'Service principal not found' });

    // Get permissions via appId
    let perms = [];
    try { perms = await g.getAppPermissions(token(req), spData.appId); } catch (_) {}
    if (!perms.length && grants.status === 'fulfilled' && Array.isArray(grants.value?.value)) {
      const scopeNames = new Set();
      for (const grant of grants.value.value) {
        const scopes = String(grant.scope || '').split(' ').map(s => s.trim()).filter(Boolean);
        scopes.forEach(s => scopeNames.add(s));
      }
      perms = [...scopeNames].map(name => ({ name, type: 'Scope (Granted)' }));
    }
    if (!perms.length && Array.isArray(spData.appRoles)) {
      perms = spData.appRoles
        .filter(r => r?.isEnabled)
        .map(r => ({ name: r.value || r.displayName || r.id, type: 'App Role' }));
    }

    let requiresConsent = false;
    try {
      const appRes = await g.getApplicationByAppId(token(req), spData.appId);
      const appObj = appRes?.value?.[0];
      const requiredResourceAccess = appObj?.requiredResourceAccess || [];
      requiresConsent = requiredResourceAccess.some(rra => (rra?.resourceAccess || []).length > 0);
    } catch (_) {
      // Non-fatal: if app registration can't be read, fallback to grants/assignments only.
    }

    const delegatedGrantCount =
      grants.status === 'fulfilled' && Array.isArray(grants.value?.value)
        ? grants.value.value.length
        : 0;
    const appRoleGrantCount =
      principalRoleAssignments.status === 'fulfilled' && Array.isArray(principalRoleAssignments.value?.value)
        ? principalRoleAssignments.value.value.length
        : 0;

    // Consent is considered granted if:
    // 1) no additional consent is required by app registration, or
    // 2) delegated grants exist, or
    // 3) app-role assignments exist for the service principal as principal.
    const hasConsent = !requiresConsent || delegatedGrantCount > 0 || appRoleGrantCount > 0;
    const tags = spData?.tags || [];
    const isHidden = tags.includes('HideApp');

    res.json({ sp: spData, permissions: perms, hasConsent, isHidden, tags });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// GET /api/apps/:spId/assignments
router.get('/:spId/assignments', requireAuth, async (req: Request, res: Response) => {
  try {
    const raw = await g.getAppRoleAssignments(token(req), req.params.spId);
    const enriched = await g.enrichAssignments(token(req), raw);
    res.json({ assignments: enriched });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// POST /api/apps/:spId/assignments
router.post('/:spId/assignments', requireAuth, async (req: Request, res: Response) => {
  const { principalId, principalType } = req.body;
  if (!principalId) return res.status(400).json({ error: 'principalId required' });
  try {
    const result = await g.assignPrincipal(token(req), req.params.spId, principalId, principalType || 'User');
    res.json({ assignment: result });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// DELETE /api/apps/:spId/assignments/:assignmentId
router.delete('/:spId/assignments/:assignmentId', requireAuth, async (req: Request, res: Response) => {
  try {
    await g.removeAppRoleAssignment(token(req), req.params.spId, req.params.assignmentId);
    res.json({ success: true });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// POST /api/apps/:spId/consent
router.post('/:spId/consent', requireAuth, async (req: Request, res: Response) => {
  try {
    const sp = await g.getServicePrincipal(token(req), req.params.spId);
    if (!sp) return res.status(404).json({ error: 'Service principal not found' });
    const results = await g.grantAdminConsent(token(req), sp);
    res.json({ success: true, results });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// PATCH /api/apps/:spId/visibility
router.patch('/:spId/visibility', requireAuth, async (req: Request, res: Response) => {
  const { hide, appId } = req.body;
  try {
    const tags = await g.setVisibility(token(req), req.params.spId, appId, !!hide);
    res.json({ success: true, tags, isHidden: tags.includes('HideApp') });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// DELETE /api/apps/:spId
router.delete('/:spId', requireAuth, async (req: Request, res: Response) => {
  try {
    // Capture appId before deleting the service principal.
    const existingSp = await g.getServicePrincipal(token(req), req.params.spId).catch(() => null);
    const existingAppId = existingSp?.appId || null;

    await g.deleteServicePrincipal(token(req), req.params.spId);

    // Try to also delete the app registration
    try {
      if (existingAppId) {
        const appRes = await g.getApplicationByAppId(token(req), existingAppId);
        const appObjId = appRes?.value?.[0]?.id;
        if (appObjId) await g.deleteApplication(token(req), appObjId);
      }
    } catch (_) {}

    res.json({ success: true });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// GET /api/apps/:spId/sync-jobs
router.get('/:spId/sync-jobs', requireAuth, async (req: Request, res: Response) => {
  try {
    const jobs = await g.getSyncJobs(token(req), req.params.spId);
    res.json({ jobs });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// POST /api/apps/:spId/sync-jobs/:jobId/restart
router.post('/:spId/sync-jobs/:jobId/restart', requireAuth, async (req: Request, res: Response) => {
  try {
    await g.restartSyncJob(token(req), req.params.spId, req.params.jobId);
    res.json({ success: true });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// GET /api/apps/:spId/group-members/:groupId
router.get('/:spId/group-members/:groupId', requireAuth, async (req: Request, res: Response) => {
  try {
    const members = await g.getGroupMembers(token(req), req.params.groupId);
    res.json({ members: members.filter(m => m['@odata.type'] === '#microsoft.graph.user') });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// GET /api/apps/user-assignments/:userId
router.get('/user-assignments/:userId', requireAuth, async (req: Request, res: Response) => {
  try {
    const all = await g.getUserAppAssignments(token(req), req.params.userId);
    res.json({ assignments: all });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

// GET /api/apps/:spId/export-csv
router.get('/:spId/export-csv', requireAuth, async (req: Request, res: Response) => {
  try {
    const raw = await g.getAppRoleAssignments(token(req), req.params.spId);
    const enriched = await g.enrichAssignments(token(req), raw);
    const lines = ['Name,Email/ID,Type', ...enriched.map(a =>
      `"${(a.principalDisplayName || '').replace(/"/g, '""')}","${(a.emailOrId || '').replace(/"/g, '""')}","${a.principalType || ''}"`
    )];
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="assignments.csv"');
    res.send(lines.join('\n'));
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

module.exports = router;
