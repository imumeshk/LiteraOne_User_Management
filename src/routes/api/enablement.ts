'use strict';

const express = require('express');
const path = require('path');
const fs = require('fs');

const router = express.Router();
type Request = import('express').Request
type Response = import('express').Response
type NextFunction = import('express').NextFunction

const DEFAULT_CONFIG = {
  resources: [
    {
      id: 'vid01',
      title: 'Getting Started with Litera One',
      type: 'Video',
      icon: '▶',
      url: 'https://www.litera.com/products/litera-one/',
      description: 'A quick overview of the Litera One platform and its capabilities.',
      showCopyLink: true,
    },
    {
      id: 'guide01',
      title: 'Litera Check User Guide',
      type: 'Guide',
      icon: '📄',
      url: 'https://www.litera.com/products/foundation/litera-check/',
      description: 'In-depth guide for using Litera Check for proofreading and document analysis.',
      showOpenButton: true,
      showCopyLink: false,
    },
    {
      id: 'article01',
      title: 'Understanding SCIM Sync',
      type: 'Article',
      icon: '📰',
      url: 'https://www.litera.com/company/news/',
      description: 'Learn how SCIM provisioning works and how to configure it with Litera.',
      tags: ['SCIM', 'Provisioning'],
    },
    {
      id: 'guide02',
      title: 'Metadact User Guide',
      type: 'Guide',
      icon: '📄',
      url: 'https://www.litera.com/products/workflow/metadact/',
      description: 'Comprehensive guide to cleaning metadata from your documents.',
      tags: ['Metadata', 'Security'],
    },
    {
      id: 'guide03',
      title: 'Content Companion User Guide',
      type: 'Guide',
      icon: '📄',
      url: 'https://www.litera.com/products/workflow/content-companion/',
      description: 'A guide to using Content Companion for document creation and review.',
      tags: ['Content', 'Productivity'],
      showOpenButton: true,
    },
  ],
  exportButtons: [
    { id: 'pack-getting-started', text: "Export 'Getting Started' Pack", resourceIds: ['vid01', 'guide01'] },
    { id: 'pack-all-guides', text: 'Export All Guides', resourceType: 'Guide' },
    { id: 'pack-scim', text: 'Export SCIM Resources', resourceTag: 'SCIM' },
    { id: 'pack-productivity', text: "Export 'Productivity' Pack", resourceIds: ['vid01', 'guide03'] },
  ],
  exportSettings: {
    pageTitle: 'Litera Enablement Hub',
    headerText: 'Your Custom Enablement Resources',
  },
  viewSettings: {
    cardsPerRow: 3,
    showOpenButtonDefault: true,
    showCopyLinkDefault: true,
  },
};

function configPath() {
  const custom = String(process.env.ENABLEMENT_HUB_CONFIG_PATH || '').trim();
  if (!custom) return path.join(process.cwd(), 'enablement-hub.json');
  return path.isAbsolute(custom) ? custom : path.join(process.cwd(), custom);
}

function ensureArray(value) {
  return Array.isArray(value) ? value : [];
}

function normalizeResource(resource) {
  const id = String(resource?.id || '').trim();
  if (!id) throw new Error('Resource id is required.');
  const title = String(resource?.title || '').trim();
  if (!title) throw new Error(`Resource "${id}" title is required.`);
  const type = String(resource?.type || '').trim();
  if (!type) throw new Error(`Resource "${id}" type is required.`);
  const url = String(resource?.url || '').trim();
  if (!url) throw new Error(`Resource "${id}" url is required.`);
  return {
    id,
    title,
    type,
    icon: String(resource?.icon || '').trim() || '📄',
    url,
    description: String(resource?.description || '').trim(),
    tags: ensureArray(resource?.tags).map(t => String(t).trim()).filter(Boolean),
    showOpenButton: resource?.showOpenButton === undefined ? undefined : !!resource.showOpenButton,
    showCopyLink: resource?.showCopyLink === undefined ? undefined : !!resource.showCopyLink,
  };
}

function normalizeExportButton(button, index = 0) {
  const text = String(button?.text || '').trim();
  if (!text) throw new Error('Export button text is required.');
  const id = String(button?.id || `export-btn-${index + 1}`).trim();
  return {
    id,
    text,
    resourceIds: ensureArray(button?.resourceIds).map(v => String(v).trim()).filter(Boolean),
    resourceType: String(button?.resourceType || '').trim() || undefined,
    resourceTag: String(button?.resourceTag || '').trim() || undefined,
  };
}

function normalizeConfig(raw) {
  const resources = ensureArray(raw?.resources).map(normalizeResource);
  const seen = new Set();
  for (const r of resources) {
    if (seen.has(r.id)) throw new Error(`Duplicate resource id "${r.id}".`);
    seen.add(r.id);
  }
  const exportButtons = ensureArray(raw?.exportButtons).map((b, i) => normalizeExportButton(b, i));
  return {
    resources,
    exportButtons,
    exportSettings: {
      pageTitle: String(raw?.exportSettings?.pageTitle || DEFAULT_CONFIG.exportSettings.pageTitle).trim(),
      headerText: String(raw?.exportSettings?.headerText || DEFAULT_CONFIG.exportSettings.headerText).trim(),
    },
    viewSettings: {
      cardsPerRow: Number.parseInt(raw?.viewSettings?.cardsPerRow, 10) || DEFAULT_CONFIG.viewSettings.cardsPerRow,
      showOpenButtonDefault: raw?.viewSettings?.showOpenButtonDefault !== undefined
        ? !!raw.viewSettings.showOpenButtonDefault
        : DEFAULT_CONFIG.viewSettings.showOpenButtonDefault,
      showCopyLinkDefault: raw?.viewSettings?.showCopyLinkDefault !== undefined
        ? !!raw.viewSettings.showCopyLinkDefault
        : DEFAULT_CONFIG.viewSettings.showCopyLinkDefault,
    },
  };
}

function readConfig() {
  const file = configPath();
  if (!fs.existsSync(file)) return normalizeConfig(DEFAULT_CONFIG);
  const parsed = JSON.parse(fs.readFileSync(file, 'utf8'));
  return normalizeConfig(parsed);
}

function writeConfig(next) {
  const file = configPath();
  const dir = path.dirname(file);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(file, JSON.stringify(next, null, 2), 'utf8');
}

function adminConfigured() {
  const username = String(process.env.ENABLEMENT_ADMIN_USERNAME || '').trim();
  const password = String(process.env.ENABLEMENT_ADMIN_PASSWORD || '').trim();
  return { username, password, ok: !!(username && password) };
}

function requireEnablementAdmin(req: Request, res: Response, next: NextFunction) {
  if (req.session?.enablementAdmin?.loggedIn) return next();
  return res.status(401).json({ error: 'Enablement admin authentication required.' });
}

function matchesType(value, type) {
  return String(value || '').toLowerCase() === String(type || '').toLowerCase();
}

function hasTag(resource, tag) {
  const tags = ensureArray(resource?.tags).map(t => String(t).toLowerCase());
  return tags.includes(String(tag || '').toLowerCase());
}

function selectByRule(resources, rule = {}) {
  if (ensureArray(rule.resourceIds).length) {
    const allowed = new Set(rule.resourceIds.map(v => String(v)));
    return resources.filter(r => allowed.has(String(r.id)));
  }
  if (rule.resourceType) return resources.filter(r => matchesType(r.type, rule.resourceType));
  if (rule.resourceTag) return resources.filter(r => hasTag(r, rule.resourceTag));
  return resources;
}

// ─── Admin auth ───────────────────────────────────────────────────────────────

router.get('/admin/status', (req: Request, res: Response) => {
  const cfg = adminConfigured();
  res.json({
    configured: cfg.ok,
    loggedIn: !!req.session?.enablementAdmin?.loggedIn,
    username: req.session?.enablementAdmin?.username || null,
  });
});

router.post('/admin/login', (req: Request, res: Response) => {
  const cfg = adminConfigured();
  if (!cfg.ok) {
    return res.status(500).json({ error: 'Enablement admin is not configured. Set ENABLEMENT_ADMIN_USERNAME and ENABLEMENT_ADMIN_PASSWORD.' });
  }
  const username = String(req.body?.username || '').trim();
  const password = String(req.body?.password || '').trim();
  if (username !== cfg.username || password !== cfg.password) {
    return res.status(401).json({ error: 'Invalid enablement admin credentials.' });
  }
  req.session.enablementAdmin = { loggedIn: true, username, loggedInAt: new Date().toISOString() };
  return res.json({ success: true, username });
});

router.post('/admin/logout', (req: Request, res: Response) => {
  if (req.session) req.session.enablementAdmin = null;
  res.json({ success: true });
});

// ─── Admin config management ─────────────────────────────────────────────────

router.get('/admin/config', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    res.json({ config: readConfig() });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

router.put('/admin/config', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    const normalized = normalizeConfig(req.body || {});
    writeConfig(normalized);
    res.json({ success: true, config: normalized });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

router.post('/admin/resources', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    const cfg = readConfig();
    const resource = normalizeResource(req.body || {});
    if (cfg.resources.some(r => r.id === resource.id)) {
      return res.status(409).json({ error: `Resource id "${resource.id}" already exists.` });
    }
    cfg.resources.push(resource);
    writeConfig(cfg);
    res.json({ success: true, resource });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

router.put('/admin/resources/:id', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    const id = String(req.params.id || '').trim();
    const cfg = readConfig();
    const idx = cfg.resources.findIndex(r => r.id === id);
    if (idx < 0) return res.status(404).json({ error: 'Resource not found.' });
    const updated = normalizeResource({ ...cfg.resources[idx], ...req.body, id });
    cfg.resources[idx] = updated;
    writeConfig(cfg);
    res.json({ success: true, resource: updated });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

router.delete('/admin/resources/:id', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    const id = String(req.params.id || '').trim();
    const cfg = readConfig();
    const before = cfg.resources.length;
    cfg.resources = cfg.resources.filter(r => r.id !== id);
    if (cfg.resources.length === before) return res.status(404).json({ error: 'Resource not found.' });
    writeConfig(cfg);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

router.post('/admin/export-buttons', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    const cfg = readConfig();
    const button = normalizeExportButton(req.body || {}, cfg.exportButtons.length);
    if (cfg.exportButtons.some(b => b.id === button.id)) {
      return res.status(409).json({ error: `Export button id "${button.id}" already exists.` });
    }
    cfg.exportButtons.push(button);
    writeConfig(cfg);
    res.json({ success: true, exportButton: button });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

router.put('/admin/export-buttons/:id', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    const id = String(req.params.id || '').trim();
    const cfg = readConfig();
    const idx = cfg.exportButtons.findIndex(b => b.id === id);
    if (idx < 0) return res.status(404).json({ error: 'Export button not found.' });
    const updated = normalizeExportButton({ ...cfg.exportButtons[idx], ...req.body, id }, idx);
    cfg.exportButtons[idx] = updated;
    writeConfig(cfg);
    res.json({ success: true, exportButton: updated });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

router.delete('/admin/export-buttons/:id', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    const id = String(req.params.id || '').trim();
    const cfg = readConfig();
    const before = cfg.exportButtons.length;
    cfg.exportButtons = cfg.exportButtons.filter(b => b.id !== id);
    if (cfg.exportButtons.length === before) return res.status(404).json({ error: 'Export button not found.' });
    writeConfig(cfg);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

router.patch('/admin/settings', requireEnablementAdmin, (req: Request, res: Response) => {
  try {
    const cfg = readConfig();
    const next = {
      ...cfg,
      exportSettings: {
        ...cfg.exportSettings,
        ...(req.body?.exportSettings || {}),
      },
      viewSettings: {
        ...cfg.viewSettings,
        ...(req.body?.viewSettings || {}),
      },
    };
    const normalized = normalizeConfig(next);
    writeConfig(normalized);
    res.json({ success: true, config: normalized });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// ─── Public read/export APIs ────────────────────────────────────────────────

router.get('/resources', (req: Request, res: Response) => {
  try {
    const cfg = readConfig();
    const { type, q, tag } = req.query;
    let filtered = cfg.resources;
    if (type && type !== 'All') filtered = filtered.filter(r => matchesType(r.type, type));
    if (tag) filtered = filtered.filter(r => hasTag(r, tag));
    if (q) {
      const ql = String(q).toLowerCase();
      filtered = filtered.filter(r =>
        (r.title || '').toLowerCase().includes(ql) ||
        (r.description || '').toLowerCase().includes(ql) ||
        ensureArray(r.tags).some(t => String(t).toLowerCase().includes(ql))
      );
    }
    res.json({
      resources: filtered,
      exportButtons: cfg.exportButtons,
      exportSettings: cfg.exportSettings,
      viewSettings: cfg.viewSettings,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

router.get('/types', (req: Request, res: Response) => {
  try {
    const cfg = readConfig();
    const types = [...new Set(cfg.resources.map(r => r.type).filter(Boolean))];
    res.json({ types });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

router.get('/tags', (req: Request, res: Response) => {
  try {
    const cfg = readConfig();
    const tags = [...new Set(cfg.resources.flatMap(r => ensureArray(r.tags).map(t => String(t))))];
    res.json({ tags });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

router.post('/export', (req: Request, res: Response) => {
  try {
    const cfg = readConfig();
    const { ids, exportButtonId, layout = 'cards', title } = req.body || {};

    let selected = cfg.resources;
    if (exportButtonId) {
      const btn = cfg.exportButtons.find(b => b.id === exportButtonId);
      if (!btn) return res.status(404).json({ error: 'Export button not found.' });
      selected = selectByRule(cfg.resources, btn);
    } else if (ensureArray(ids).length) {
      const allowed = new Set(ids.map(v => String(v)));
      selected = cfg.resources.filter(r => allowed.has(String(r.id)));
    }
    if (!selected.length) return res.status(400).json({ error: 'No resources to export' });

    const viewDefaults = cfg.viewSettings || {};
    const pageTitle = String(title || cfg.exportSettings?.pageTitle || 'Litera Enablement Hub');
    const headerText = String(cfg.exportSettings?.headerText || pageTitle);
    const typeColors = {
      video: '#fee2e2:#b91c1c',
      guide: '#e0f2fe:#0369a1',
      article: '#ecfeff:#0e7490',
      doc: '#dcfce7:#15803d',
      template: '#ede9fe:#7c3aed',
    };

    const cards = selected.map(r => {
      const typeKey = String(r.type || '').toLowerCase();
      const [bg, fg] = (typeColors[typeKey] || '#f1f5f9:#475569').split(':');
      const showOpen = r.showOpenButton === undefined ? !!viewDefaults.showOpenButtonDefault : !!r.showOpenButton;
      const openBtn = showOpen
        ? `<a href="${r.url}" target="_blank" style="display:inline-block;margin-top:14px;padding:7px 14px;background:#0284c7;color:#fff;border-radius:7px;font-size:12px;font-weight:600;text-decoration:none">↗ Open Resource</a>`
        : '';
      return layout === 'list'
        ? `<li style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin-bottom:12px;display:flex;align-items:center;gap:16px">
            <span style="font-size:24px">${r.icon || '📄'}</span>
            <div style="flex:1">
              <a href="${r.url}" target="_blank" style="font-weight:700;font-size:15px;color:#0284c7;text-decoration:none">${r.title}</a>
              <p style="margin:4px 0 0;color:#64748b;font-size:13px">${r.description || ''}</p>
            </div>
            <span style="background:${bg};color:${fg};padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;text-transform:uppercase;white-space:nowrap">${r.type}</span>
          </li>`
        : `<div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:18px;display:flex;flex-direction:column">
            <span style="background:${bg};color:${fg};padding:2px 9px;border-radius:20px;font-size:10px;font-weight:700;text-transform:uppercase;display:inline-block;margin-bottom:10px;align-self:flex-start">${r.type}</span>
            <div style="font-size:16px;font-weight:700;margin-bottom:8px">${r.icon || ''} ${r.title}</div>
            <div style="font-size:13px;color:#64748b;flex:1;line-height:1.5">${r.description || ''}</div>
            ${openBtn}
          </div>`;
    }).join('\n');

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>${pageTitle}</title>
<style>
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;max-width:960px;margin:2rem auto;padding:0 1.5rem;color:#0f172a;background:#f8fafc}
  h1{font-size:24px;font-weight:800;margin-bottom:.4rem;color:#0f172a}
  p.sub{color:#64748b;margin-bottom:1.5rem}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:16px}
  ul{list-style:none;padding:0;margin:0}
  @media(max-width:600px){.grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<h1>${pageTitle}</h1>
<p class="sub">${headerText}</p>
${layout === 'list' ? `<ul>${cards}</ul>` : `<div class="grid">${cards}</div>`}
<p style="margin-top:2rem;color:#94a3b8;font-size:12px">Exported from Litera One Manager Portal — ${new Date().toLocaleDateString()}</p>
</body>
</html>`;

    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="litera-enablement-${Date.now()}.html"`);
    res.send(html);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

module.exports = router;
