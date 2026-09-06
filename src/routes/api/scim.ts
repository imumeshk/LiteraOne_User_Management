'use strict';

const express = require('express');
const router = express.Router();
const { requireAuth } = require('../../middleware/auth.ts');
const g = require('../../services/graphService.ts');

type Request = import('express').Request
type Response = import('express').Response

const token = (req: Request) => req.session.auth?.accessToken || '';

/**
 * POST /api/scim/create
 * Creates a SCIM app. Streams progress via Server-Sent Events.
 * Accepts JSON body: { appName, scimUrl, scimToken, principals: [{id, name, type}], probeSpId }
 */
router.post('/create', requireAuth, async (req: Request, res: Response) => {
  const { appName, scimUrl, scimToken, principals = [], probeSpId } = req.body;

  if (!appName) {
    return res.status(400).json({ error: 'appName is required' });
  }
  if (!String(scimUrl || '').trim()) {
    return res.status(400).json({ error: 'scimUrl is required' });
  }
  if (!String(scimToken || '').trim()) {
    return res.status(400).json({ error: 'scimToken is required' });
  }

  // Set up SSE
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('X-Accel-Buffering', 'no');
  res.flushHeaders();

  const send = (data) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  try {
    const gen = g.createScimApp(token(req), { appName, scimUrl, scimToken, principals, probeSpId });
    for await (const event of gen) {
      send(event);
    }
  } catch (e) {
    send({ type: 'result', success: false, error: e.message });
  } finally {
    res.end();
  }
});

// POST /api/scim/test-connection
router.post('/test-connection', requireAuth, async (req: Request, res: Response) => {
  try {
    const { scimUrl, scimToken } = req.body || {};
    const result = await g.testScimConnection(scimUrl, scimToken);
    if (!result.ok) return res.status(400).json(result);
    res.json(result);
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

module.exports = router;
