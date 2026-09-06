'use strict';

const express = require('express');
const router = express.Router();
const { requireAuth } = require('../../middleware/auth.ts');
const g = require('../../services/graphService.ts');

type Request = import('express').Request
type Response = import('express').Response

const token = (req: Request) => req.session.auth?.accessToken || '';

// GET /api/principals/search?q=john&type=users
router.get('/search', requireAuth, async (req: Request, res: Response) => {
  try {
    const result = await g.searchPrincipals(
      token(req),
      req.query.q,
      req.query.type || 'users',
      {
        top: req.query.top,
        nextLink: req.query.nextLink,
      }
    );
    res.json({
      principals: result?.value || [],
      nextLink: result?.['@odata.nextLink'] || null,
    });
  } catch (e) {
    res.status(e.status || 500).json({ error: e.message });
  }
});

module.exports = router;
