'use strict';

const crypto = require('crypto');
const express = require('express');
const router = express.Router();
const { getAuthorizationCodeToken, getOrganization, getMe } = require('../services/graphService.ts');

type Request = import('express').Request
type Response = import('express').Response

interface EntraConfig {
  tenantId: string
  clientId: string
  clientSecret: string
  redirectUri: string
  postLogoutRedirectUri: string
  scopes: string
}

function getEntraConfig(): EntraConfig {
  return {
    tenantId: (process.env.ENTRA_TENANT_ID || 'organizations').trim() || 'organizations',
    clientId: (process.env.ENTRA_CLIENT_ID || '').trim(),
    clientSecret: (process.env.ENTRA_CLIENT_SECRET || '').trim(),
    redirectUri: (process.env.ENTRA_REDIRECT_URI || 'http://localhost:3000/api/auth/callback').trim(),
    postLogoutRedirectUri: (process.env.ENTRA_POST_LOGOUT_REDIRECT_URI || 'http://localhost:3000/').trim(),
    scopes: (process.env.ENTRA_SCOPES || 'openid profile offline_access User.Read Directory.Read.All Application.ReadWrite.All AppRoleAssignment.ReadWrite.All DelegatedPermissionGrant.ReadWrite.All Group.Read.All Synchronization.ReadWrite.All').trim(),
  };
}

function normalizeTenant(tenant: unknown): string {
  const raw = String(tenant || '').trim();
  if (!raw) return 'organizations';
  if (/^(organizations|common|consumers)$/i.test(raw)) return raw.toLowerCase();
  if (/^[A-Za-z0-9.-]+$/.test(raw)) return raw;
  return 'organizations';
}

function validateConfig(cfg: EntraConfig): void {
  if (!cfg.clientId || !cfg.clientSecret || !cfg.redirectUri) {
    throw new Error('Missing ENTRA_CLIENT_ID, ENTRA_CLIENT_SECRET, or ENTRA_REDIRECT_URI in .env');
  }
}

function buildAuthorizeUrl(cfg: EntraConfig, state: string): string {
  const params = new URLSearchParams({
    client_id: cfg.clientId,
    response_type: 'code',
    redirect_uri: cfg.redirectUri,
    response_mode: 'query',
    scope: cfg.scopes,
    state,
  });
  return `https://login.microsoftonline.com/${cfg.tenantId}/oauth2/v2.0/authorize?${params.toString()}`;
}

async function buildSessionFromCode(code: string, cfg: EntraConfig) {
  const tokenData = await getAuthorizationCodeToken({
    tenantId: cfg.tenantId,
    clientId: cfg.clientId,
    clientSecret: cfg.clientSecret,
    code,
    redirectUri: cfg.redirectUri,
    scopes: cfg.scopes,
  });

  let orgName = 'Connected';
  let orgId = cfg.tenantId;
  let tenantId = cfg.tenantId;
  let lastSync = null;
  let userName = '';

  try {
    const me = await getMe(tokenData.accessToken);
    userName = me?.displayName || me?.userPrincipalName || me?.mail || '';
  } catch (_) { /* non-fatal */ }

  try {
    const orgRes = await getOrganization(tokenData.accessToken);
    const org = orgRes?.value?.[0];
    if (org) {
      orgName = org.displayName || orgName;
      orgId = org.id || orgId;
      tenantId = org.id || tenantId;
      lastSync = org.onPremisesLastSyncDateTime || null;
    } else if (userName) {
      orgName = userName;
    }
  } catch (_) {
    if (userName) orgName = userName;
  }

  return {
    accessToken: tokenData.accessToken,
    refreshToken: tokenData.refreshToken,
    expiresAt: tokenData.expiresAt,
    tenantId,
    clientId: cfg.clientId,
    orgName,
    orgId,
    userName,
    lastSync,
    connectedAt: new Date().toISOString(),
  };
}

function renderLogin(res: Response, error: string | null = null) {
  const cfg = getEntraConfig();
  res.render('login', {
    title: 'Sign In — Litera One Portal',
    error,
    tenantAuthority: cfg.tenantId,
    clientId: cfg.clientId || 'Not configured',
    redirectUri: cfg.redirectUri || 'Not configured',
    isConfigured: !!(cfg.clientId && cfg.clientSecret && cfg.redirectUri),
  });
}

// GET /auth/login => redirect to Microsoft Entra sign-in page
router.get('/login', (req: Request, res: Response) => {
  if (req.session?.auth?.accessToken) return res.redirect('/');

  try {
    const cfg = getEntraConfig();
    validateConfig(cfg);

    const selectedTenant = normalizeTenant(req.query.tenant || cfg.tenantId);
    cfg.tenantId = selectedTenant;

    const state = crypto.randomBytes(16).toString('hex');
    req.session.oauthState = state;
    req.session.oauthTenant = selectedTenant;

    const requestedReturnTo = typeof req.query.returnTo === 'string' ? req.query.returnTo : '';
    if (requestedReturnTo.startsWith('/')) req.session.returnTo = requestedReturnTo;

    return res.redirect(buildAuthorizeUrl(cfg, state));
  } catch (err: any) {
    return renderLogin(res, err.message || 'Unable to start Microsoft sign-in.');
  }
});

// GET /auth/callback
router.get('/callback', async (req: Request, res: Response) => {
  const cfg = getEntraConfig();
  cfg.tenantId = normalizeTenant(req.session.oauthTenant || cfg.tenantId);

  if (req.query.error) {
    const msg = req.query.error_description || req.query.error || 'Sign-in was cancelled or failed.';
    return renderLogin(res, String(msg));
  }

  const code = typeof req.query.code === 'string' ? req.query.code : '';
  const state = typeof req.query.state === 'string' ? req.query.state : '';
  const expectedState = req.session.oauthState;
  delete req.session.oauthState;
  delete req.session.oauthTenant;

  if (!code) return renderLogin(res, 'Missing authorization code from Microsoft sign-in response.');
  if (!expectedState || state !== expectedState) return renderLogin(res, 'Invalid sign-in state. Please try again.');

  try {
    validateConfig(cfg);
    const authSession = await buildSessionFromCode(code, cfg);
    const returnTo = req.session.returnTo || '/';
    await new Promise((resolve, reject) => {
      req.session.regenerate(err => (err ? reject(err) : resolve()));
    });
    req.session.auth = authSession;
    return res.redirect(returnTo);
  } catch (err: any) {
    return renderLogin(res, err.message || 'Authentication failed during callback processing.');
  }
});

// POST /auth/login (legacy form submit support)
router.post('/login', (_req: Request, res: Response) => res.redirect('/auth/login'));

// POST /auth/logout
router.post('/logout', (req: Request, res: Response) => {
  const cfg = getEntraConfig();
  const logoutParams = new URLSearchParams({
    post_logout_redirect_uri: cfg.postLogoutRedirectUri,
  });
  const logoutUrl = `https://login.microsoftonline.com/${cfg.tenantId}/oauth2/v2.0/logout?${logoutParams.toString()}`;

  req.session.destroy(() => res.redirect(logoutUrl));
});

// GET /auth/status
router.get('/status', (req: Request, res: Response) => {
  if (req.session?.auth?.accessToken) {
    res.json({
      authenticated: true,
      orgName: req.session.auth.orgName,
      tenantId: req.session.auth.tenantId,
      connectedAt: req.session.auth.connectedAt,
      lastSync: req.session.auth.lastSync,
      expiresAt: req.session.auth.expiresAt,
    });
  } else {
    res.json({ authenticated: false });
  }
});

module.exports = router;
