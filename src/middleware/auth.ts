'use strict';

type Request = import('express').Request
type Response = import('express').Response
type NextFunction = import('express').NextFunction

/**
 * Require an authenticated session. Stores original URL and redirects to /auth/login.
 */
function requireAuth(req: Request, res: Response, next: NextFunction) {
  if (req.session?.auth?.accessToken) {
    // Check token expiry
    if (req.session.auth.expiresAt && Date.now() > req.session.auth.expiresAt) {
      req.session.auth = null;
      if (req.xhr || req.headers.accept?.includes('application/json')) {
        return res.status(401).json({ error: 'Session expired. Please sign in again.' });
      }
      req.session.returnTo = req.originalUrl;
      return res.redirect('/auth/login?expired=1');
    }
    return next();
  }

  if (req.xhr || req.headers.accept?.includes('application/json')) {
    return res.status(401).json({ error: 'Not authenticated' });
  }
  req.session.returnTo = req.originalUrl;
  res.redirect('/auth/login');
}

/**
 * Attach auth info to res.locals for templates
 */
function attachLocals(req: Request, res: Response, next: NextFunction) {
  res.locals.isAuthenticated = !!(req.session?.auth?.accessToken);
  res.locals.tenantId = req.session?.auth?.tenantId || '';
  res.locals.orgName = req.session?.auth?.orgName || 'Not Connected';
  res.locals.userName = req.session?.auth?.userName || '';
  res.locals.lastSync = req.session?.auth?.lastSync || null;
  res.locals.currentPath = req.path;
  next();
}

module.exports = { requireAuth, attachLocals };
