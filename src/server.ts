'use strict';

require('dotenv').config();
const express = require('express');
const session = require('express-session');
const helmet = require('helmet');
const morgan = require('morgan');
const path = require('path');
const crypto = require('crypto');
const rateLimit = require('express-rate-limit');

const { attachLocals } = require('./middleware/auth.ts');
const authRouter = require('./routes/auth.ts');
const appsRouter = require('./routes/api/apps.ts');
const scimRouter = require('./routes/api/scim.ts');
const principalsRouter = require('./routes/api/principals.ts');
const deployRouter = require('./routes/api/deploy.ts');
const enablementRouter = require('./routes/api/enablement.ts');

const app = express();
const PORT = process.env.PORT || 3000;
const sessionSecret = String(process.env.SESSION_SECRET || '').trim();
const sessionCookieSecureRaw = String(process.env.SESSION_COOKIE_SECURE || '').trim().toLowerCase();
const sessionCookieSecure =
  sessionCookieSecureRaw === 'true'
    ? true
    : sessionCookieSecureRaw === 'false'
      ? false
      : process.env.NODE_ENV === 'production';
type Request = import('express').Request
type Response = import('express').Response
type NextFunction = import('express').NextFunction

function getEnablementAdminUrlPath() {
  const raw = String(process.env.ENABLEMENT_ADMIN_LOGIN_URL || process.env.ENABLEMENT_ADMIN_URL_PATH || '').trim();
  if (!raw) return '/';
  try {
    const url = new URL(raw);
    const normalized = url.pathname.replace(/\/+$/, '') || '/';
    if (normalized.includes(' ')) return '/';
    return normalized;
  } catch (error) {
    if (!raw.startsWith('/')) return '/';
    // Keep the path segment clean and predictable for routing.
    const normalized = raw.replace(/\/+$/, '') || '/';
    if (normalized.includes(' ')) return '/';
    return normalized;
  }
}

if (!sessionSecret) {
  if (process.env.NODE_ENV === 'production') {
    throw new Error('SESSION_SECRET is required in production.');
  }
  console.warn('[security] SESSION_SECRET is not set. Using an insecure development fallback.');
}

// ─── Security ─────────────────────────────────────────────────────────────────
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'", 'https://fonts.googleapis.com'],
      scriptSrcAttr: ["'unsafe-inline'"],
      styleSrc: ["'self'", "'unsafe-inline'", 'https://fonts.googleapis.com', 'https://fonts.gstatic.com'],
      fontSrc: ["'self'", 'https://fonts.googleapis.com', 'https://fonts.gstatic.com'],
      connectSrc: ["'self'"],
      imgSrc: ["'self'", 'data:'],
    },
  },
}));

// ─── Rate limiting ────────────────────────────────────────────────────────────
const apiLimiter = rateLimit({ windowMs: 60 * 1000, max: 120, standardHeaders: true, legacyHeaders: false });
const authLimiter = rateLimit({ windowMs: 15 * 60 * 1000, max: 20 });

// ─── Middleware ───────────────────────────────────────────────────────────────
app.use(morgan(process.env.NODE_ENV === 'production' ? 'combined' : '[:date[iso]] :method :url :status :response-time ms'));
app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true, limit: '2mb' }));
app.use(express.static(path.join(__dirname, '../public')));

// ─── Sessions ─────────────────────────────────────────────────────────────────
app.use(session({
  name: 'litera.sid',
  secret: sessionSecret || 'litera-one-dev-secret-change-me',
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: sessionCookieSecure,
    httpOnly: true,
    sameSite: 'lax',
    maxAge: 8 * 60 * 60 * 1000, // 8 hours
  },
}));

function ensureCsrfToken(req: Request, res: Response, next: NextFunction) {
  if (!req.session.csrfToken) req.session.csrfToken = crypto.randomBytes(24).toString('hex');
  res.locals.csrfToken = req.session.csrfToken;
  next();
}

function verifyCsrf(req: Request, res: Response, next: NextFunction) {
  const protectedMethods = new Set(['POST', 'PUT', 'PATCH', 'DELETE']);
  if (!protectedMethods.has(req.method)) return next();

  const sessionToken = req.session?.csrfToken;
  const providedToken = req.get('x-csrf-token') || req.body?._csrf;
  if (sessionToken && providedToken && sessionToken === providedToken) return next();

  if (req.path.startsWith('/api/')) {
    return res.status(403).json({ error: 'CSRF token invalid or missing.' });
  }
  return res.status(403).render('error', {
    title: '403 Forbidden',
    message: 'Request blocked by CSRF protection. Refresh and try again.',
    code: 403,
  });
}

// ─── View engine ──────────────────────────────────────────────────────────────
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, '../public/views'));
app.use(ensureCsrfToken);
app.use(verifyCsrf);
app.use(attachLocals);

// Prevent stale API responses from browser/proxy caches.
app.use('/api', (_req: Request, res: Response, next: NextFunction) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  next();
});

// ─── Routes ───────────────────────────────────────────────────────────────────
app.use('/auth', authLimiter, authRouter);
app.use('/api/auth', authLimiter, authRouter);
app.use('/api/apps', apiLimiter, appsRouter);
app.use('/api/scim', apiLimiter, scimRouter);
app.use('/api/principals', apiLimiter, principalsRouter);
app.use('/api/deploy', apiLimiter, deployRouter);
app.use('/api/enablement', apiLimiter, enablementRouter);

app.get('/health', (_req: Request, res: Response) => {
  res.status(200).json({ status: 'ok' });
});

// Main portal routes
app.get('/', (_req: Request, res: Response) => {
  res.render('portal', {
    title: 'Litera One Manager Portal',
    initialPage: 'dashboard',
    standaloneEnablementAdmin: false,
  });
});

app.get('/enablement-admin-login', (_req: Request, res: Response) => {
  res.render('portal', {
    title: 'Litera One Manager Portal',
    initialPage: 'enablement-admin',
    standaloneEnablementAdmin: true,
  });
});

// Dedicated URL for Enablement Admin login/config page (not in sidebar nav)
const enablementAdminUrlPath = getEnablementAdminUrlPath();
app.get(enablementAdminUrlPath, (_req: Request, res: Response) => {
  res.render('portal', {
    title: 'Litera One Manager Portal',
    initialPage: 'enablement-admin',
    standaloneEnablementAdmin: true,
  });
});

// ─── 404 / Error handlers ─────────────────────────────────────────────────────
app.use((req: Request, res: Response) => {
  if (req.path.startsWith('/api/')) return res.status(404).json({ error: 'Not found' });
  res.status(404).render('error', { title: '404 Not Found', message: 'Page not found', code: 404 });
});

app.use((err: any, req: Request, res: Response, _next: NextFunction) => {
  console.error(err.stack);
  if (req.path.startsWith('/api/')) return res.status(500).json({ error: err.message || 'Internal server error' });
  res.status(500).render('error', { title: 'Server Error', message: err.message, code: 500 });
});

app.listen(PORT, () => {
  console.log(`\n  ┌─────────────────────────────────────────┐`);
  console.log(`  │  Litera One Manager Portal               │`);
  console.log(`  │  Running at http://localhost:${PORT}        │`);
  console.log(`  └─────────────────────────────────────────┘\n`);
  console.log(`Server accessible at: http://localhost:${PORT}`);
  console.log(`Enablement Admin URL: http://localhost:${PORT}${enablementAdminUrlPath}`);
});

module.exports = app;
