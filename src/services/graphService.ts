'use strict';

const { GRAPH_BASE_V1, GRAPH_BASE_BETA, NON_GALLERY_TEMPLATE_ID } = require('../../config/graph');

// ─── Core request helper ─────────────────────────────────────────────────────
function formatFetchError(context, err) {
  const cause = err?.cause;
  const bits = [];
  if (cause?.code) bits.push(`code=${cause.code}`);
  if (cause?.errno) bits.push(`errno=${cause.errno}`);
  if (cause?.syscall) bits.push(`syscall=${cause.syscall}`);
  if (cause?.hostname) bits.push(`host=${cause.hostname}`);
  if (cause?.message) bits.push(`cause=${cause.message}`);
  const detail = bits.length ? ` (${bits.join(', ')})` : '';
  return `${context}: ${err?.message || 'Network request failed'}${detail}`;
}

async function graphRequest(token, method, path, body = null, { beta = false, expectNoContent = false } = {}) {
  const base = beta ? GRAPH_BASE_BETA : GRAPH_BASE_V1;
  const url = path.startsWith('http') ? path : `${base}${path.startsWith('/') ? path : '/' + path}`;

  const opts = {
    method: method.toUpperCase(),
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ConsistencyLevel: 'eventual',
    },
  };
  if (body) opts.body = JSON.stringify(body);

  let res, data;
  try {
    res = await fetch(url, opts);
  } catch (err) {
    throw new Error(`Network error calling ${method} ${url}: ${err.message}`);
  }

  if (!res.ok) {
    const text = await res.text().catch(() => '');
    try { data = JSON.parse(text); } catch { data = { rawText: text }; }
    const msg = data?.error?.message || data?.error_description || `HTTP ${res.status}`;
    const err = new Error(msg);
    err.status = res.status;
    err.graphError = data?.error;
    throw err;
  }

  if (res.status === 204 || expectNoContent) return true;

  const text = await res.text();
  try { data = JSON.parse(text); } catch { data = { rawText: text }; }
  return data;
}

async function graphGetAll(token, path, { beta = false } = {}) {
  const results = [];
  let url = path;
  while (url) {
    const data = await graphRequest(token, 'GET', url, null, { beta });
    if (data?.value) results.push(...data.value);
    url = data?.['@odata.nextLink'] || null;
  }
  return results;
}

function wait(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ─── Auth ─────────────────────────────────────────────────────────────────────

async function getClientCredentialsToken(tenantId, clientId, clientSecret) {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
  });
  let res;
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
  } catch (err) {
    throw new Error(formatFetchError('Could not reach Microsoft token endpoint', err));
  }
  const data = await res.json();
  if (!res.ok) {
    const msg = data.error_description || data.error || 'Authentication failed';
    const err = new Error(msg);
    err.status = res.status;
    throw err;
  }
  return {
    accessToken: data.access_token,
    expiresIn: data.expires_in,
    expiresAt: Date.now() + (data.expires_in - 60) * 1000,
  };
}

async function getAuthorizationCodeToken({
  tenantId,
  clientId,
  clientSecret,
  code,
  redirectUri,
  scopes,
}) {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'authorization_code',
    client_id: clientId,
    client_secret: clientSecret,
    code,
    redirect_uri: redirectUri,
    scope: scopes,
  });
  let res;
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
  } catch (err) {
    throw new Error(formatFetchError('Could not reach Microsoft token endpoint', err));
  }
  const data = await res.json();
  if (!res.ok) {
    const msg = data.error_description || data.error || 'Authentication failed';
    const err = new Error(msg);
    err.status = res.status;
    throw err;
  }
  return {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || null,
    idToken: data.id_token || null,
    expiresIn: data.expires_in,
    expiresAt: Date.now() + (data.expires_in - 60) * 1000,
  };
}

// ─── Organisation / Profile ───────────────────────────────────────────────────

async function getOrganization(token) {
  return graphRequest(token, 'GET', '/organization?$select=id,displayName,onPremisesLastSyncDateTime');
}

async function getMe(token) {
  return graphRequest(token, 'GET', '/me?$select=id,displayName,userPrincipalName,mail');
}

// ─── Service Principals / Apps ────────────────────────────────────────────────

async function searchServicePrincipals(token, query) {
  const q = (query || 'Litera').replace(/'/g, "''");
  return graphRequest(token, 'GET',
    `/servicePrincipals?$filter=startswith(displayName,'${q}')&$select=id,displayName,appId,tags,accountEnabled,appRoleAssignmentRequired&$top=100`
  );
}

async function getServicePrincipal(token, spId, { beta = false } = {}) {
  return graphRequest(token, 'GET',
    `/servicePrincipals/${spId}?$select=id,displayName,appId,tags,accountEnabled,appRoles,appRoleAssignmentRequired`,
    null, { beta }
  );
}

async function updateServicePrincipal(token, spId, patchBody) {
  return graphRequest(token, 'PATCH', `/servicePrincipals/${spId}`, patchBody, { expectNoContent: true });
}

async function deleteServicePrincipal(token, spId) {
  return graphRequest(token, 'DELETE', `/servicePrincipals/${spId}`, null, { expectNoContent: true });
}

async function createServicePrincipal(token, appId) {
  return graphRequest(token, 'POST', '/servicePrincipals', { appId });
}

// ─── Application Registrations ────────────────────────────────────────────────

async function getApplicationByAppId(token, appId) {
  const q = appId.replace(/'/g, "''");
  return graphRequest(token, 'GET',
    `/applications?$filter=appId eq '${q}'&$select=id,appId,displayName,requiredResourceAccess,tags`
  );
}

async function createApplication(token, displayName) {
  return graphRequest(token, 'POST', '/applications', {
    displayName,
    signInAudience: 'AzureADMyOrg',
    tags: ['HideApp', 'WindowsAzureActiveDirectoryIntegratedApp'],
  });
}

async function deleteApplication(token, appObjId) {
  return graphRequest(token, 'DELETE', `/applications/${appObjId}`, null, { expectNoContent: true });
}

async function updateApplicationTags(token, appObjId, tags) {
  return graphRequest(token, 'PATCH', `/applications/${appObjId}`, { tags }, { expectNoContent: true });
}

async function instantiateTemplate(token, templateId, displayName) {
  return graphRequest(token, 'POST', `/applicationTemplates/${templateId}/instantiate`, { displayName });
}

// ─── Permissions / Consent ────────────────────────────────────────────────────

async function getAppPermissions(token, appId) {
  const appRes = await getApplicationByAppId(token, appId);
  const app = appRes?.value?.[0];
  if (!app) return [];

  const chips = [];
  for (const rra of (app.requiredResourceAccess || [])) {
    try {
      const spRes = await graphRequest(token, 'GET',
        `/servicePrincipals?$filter=appId eq '${rra.resourceAppId}'&$select=displayName,oauth2PermissionScopes,appRoles`
      );
      const sp = spRes?.value?.[0];
      if (!sp) continue;
      for (const ra of (rra.resourceAccess || [])) {
        let name = ra.id;
        if (ra.type === 'Scope') {
          const s = (sp.oauth2PermissionScopes || []).find(x => x.id === ra.id);
          if (s) name = s.value;
        } else {
          const r = (sp.appRoles || []).find(x => x.id === ra.id);
          if (r) name = r.value || r.displayName || ra.id;
        }
        chips.push({ name, type: ra.type });
      }
    } catch (_) { /* non-fatal */ }
  }
  return chips;
}

async function getOAuth2Grants(token, clientSpId) {
  return graphRequest(token, 'GET',
    `/oauth2PermissionGrants?$filter=clientId eq '${clientSpId}'`
  );
}

async function getPrincipalAppRoleAssignments(token, clientSpId) {
  return graphRequest(token, 'GET',
    `/servicePrincipals/${clientSpId}/appRoleAssignments?$top=999`
  );
}

async function grantAdminConsent(token, clientSp) {
  const appRes = await getApplicationByAppId(token, clientSp.appId);
  const app = appRes?.value?.[0];
  if (!app) throw new Error('Application registration not found for consent.');

  const results = [];
  for (const rra of (app.requiredResourceAccess || [])) {
    try {
      const spR = await graphRequest(token, 'GET',
        `/servicePrincipals?$filter=appId eq '${rra.resourceAppId}'&$select=id,displayName,oauth2PermissionScopes,appRoles`
      );
      const resSp = spR?.value?.[0];
      if (!resSp) continue;

      const scopes = [];
      const roles = [];
      for (const ra of (rra.resourceAccess || [])) {
        if (ra.type === 'Scope') {
          const s = (resSp.oauth2PermissionScopes || []).find(x => x.id === ra.id);
          if (s) scopes.push(s.value);
        } else {
          roles.push({ id: ra.id, resourceId: resSp.id, resourceName: resSp.displayName });
        }
      }

      if (scopes.length) {
        const existing = await graphRequest(token, 'GET',
          `/oauth2PermissionGrants?$filter=clientId eq '${clientSp.id}' and resourceId eq '${resSp.id}'`
        );
        const grant = existing?.value?.[0];
        if (grant) {
          const merged = [...new Set([...grant.scope.split(' '), ...scopes])].join(' ');
          if (merged !== grant.scope) {
            await graphRequest(token, 'PATCH', `/oauth2PermissionGrants/${grant.id}`, { scope: merged }, { expectNoContent: true });
          }
        } else {
          await graphRequest(token, 'POST', '/oauth2PermissionGrants', {
            clientId: clientSp.id, consentType: 'AllPrincipals',
            resourceId: resSp.id, scope: scopes.join(' '),
          });
        }
        results.push({ resource: resSp.displayName, type: 'delegated', scopes });
      }

      for (const role of roles) {
        try {
          await graphRequest(token, 'POST', `/servicePrincipals/${clientSp.id}/appRoleAssignments`, {
            principalId: clientSp.id, resourceId: role.resourceId, appRoleId: role.id,
          });
          results.push({ resource: role.resourceName, type: 'application', roleId: role.id });
        } catch (e) {
          if (!e.message.includes('already exists') && !e.message.includes('Permission being assigned already')) throw e;
        }
      }
    } catch (e) {
      results.push({ error: e.message });
    }
  }
  return results;
}

// ─── Assignments ──────────────────────────────────────────────────────────────

async function getAppRoleAssignments(token, spId) {
  const path = `/servicePrincipals/${spId}/appRoleAssignedTo?$top=999`;
  let v1 = [];
  let beta = [];

  try {
    v1 = await graphGetAll(token, path);
  } catch (_) {
    v1 = [];
  }
  try {
    beta = await graphGetAll(token, path, { beta: true });
  } catch (_) {
    beta = [];
  }

  if (!v1.length && !beta.length) return [];

  // Some tenants return partial assignment visibility in v1 vs beta. Merge by assignment id.
  const byId = new Map();
  for (const row of [...v1, ...beta]) {
    if (!row?.id) continue;
    if (!byId.has(row.id)) byId.set(row.id, row);
  }
  return [...byId.values()];
}

async function enrichAssignments(token, assignments) {
  const ids = [...new Set(
    assignments
      .filter(a => ['User', 'Group'].includes(a.principalType) && a.principalId)
      .map(a => a.principalId)
  )];

  const lookup = {};
  const chunkSize = 900;
  for (let i = 0; i < ids.length; i += chunkSize) {
    const chunk = ids.slice(i, i + chunkSize);
    try {
      const res = await graphRequest(token, 'POST', '/directoryObjects/getByIds',
        { ids: chunk, types: ['user', 'group'] }
      );
      for (const obj of (res?.value || [])) {
        lookup[obj.id] = obj.mail || obj.userPrincipalName || obj.id;
      }
    } catch (_) { /* non-fatal */ }
  }

  return assignments.map(a => ({
    ...a,
    emailOrId: lookup[a.principalId] || a.principalId,
  }));
}

async function assignPrincipal(token, spId, principalId, principalType) {
  // Resolve best appRoleId
  let appRoleId = '00000000-0000-0000-0000-000000000000';
  try {
    const sp = await graphRequest(token, 'GET', `/servicePrincipals/${spId}?$select=appRoles,appRoleAssignmentRequired`);
    // In Entra app roles, groups are assigned through roles that allow "User".
    const memberType = principalType === 'Application' ? 'Application' : 'User';
    const eligible = (sp?.appRoles || []).filter(r =>
      r.isEnabled && (r.allowedMemberTypes || []).includes(memberType)
    );
    if (eligible.length) appRoleId = eligible[0].id;
  } catch (e) {
    // Fall back to default role when role discovery is unavailable.
  }

  return graphRequest(token, 'POST', `/servicePrincipals/${spId}/appRoleAssignedTo`, {
    principalId, resourceId: spId, appRoleId,
  });
}

async function testScimConnection(scimUrl, scimToken) {
  const rawUrl = String(scimUrl || '').trim();
  const token = String(scimToken || '').trim();
  if (!rawUrl) throw new Error('SCIM tenant URL is required.');
  if (!token) throw new Error('SCIM secret token is required.');

  const normalized = rawUrl.replace(/\/+$/, '');
  const candidates = [
    `${normalized}/ServiceProviderConfig`,
    `${normalized}/Schemas`,
    normalized,
  ];

  const errors = [];
  for (const url of candidates) {
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 10000);
      const res = await fetch(url, {
        method: 'GET',
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/scim+json, application/json' },
        signal: controller.signal,
      });
      clearTimeout(timeout);

      const text = await res.text().catch(() => '');
      if (res.ok) {
        return { ok: true, url, status: res.status, responsePreview: text.slice(0, 500) };
      }
      errors.push(`${url} -> HTTP ${res.status}${text ? `: ${text.slice(0, 120)}` : ''}`);
    } catch (e) {
      errors.push(`${url} -> ${e.name === 'AbortError' ? 'Timeout' : e.message}`);
    }
  }

  return { ok: false, errors };
}

async function removeAppRoleAssignment(token, spId, assignmentId) {
  return graphRequest(token, 'DELETE', `/servicePrincipals/${spId}/appRoleAssignedTo/${assignmentId}`,
    null, { expectNoContent: true }
  );
}

// ─── Groups & Users ───────────────────────────────────────────────────────────

async function searchPrincipals(token, query, type = 'users', options = {}) {
  const safeType = type === 'groups' ? 'groups' : 'users';
  const selectF = safeType === 'groups' ? 'id,displayName,mail' : 'id,displayName,userPrincipalName,mail';
  const q = (query || '').trim();
  const rawTop = Number.parseInt(options.top, 10);
  const top = Number.isFinite(rawTop) ? Math.max(1, Math.min(rawTop, 100)) : 20;
  const nextLink = String(options.nextLink || '').trim();

  if (nextLink) {
    if (!/^https:\/\/graph\.microsoft\.com\//i.test(nextLink)) {
      throw new Error('Invalid nextLink host.');
    }
    return graphRequest(token, 'GET', nextLink);
  }

  if (!q) {
    return graphRequest(token, 'GET', `/${safeType}?$select=${selectF}&$top=${top}`);
  }
  return graphRequest(token, 'GET',
    `/${safeType}?$search="displayName:${encodeURIComponent(q)}"&$select=${selectF}&$top=${top}`
  );
}

async function getGroupMembers(token, groupId) {
  return graphGetAll(token, `/groups/${groupId}/members?$select=displayName,userPrincipalName,id`);
}

async function getUserAppAssignmentsOnce(token, userId) {
  const path = `/users/${userId}/appRoleAssignments?$select=resourceDisplayName,resourceId,id`;
  let v1 = [];
  let beta = [];

  try {
    v1 = await graphGetAll(token, path);
  } catch (_) {
    v1 = [];
  }
  try {
    beta = await graphGetAll(token, path, { beta: true });
  } catch (_) {
    beta = [];
  }

  const byId = new Map();
  for (const row of [...v1, ...beta]) {
    if (!row?.id) continue;
    if (!byId.has(row.id)) byId.set(row.id, row);
  }
  return [...byId.values()];
}

async function getUserAppAssignments(token, userId, { attempts = 4, delayMs = 900 } = {}) {
  let best = [];
  let lastError = null;

  for (let i = 0; i < attempts; i++) {
    try {
      const current = await getUserAppAssignmentsOnce(token, userId);
      if (current.length >= best.length) best = current;
      if (best.length > 0) return best;
    } catch (e) {
      lastError = e;
      if (i === attempts - 1) throw e;
    }

    if (i < attempts - 1) {
      await wait(delayMs);
    }
  }

  if (best.length > 0) return best;
  if (lastError) throw lastError;
  return [];
}

// ─── Tags / Visibility ────────────────────────────────────────────────────────

async function getSpTagsDetailed(token, spId) {
  try {
    const sp = await graphRequest(token, 'GET',
      `/servicePrincipals/${spId}?$select=id,appId,tags,servicePrincipalType,appOwnerOrganizationId`
    );
    return sp;
  } catch (_) {
    return graphRequest(token, 'GET',
      `/servicePrincipals/${spId}?$select=id,appId,tags,servicePrincipalType,appOwnerOrganizationId`,
      null, { beta: true }
    );
  }
}

async function setVisibility(token, spId, appId, hide) {
  const spInfo = await getSpTagsDetailed(token, spId);
  const tags = [...new Set([...(spInfo?.tags || [])])];
  const ensureTag = 'WindowsAzureActiveDirectoryIntegratedApp';
  if (!tags.includes(ensureTag)) tags.push(ensureTag);

  if (hide) {
    if (!tags.includes('HideApp')) tags.push('HideApp');
  } else {
    const idx = tags.indexOf('HideApp');
    if (idx > -1) tags.splice(idx, 1);
  }

  await updateServicePrincipal(token, spId, { tags });

  // Also sync to app registration tags if possible
  try {
    if (appId) {
      const appRes = await getApplicationByAppId(token, appId);
      const appObjId = appRes?.value?.[0]?.id;
      if (appObjId) await updateApplicationTags(token, appObjId, tags);
    }
  } catch (_) { /* non-fatal */ }

  return tags;
}

// ─── Synchronization / SCIM ───────────────────────────────────────────────────

async function getSyncJobs(token, spId) {
  try {
    const res = await graphRequest(token, 'GET', `/servicePrincipals/${spId}/synchronization/jobs`);
    return res?.value || [];
  } catch (_) {
    try {
      const res = await graphRequest(token, 'GET', `/servicePrincipals/${spId}/synchronization/jobs`, null, { beta: true });
      return res?.value || [];
    } catch (_) { return []; }
  }
}

async function getSyncTemplates(token, spId) {
  try {
    const res = await graphRequest(token, 'GET', `/servicePrincipals/${spId}/synchronization/templates`);
    return res?.value || [];
  } catch (_) {
    try {
      const res = await graphRequest(token, 'GET', `/servicePrincipals/${spId}/synchronization/templates`, null, { beta: true });
      return res?.value || [];
    } catch (_) { return []; }
  }
}

async function createSyncJob(token, spId, templateId = 'scim') {
  try {
    return graphRequest(token, 'POST', `/servicePrincipals/${spId}/synchronization/jobs`, { templateId });
  } catch (_) {
    return graphRequest(token, 'POST', `/servicePrincipals/${spId}/synchronization/jobs`, { templateId }, { beta: true });
  }
}

async function startSyncJob(token, spId, jobId) {
  try {
    return graphRequest(token, 'POST', `/servicePrincipals/${spId}/synchronization/jobs/${jobId}/start`, {}, { expectNoContent: true });
  } catch (_) { return false; }
}

async function restartSyncJob(token, spId, jobId) {
  try {
    return graphRequest(token, 'POST', `/servicePrincipals/${spId}/synchronization/jobs/${jobId}/restart`,
      { criteria: { resetScope: 'Full' } }, { expectNoContent: true }
    );
  } catch (_) {
    return graphRequest(token, 'POST', `/servicePrincipals/${spId}/synchronization/jobs/${jobId}/restart`,
      { criteria: { resetScope: 'Full' } }, { beta: true, expectNoContent: true }
    );
  }
}

async function setSyncSecrets(token, spId, scimUrl, scimToken) {
  const body = {
    value: [
      { key: 'BaseAddress', value: scimUrl },
      { key: 'SecretToken', value: scimToken },
    ],
  };
  const uris = [
    `/servicePrincipals/${spId}/synchronization/secrets`,
  ];
  for (const uri of uris) {
    for (const method of ['PUT', 'PATCH']) {
      try {
        await graphRequest(token, method, uri, body, { expectNoContent: true });
        return true;
      } catch (_) { /* try next */ }
    }
  }
  return false;
}

// ─── Full SCIM creation flow ──────────────────────────────────────────────────

async function* createScimApp(token, { appName, scimUrl, scimToken, principals = [], probeSpId = null }) {
  yield { type: 'progress', message: 'Validating permissions...', percent: 5 };

  // Preflight: check synchronization access on any known SP
  if (probeSpId) {
    try {
      await graphRequest(token, 'GET', `/servicePrincipals/${probeSpId}/synchronization/templates`);
    } catch (e) {
      if (e.status === 401 || e.status === 403) {
        throw new Error('Access denied reading synchronization templates. Ensure Synchronization.ReadWrite.All is consented.');
      }
    }
  }

  yield { type: 'progress', message: 'Creating enterprise application...', percent: 15 };

  let newSpId = null, newAppObjId = null;
  let creationMode = 'unknown';

  // Try non-gallery template first
  try {
    yield { type: 'diag', message: `Instantiating non-gallery template ${NON_GALLERY_TEMPLATE_ID}` };
    const inst = await instantiateTemplate(token, NON_GALLERY_TEMPLATE_ID, appName);
    if (inst?.servicePrincipal?.id) {
      newSpId = inst.servicePrincipal.id;
      newAppObjId = inst.application?.id;
      creationMode = 'nongallery';
      yield { type: 'diag', message: `Non-gallery SP created: spId=${newSpId}` };
    }
  } catch (e) {
    yield { type: 'diag', message: `Non-gallery template failed: ${e.message}. Falling back.` };
  }

  // Fallback: create app + SP directly
  if (!newSpId) {
    try {
      const newApp = await createApplication(token, appName);
      newAppObjId = newApp.id;
      const newSp = await createServicePrincipal(token, newApp.appId);
      newSpId = newSp.id;
      creationMode = 'appPlusSp';
      yield { type: 'diag', message: `Fallback: created app+SP: spId=${newSpId}` };
    } catch (e) {
      throw new Error('Could not create enterprise application: ' + e.message);
    }
  }

  yield { type: 'progress', message: 'Configuring service principal...', percent: 30 };

  try {
    await updateServicePrincipal(token, newSpId, {
      appRoleAssignmentRequired: true,
      preferredSingleSignOnMode: 'notSupported',
      loginUrl: null,
      logoutUrl: null,
      tags: ['WindowsAzureActiveDirectoryIntegratedApp', 'WindowsAzureActiveDirectoryOnPremApp', 'HideApp'],
    });
    yield { type: 'diag', message: 'SP properties configured.' };
  } catch (e) {
    yield { type: 'diag', message: `SP config non-fatal: ${e.message}` };
  }

  yield { type: 'progress', message: 'Granting admin consent...', percent: 40 };

  // Grant delegated scope for User.Read at minimum
  try {
    const msSp = await graphRequest(token, 'GET', `/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'&$select=id`);
    const msSpId = msSp?.value?.[0]?.id;
    if (msSpId) {
      const grants = await graphRequest(token, 'GET', `/oauth2PermissionGrants?$filter=clientId eq '${newSpId}' and resourceId eq '${msSpId}'`);
      if (!grants?.value?.length) {
        await graphRequest(token, 'POST', '/oauth2PermissionGrants', {
          clientId: newSpId, consentType: 'AllPrincipals', resourceId: msSpId, scope: 'User.Read',
        });
        yield { type: 'diag', message: 'Granted User.Read delegated scope.' };
      }
    }
  } catch (e) {
    yield { type: 'diag', message: `Consent non-fatal: ${e.message}` };
  }

  yield { type: 'progress', message: 'Setting up SCIM synchronization...', percent: 55 };

  // Wait for SP to be ready, then get sync template
  let templateId = 'scim';
  for (let attempt = 1; attempt <= 8; attempt++) {
    try {
      const templates = await getSyncTemplates(token, newSpId);
      const t = templates.find(t => t.id === 'scim') || templates.find(t => t.id?.toLowerCase().includes('scim')) || templates[0];
      if (t) { templateId = t.id; break; }
    } catch (_) {}
    if (attempt < 8) {
      yield { type: 'diag', message: `Waiting for SCIM template (attempt ${attempt}/8)...` };
      await new Promise(r => setTimeout(r, 3000));
    }
  }
  yield { type: 'diag', message: `Using sync template: ${templateId}` };

  yield { type: 'progress', message: 'Creating sync job...', percent: 65 };

  let syncJobId = null;
  try {
    const job = await createSyncJob(token, newSpId, templateId);
    syncJobId = job?.id;
    yield { type: 'diag', message: `Sync job created: ${syncJobId}` };
  } catch (e) {
    yield { type: 'diag', message: `Sync job creation non-fatal: ${e.message}` };
  }

  yield { type: 'progress', message: 'Configuring SCIM endpoint secrets...', percent: 75 };

  if (scimUrl && scimToken) {
    const secretsOk = await setSyncSecrets(token, newSpId, scimUrl, scimToken);
    yield { type: 'diag', message: secretsOk ? 'SCIM secrets configured.' : 'SCIM secrets config failed (may require manual setup).' };
  }

  yield { type: 'progress', message: 'Assigning users and groups...', percent: 82 };

  let assignedCount = 0, assignFailCount = 0;
  for (const p of principals) {
    try {
      await assignPrincipal(token, newSpId, p.id, p.type || 'User');
      assignedCount++;
      yield { type: 'diag', message: `Assigned: ${p.name}` };
    } catch (e) {
      assignFailCount++;
      yield { type: 'diag', message: `Assign failed for ${p.name}: ${e.message}` };
    }
  }

  yield { type: 'progress', message: 'Starting synchronization...', percent: 92 };

  if (syncJobId) {
    try {
      await startSyncJob(token, newSpId, syncJobId);
      yield { type: 'diag', message: 'Sync job started.' };
    } catch (e) {
      yield { type: 'diag', message: `Sync start non-fatal: ${e.message}` };
    }
  }

  yield { type: 'progress', message: 'Done!', percent: 100 };
  yield {
    type: 'result',
    success: true,
    appName,
    spId: newSpId,
    appObjId: newAppObjId,
    syncJobId,
    creationMode,
    assignedRequested: principals.length,
    assignedCount,
    assignFailCount,
  };
}

// ─── Deployments ──────────────────────────────────────────────────────────────

async function getOfficeDeployments(token) {
  try {
    const res = await graphRequest(token, 'GET', '/admin/officeConfiguration/clientConfigurations', null, { beta: true });
    return res?.value || [];
  } catch (e) {
    if (e.status === 404 || e.message.includes('officeConfiguration')) {
      return null; // not supported
    }
    throw e;
  }
}

module.exports = {
  graphRequest,
  graphGetAll,
  getClientCredentialsToken,
  getAuthorizationCodeToken,
  getOrganization,
  getMe,
  searchServicePrincipals,
  getServicePrincipal,
  updateServicePrincipal,
  deleteServicePrincipal,
  createServicePrincipal,
  getApplicationByAppId,
  createApplication,
  deleteApplication,
  updateApplicationTags,
  instantiateTemplate,
  getAppPermissions,
  getOAuth2Grants,
  getPrincipalAppRoleAssignments,
  grantAdminConsent,
  getAppRoleAssignments,
  enrichAssignments,
  assignPrincipal,
  removeAppRoleAssignment,
  searchPrincipals,
  getGroupMembers,
  getUserAppAssignments,
  getSpTagsDetailed,
  setVisibility,
  getSyncJobs,
  getSyncTemplates,
  createSyncJob,
  startSyncJob,
  restartSyncJob,
  setSyncSecrets,
  createScimApp,
  getOfficeDeployments,
  testScimConnection,
};
