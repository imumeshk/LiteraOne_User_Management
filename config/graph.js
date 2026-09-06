'use strict';

module.exports = {
  GRAPH_BASE_V1: 'https://graph.microsoft.com/v1.0',
  GRAPH_BASE_BETA: 'https://graph.microsoft.com/beta',
  TOKEN_ENDPOINT: 'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token',
  SCOPE: 'https://graph.microsoft.com/.default',

  NON_GALLERY_TEMPLATE_ID: '8adf8e6e-67b2-4cf2-a259-e3dc5476c621',

  REQUIRED_PERMISSIONS: [
    'Directory.Read.All',
    'Application.ReadWrite.All',
    'AppRoleAssignment.ReadWrite.All',
    'Group.Read.All',
    'Synchronization.ReadWrite.All',
    'DelegatedPermissionGrant.ReadWrite.All',
  ],

  HIDE_TAGS: ['HideApp', 'WindowsAzureActiveDirectoryIntegratedApp'],
};
