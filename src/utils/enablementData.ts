import type { EnablementResource } from '../types/enablement'

export const SAMPLE_ENABLEMENT = [
  {
    id: 1, title: 'Getting Started with Litera One', type: 'video',
    description: 'A complete walkthrough of the Litera One platform covering core workflows and key features for new users.',
    url: 'https://litera.com/resources', icon: '🎬', tags: ['onboarding', 'intro', 'overview']
  },
  {
    id: 2, title: 'Litera Transact Quick Start', type: 'guide',
    description: 'Step-by-step guide to setting up and using Litera Transact for deal management and closing.',
    url: 'https://litera.com/resources', icon: '📘', tags: ['transact', 'guide', 'deals']
  },
  {
    id: 3, title: 'Compare & Review Feature Overview', type: 'video',
    description: 'Learn how to use document comparison and collaborative review workflows to accelerate deal processes.',
    url: 'https://litera.com/resources', icon: '🎬', tags: ['compare', 'review', 'documents']
  },
  {
    id: 4, title: 'Admin Configuration Manual', type: 'doc',
    description: 'Comprehensive reference for IT administrators configuring Litera One, including SCIM, SSO, and permissions.',
    url: 'https://litera.com/resources', icon: '📄', tags: ['admin', 'config', 'it']
  },
  {
    id: 5, title: 'Email Template Library', type: 'template',
    description: 'Pre-built email templates for common legal communications, deal notifications, and client correspondence.',
    url: 'https://litera.com/resources', icon: '📧', tags: ['templates', 'email', 'communication']
  },
  {
    id: 6, title: 'Outlook Add-in User Guide', type: 'guide',
    description: 'End-user guide for the Litera One Outlook integration — accessing matters, documents, and contacts from email.',
    url: 'https://litera.com/resources', icon: '📘', tags: ['outlook', 'addin', 'email']
  },
  {
    id: 7, title: 'Word Add-in Walkthrough', type: 'video',
    description: 'Video tutorial covering the Litera One Word add-in for drafting, comparison, and clause management.',
    url: 'https://litera.com/resources', icon: '🎬', tags: ['word', 'addin', 'drafting']
  },
  {
    id: 8, title: 'Matter & Deal Templates', type: 'template',
    description: 'Standard matter management and deal closing checklists and templates for your legal team.',
    url: 'https://litera.com/resources', icon: '🗂', tags: ['templates', 'matter', 'checklist']
  },
  {
    id: 9, title: 'SCIM Provisioning Setup Guide', type: 'doc',
    description: 'Technical guide for configuring SCIM user provisioning with Microsoft Entra ID / Azure AD.',
    url: 'https://litera.com/resources', icon: '📄', tags: ['scim', 'admin', 'azure', 'provisioning']
  },
  {
    id: 10, title: 'End-User Adoption Playbook', type: 'guide',
    description: 'Proven strategies and communication templates for rolling out Litera One across your organization.',
    url: 'https://litera.com/resources', icon: '📘', tags: ['adoption', 'change management', 'rollout']
  },
  {
    id: 11, title: 'Advanced Search & Drafting', type: 'video',
    description: 'Deep-dive into Litera Precedent for clause reuse, advanced search, and AI-assisted drafting.',
    url: 'https://litera.com/resources', icon: '🎬', tags: ['drafting', 'search', 'precedent', 'ai']
  },
  {
    id: 12, title: 'Billing & Matter Intake Templates', type: 'template',
    description: 'Ready-to-use templates for billing narratives, matter intake forms, and time entry descriptions.',
    url: 'https://litera.com/resources', icon: '🗂', tags: ['billing', 'matter', 'intake', 'templates']
  },
] as EnablementResource[]

export const TYPE_LABELS = {
  All: 'All Resources',
  video: 'Videos',
  guide: 'Guides',
  doc: 'Documents',
  template: 'Templates'
} as const
