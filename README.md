# NDIS Invoice Management System

A participant invoicing platform built for the Witty Data assessment — covering participant/provider management, NDIS rate-set Excel import, and invoice management with rate/price matching.

## Tech Stack

- Node.js v24.18.0, Next.js v16.2.10, TypeScript
- PostgreSQL (via Kysely)
- antd, TailwindCSS

## Setup

### 1. Prerequisites

- Node.js v24.18.0
- Docker (for Postgres [+ MinIO if used]), or your own local instances

### 2. Install

```bash
npm install
```

### 3. Start infrastructure

```bash
docker compose up -d
```

### 4. Configure environment

```bash
cp .env.example .env
```

### 5. Run the app

```bash
npm run dev
```

Visit `http://localhost:3000`.

[Confirm you actually ran steps 1–5 in order against a clean database before submitting.]

## What's Implemented

### Primary Requirements (Section 9)

- [x] Participant (Client) Management
- [x] Provider Management
- [x] Invoice Management (create/edit/list/view, rate-set matching, price lookup, draft vs completed validation)
- [x] Rate Set Management
- [x] NDIS Excel Import (idempotent, all worksheets, categories/items/attributes/prices)

### Secondary Requirements (Section 10) — optional for junior candidates

- [ ] System User Management
- [ ] Authentication
- [ ] RBAC
- [ ] Audit Logging

### Extended: AI-based PDF Invoice Extraction (Section 11) — optional

- [ ] Invoice upload
- [ ] MinIO file storage
- [ ] AI extraction
- [ ] Draft creation + mapping cascade (participant/provider/rate-set/support-item)

## What's Incomplete

- **Section 10 (Secondary)** — User management, login/session (Argon2), RBAC permission checks, and audit logging are not implemented.
- **Section 11 (AI Extraction)** — PDF upload, MinIO storage, OpenAI/OpenRouter extraction, and auto draft-invoice creation are not implemented.

## Known Limitations

- No authentication — all API routes are currently open.
- No role-based permission checks.
- No audit trail for create/update/delete actions.
- Invoice AI upload flow is not available.

## If I Had More Time

1. **Authentication + RBAC**  
   Implement Argon2 login, session tokens, and enforce permissions on every write endpoint. This is the highest-priority missing piece for a usable system.

2. **Audit logging**  
   Record actor, action, entity, and changes into `audit_log` on every mutation.

3. **AI invoice extraction**  
   Add PDF upload → MinIO → OpenRouter/OpenAI structured extraction → draft invoice creation with participant/provider mapping.
