# module/graph - V8 module

This module keeps Microsoft Graph calls simple inside Instago workflows. It bundles auth lookup and request helpers so step code stays small and readable.

If you are new to the platform, start at https://instago.ai.

## Technical reference

Exports:
- `configure({ tenant?, clientId?, clientSecret?, refreshToken?, scope?, baseUrl? })`
- `ensureAccessToken(overrides?)`
- `json({ path, method='GET', headers?, bodyObj?, debug?, auth? })`

Usage:
- `configure(...)` stores default auth fields and `baseUrl` (defaults to `https://graph.microsoft.com/v1.0`).
- `ensureAccessToken(overrides?)` checks `overrides`, configured values, then env keys to return an access token string (or an empty string).
- `json(...)` joins `baseUrl` + `path`, adds `Authorization` if missing, and returns `{ ok, data, status, raw, error }`.

Configuration:
- Optional env keys:
  - `graph.accessToken`
  - `graph.clientId`
  - `graph.clientSecret`
  - `graph.refreshToken`
  - `graph.scope`
  - `graph.tenant`
  - `ms.accessToken`
  - `ms.clientId`
  - `ms.clientSecret`
  - `ms.refreshToken`
  - `ms.scope`
  - `ms.tenant`
  - `msAccessToken`
  - `msClientId`
  - `msClientSecret`
  - `msRefreshToken`
  - `msScope`
  - `msTenant`

Examples:
```js
const graph = require('graph@latest');

graph.configure({
  tenant: env.graph.tenant,
  clientId: env.graph.clientId,
  clientSecret: env.graph.clientSecret,
  refreshToken: env.graph.refreshToken,
  scope: env.graph.scope
});

const res = await graph.json({ path: '/me', method: 'GET' });
```

```js
const graph = require('graph@latest');

const res = await graph.json({
  path: '/me/drive/root',
  auth: { accessToken: env.graph.accessToken }
});
```
