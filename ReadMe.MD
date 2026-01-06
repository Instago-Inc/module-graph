# module/graph - V8 module

This module is a compact helper for calling Microsoft Graph in Instago workflows. It keeps token handling and request setup in one place so flows can stay small and focused.

Learn more about the platform at https://instago.ai.

## Technical reference

Exports:
- `configure({ tenant?, clientId?, clientSecret?, refreshToken?, scope?, baseUrl? })`
- `ensureAccessToken(overrides?)`
- `json({ path, method='GET', headers?, bodyObj?, debug?, auth? })`

Usage:
- `configure(...)` stores defaults for auth and `baseUrl` (`https://graph.microsoft.com/v1.0` by default).
- `ensureAccessToken(overrides?)` returns a token from `overrides`, configured values, or env keys.
- `json(...)` performs a Graph request, adds `Authorization` when missing, and returns `{ ok, data, status, raw, error }`.

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
  refreshToken: env.graph.refreshToken
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
