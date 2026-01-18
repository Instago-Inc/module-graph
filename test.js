const t = require('utest@latest');
const graph = require('graph@latest');

t.test('ensureAccessToken returns empty when unconfigured', async () => {
  const tok = await graph.ensureAccessToken({});
  t.expect(typeof tok === 'string').toBe(true);
  t.expect(tok.length === 0).toBe(true);
});

module.exports = { run: async (opts) => { await t.run(Object.assign({ quiet: true }, opts)); t.reset(); } };
