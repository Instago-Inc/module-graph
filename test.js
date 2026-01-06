const t = require('utest@1.0.0');
const graph = require('graph@1.0.0');

t.test('ensureAccessToken returns empty when unconfigured', async () => {
  const tok = await graph.ensureAccessToken({});
  t.expect(typeof tok === 'string').toBe(true);
  t.expect(tok.length === 0).toBe(true);
});

module.exports = { run: async (opts) => { await t.run(Object.assign({ quiet: true }, opts)); t.reset(); } };
