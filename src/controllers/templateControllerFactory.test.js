import { describe, it, expect, vi, beforeEach } from 'vitest';
import { createRequire } from 'module';

// templateControllerFactory.js is loaded via Node's native require (not Vite's ESM
// transform) because its existing `module.exports = makeTemplateFactory = ...` line
// relies on CommonJS sloppy-mode implicit globals, which throws under strict-mode ESM.
const require = createRequire(import.meta.url);
const makeTemplateController = require('./templateControllerFactory.js');

function mockRes() {
  const res = {};
  res.status = vi.fn().mockReturnValue(res);
  res.json = vi.fn().mockReturnValue(res);
  return res;
}

describe('templateControllerFactory - PPT Master create (minimal fields)', () => {
  let Model;
  let ctrl;

  beforeEach(() => {
    Model = {
      create: vi.fn(async (payload) => ({ _id: 'tpl1', ...payload })),
      find: vi.fn(() => ({ sort: vi.fn(async () => []) })),
      findById: vi.fn(async () => null),
    };
    ctrl = makeTemplateController(Model);
  });

  it('creates a template using only name, description and status', async () => {
    const req = { body: { name: 'Q1 Proposal', description: 'Q1 pitch deck', status: 'active' }, user: { _id: 'u1' }, file: undefined };
    const res = mockRes();

    await ctrl.create(req, res);

    expect(Model.create).toHaveBeenCalledWith(
      expect.objectContaining({ name: 'Q1 Proposal', description: 'Q1 pitch deck', status: 'active', createdBy: 'u1' })
    );
    const sentPayload = Model.create.mock.calls[0][0];
    expect(sentPayload.version).toBeUndefined();
    expect(sentPayload.variant).toBeUndefined();
    expect(sentPayload.fileUrl).toBeUndefined();
    expect(res.status).toHaveBeenCalledWith(201);
  });

  it('does not require version, variant or file in the request body', async () => {
    const req = { body: { name: 'No Extras' }, user: { _id: 'u1' }, file: undefined };
    const res = mockRes();

    await expect(ctrl.create(req, res)).resolves.not.toThrow();
    expect(Model.create).toHaveBeenCalled();
  });

  it('still forwards version/variant when the caller sends them (edit / legacy flows unaffected)', async () => {
    const req = {
      body: { name: 'Legacy Upload', version: '2.0', variant: 'Premium' },
      user: { _id: 'u1' },
      file: undefined,
    };
    const res = mockRes();

    await ctrl.create(req, res);

    expect(Model.create).toHaveBeenCalledWith(expect.objectContaining({ version: '2.0', variant: 'Premium' }));
  });
});

describe('templateControllerFactory - existing template APIs unaffected', () => {
  let Model;
  let ctrl;

  beforeEach(() => {
    Model = {
      create: vi.fn(async (payload) => ({ _id: 'tpl1', ...payload })),
      find: vi.fn(() => ({ sort: vi.fn(async () => [{ _id: 'tpl1', name: 'Existing' }]) })),
      findById: vi.fn(async (id) => ({
        _id: id,
        name: 'Existing',
        status: 'active',
        save: vi.fn(async function () {
          return this;
        }),
        deleteOne: vi.fn(async () => {}),
      })),
    };
    ctrl = makeTemplateController(Model);
  });

  it('getAll lists templates', async () => {
    const req = { query: {} };
    const res = mockRes();
    await ctrl.getAll(req, res);
    expect(res.json).toHaveBeenCalledWith([{ _id: 'tpl1', name: 'Existing' }]);
  });

  it('getOne returns a single template', async () => {
    const req = { params: { id: 'tpl1' } };
    const res = mockRes();
    await ctrl.getOne(req, res);
    expect(res.json).toHaveBeenCalled();
  });

  it('setStatus toggles status on an existing template', async () => {
    const req = { params: { id: 'tpl1' }, body: { status: 'inactive' } };
    const res = mockRes();
    await ctrl.setStatus(req, res);
    expect(res.json).toHaveBeenCalled();
  });

  it('remove deletes an existing template', async () => {
    const req = { params: { id: 'tpl1' } };
    const res = mockRes();
    await ctrl.remove(req, res);
    expect(res.json).toHaveBeenCalledWith({ message: 'Template deleted' });
  });
});
