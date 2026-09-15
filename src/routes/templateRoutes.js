const express = require('express');
const { protect, authorize } = require('../middleware/auth');
const upload = require('../middleware/upload');
const makeTemplateController = require('../controllers/templateControllerFactory');
const PPTTemplate = require('../models/PPTTemplate');
const ExcelTemplate = require('../models/ExcelTemplate');

function buildRouter(Model) {
  const router = express.Router();
  const ctrl = makeTemplateController(Model);

  router.use(protect);

  router.get('/', ctrl.getAll);
  router.get('/:id', ctrl.getOne);
  router.post('/', authorize('admin', 'tl'), upload.single('file'), ctrl.create);
  router.put('/:id', authorize('admin', 'tl'), ctrl.update);
  router.patch('/:id/status', authorize('admin', 'tl'), ctrl.setStatus);
  router.delete('/:id', authorize('admin', 'tl'), ctrl.remove);

  return router;
}

module.exports = {
  pptRouter: buildRouter(PPTTemplate),
  excelRouter: buildRouter(ExcelTemplate),
};
