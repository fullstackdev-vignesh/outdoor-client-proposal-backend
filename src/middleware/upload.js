const multer = require('multer');

const ALLOWED_TYPES = [
  'image/jpeg',
  'image/png',
  'image/webp',
  'image/gif',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/ms-excel',
  'application/vnd.ms-excel',
];
const MAX_SIZE_BYTES = 20 * 1024 * 1024; // 20MB

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_SIZE_BYTES },
  fileFilter: (req, file, cb) => {
    if (!ALLOWED_TYPES.includes(file.mimetype) && !/\.(pptx|xlsx|xls|jpg|jpeg|png|webp|gif)$/i.test(file.originalname)) {
      return cb(new Error('Only images, PPTX, or XLSX files are allowed'));
    }
    cb(null, true);
  },
});

module.exports = upload;
