const { Readable } = require('stream');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// تحميل سكربت التحويل كـ string
const scriptPath = path.join(__dirname, '..', 'convert_html_to_docx.js');
const scriptContent = fs.readFileSync(scriptPath, 'utf8');

// عمل sandbox لتشغيل الكود
const sandbox = {
  window: {},
  document: {},
  console,
  setTimeout,
  JSZip: require('../jszip-master/dist/jszip.min.js'), // أو استخدم نسخة داخلية
  saveAs: () => {},
};
vm.createContext(sandbox);
vm.runInContext(scriptContent, sandbox);

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method Not Allowed' });
  }

  const { html } = req.body;

  if (!html) {
    return res.status(400).json({ error: 'Missing HTML content' });
  }

  try {
    const convertFn = sandbox.window.convertHTMLToDOCX || sandbox.convertHTMLToDOCX;


    const blob = await convertFn(html, {
      orientation: "portrait",
      margins: { top: 720, right: 720, bottom: 720, left: 720 }
    });

    const buffer = Buffer.from(await blob.arrayBuffer());

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=converted.docx');
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.toString() });
  }
};
