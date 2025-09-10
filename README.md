<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>TargetFill: One Click, All Cells</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    body { font-family: Inter, system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial; }
    .app-container { max-width: 980px; margin: 2rem auto; padding: 2rem; }
    .file-drop-zone { border: 2px dashed #e2e8f0; padding: 2.5rem; border-radius: 0.75rem; cursor: pointer; }
    .file-drop-zone.dragover { background: #eff6ff; border-color: #60a5fa; }
    .btn-primary { background: #2563eb; color: white; }
    .btn-primary:disabled { background: #94a3b8; cursor: not-allowed; }
    .small { font-size: 0.9rem; }
    pre.debug { max-height: 200px; overflow: auto; background: #0f172a; color: #e2e8f0; padding: 0.5rem; border-radius: 6px; }
  </style>
</head>
<body class="bg-slate-50 text-slate-800">
  <div class="app-container bg-white rounded-xl shadow-md">
    <header class="text-center mb-6">
      <h1 class="text-2xl font-semibold">TargetFill: One Click, All Cells</h1>
      <p class="text-slate-500 small">Analyze Title/Bullets/Description in the <strong>data</strong> sheet and auto-fill attributes starting at column 10.</p>
    </header>

    <main>
      <section class="mb-6 p-4 border border-slate-200 rounded">
        <h2 class="font-medium mb-2">1. Category & Attributes</h2>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label class="small block mb-1">Category</label>
            <select id="category" class="w-full p-2 border rounded">
              <option value="dietary-supplement">Dietary Supplement</option>
              <option value="vitamins">Vitamins</option>
              <option value="more" disabled>More (Coming Soon)</option>
            </select>
          </div>

          <div>
            <label class="small block mb-1">Attributes to Fill</label>
            <div class="space-y-2">
              <label class="flex items-center space-x-2"><input type="checkbox" value="subtype"> <span>Health and Beauty Subtype*</span></label>
              <label class="flex items-center space-x-2"><input type="checkbox" value="gender"> <span>Gender*</span></label>
              <label class="flex items-center space-x-2"><input type="checkbox" value="application"> <span>Health Application*</span></label>
              <label class="flex items-center space-x-2"><input type="checkbox" value="form"> <span>Product Form*</span></label>
            </div>
          </div>
        </div>
      </section>

      <section class="mb-6 p-4 border border-slate-200 rounded">
        <h2 class="font-medium mb-2">2. Upload Excel file with Data</h2>
        <div id="dropZone" class="file-drop-zone text-center">
          <p id="dropZoneText" class="text-slate-500">Drag & drop your .xlsx or .xlsm file here, or click to select a file.</p>
          <input id="fileInput" type="file" accept=".xlsx,.xlsm" style="display:none">
          <div id="fileName" class="mt-2 text-blue-600 font-medium"></div>
        </div>
        <p class="small mt-3 text-slate-500">Expected sheet name: <strong>data</strong> — Row 1 headers: SKU | Title | Bullet1..Bullet5 | Description; leave Column 9 blank; attributes are written from Column 10 onward.</p>
      </section>

      <div class="text-center mb-6">
        <button id="generateBtn" class="btn-primary px-6 py-2 rounded font-medium" disabled>Analyze & Fill</button>
      </div>

      <div id="status" class="text-center small text-slate-600 mb-6"></div>

      <section class="p-4 border border-slate-200 rounded bg-slate-50">
        <h3 class="font-medium mb-2">File Format Instructions</h3>
        <ul class="list-disc list-inside small text-slate-600">
          <li>Sheet name must be <strong>data</strong></li>
          <li>Row 1 = headers</li>
          <li>Column 1: SKU</li>
          <li>Column 2: Title</li>
          <li>Column 3–7: Bullet Points (1–5)</li>
          <li>Column 8: Product Description</li>
          <li>Column 9: (keep blank)</li>
          <li>Column 10 onward: attributes filled from selected checkboxes</li>
        </ul>
      </section>

    </main>
  </div>

  <script>
    // --- Robust keyword maps (rules list) ---
    const kw = {
      subtype: [
        { terms: ['fiber','psyllium','inulin'], value: 'Fiber Supplements' },
        { terms: ['green','superfood','spirulina','chlorella','wheatgrass'], value: 'Greens and Superfood Supplements' },
        { terms: ['herb','herbal','ashwagandha','turmeric','ginseng','echinacea','milk thistle','valerian','fenugreek','ginger','garlic'], value: 'Herbal Supplements' },
        { terms: ['magnesium','zinc','iron','calcium','potassium','selenium','iodine','chromium','copper','manganese'], value: 'Mineral Supplements' },
        { terms: ['supplement','dietary'], value: 'Dietary Supplements' }
      ],
      gender: [
        { terms: ["men's", 'for men', 'for him', 'male'], value: 'Men' },
        { terms: ["women's", 'for women', 'for her', 'female'], value: 'Women' }
      ],
      application: [
        { terms: ['adrenal'], value: 'Adrenal Health' },
        { terms: ['aging','anti-aging','anti aging'], value: 'Aging' },
        { terms: ['allergy','allergies','hay fever'], value: 'Allergies' },
        { terms: ['anxiety','anxious'], value: 'Anxiety' },
        { terms: ['bladder infection','urinary tract infection','uti'], value: 'Bladder Infection' },
        { terms: ['bladder support'], value: 'Bladder Support' },
        { terms: ['bloat','bloating'], value: 'Bloating' },
        { terms: ['blood sugar','glycemic','glucose'], value: 'Blood Sugar Imbalance' },
        { terms: ['bone','osteoporosis','calcium'], value: 'Bone Health' },
        { terms: ['child','kids','children'], value: "Children's Health" },
        { terms: ['cholesterol'], value: 'Cholesterol Level Maintenance' },
        { terms: ['circulat','circulatory'], value: 'Circulatory System Health' },
        { terms: ['constipation'], value: 'Constipation' },
        { terms: ['dental','teeth','tooth'], value: 'Dental Health' },
        { terms: ['diabetes','diabetic'], value: 'Diabetes' },
        { terms: ['diarrhea','diarrhoea'], value: 'Diarrhea' },
        { terms: ['digest','digestion','probiotic','prebiotic','gut'], value: 'Digestive Health' },
        { terms: ['endurance'], value: 'Endurance' },
        { terms: ['energy','energiz'], value: 'Energy' },
        { terms: ['eye','vision','lutein','zeaxanthin'], value: 'Eye Health' },
        { terms: ['fertility','ttc','ovulation'], value: 'Fertility' },
        { terms: ['gout'], value: 'Gout' },
        { terms: ['hair','skin','nail','collagen','biotin'], value: 'Hair, Skin and Nail Health' },
        { terms: ['heart','cardio','omega-3','omega 3','epa','dha'], value: 'Heart Health' },
        { terms: ['hydrate','hydration','electrolyte'], value: 'Hydration' },
        { terms: ['immune','immunity','elderberry','vitamin c'], value: 'Immune System Health' },
        { terms: ['inflammation','anti-inflammatory','turmeric','curcumin'], value: 'Inflammation' },
        { terms: ['insomnia','sleep','melatonin'], value: 'Insomnia' },
        { terms: ['iron deficiency','anemia','anaemia'], value: 'Iron Deficiency' },
        { terms: ['ibs','irritable bowel'], value: 'Irritable Bowel Syndrome (IBS)' },
        { terms: ['joint','glucosamine','chondroitin','msm'], value: 'Joint Support' },
        { terms: ['liver','silymarin','milk thistle'], value: 'Liver Health' },
        { terms: ['memory','brain','focus','nootropic'], value: 'Memory and Brain Health' },
        { terms: ['menopause'], value: 'Menopause' },
        { terms: ['metabolism','thermogenic'], value: 'Metabolism' },
        { terms: ['mood','serotonin','5-htp'], value: 'Mood' },
        { terms: ['muscle'], value: 'Muscle Growth' },
        { terms: ['nausea'], value: 'Nausea' },
        { terms: ['nerve','neuropathy'], value: 'Nerve Pain' },
        { terms: ['overall','multivitamin','daily multivitamin','multivitamin'], value: 'Overall Health' },
        { terms: ['pain relief','analgesic'], value: 'Pain Relief' },
        { terms: ['pms'], value: 'PMS' },
        { terms: ['pregnan','prenatal','folic acid','folate'], value: 'Pregnancy' },
        { terms: ['prenatal'], value: 'Prenatal Health' },
        { terms: ['prostate','saw palmetto'], value: 'Prostate Health' },
        { terms: ['respiratory','lung','bronchial'], value: 'Respiratory Health' },
        { terms: ['seasonal allergy','seasonal allergies'], value: 'Seasonal Allergies' },
        { terms: ['skin','dermat'], value: 'Skin Health' },
        { terms: ['sleep disturbance'], value: 'Sleep Disturbance' },
        { terms: ['sports','athlet'], value: 'Sports Performance' },
        { terms: ['stress','adaptogen','ashwagandha','rhodiola'], value: 'Stress' },
        { terms: ['thyroid'], value: 'Thyroid Health' },
        { terms: ['urinary','uti','urinary tract infection'], value: 'Urinary Tract Infection' },
        { terms: ['vaginal'], value: 'Vaginal Health' },
        { terms: ['weight loss','fat burner'], value: 'Weight Loss' },
        { terms: ['weight management'], value: 'Weight Management' },
        { terms: ['women','female'], value: "Women's Health" },
        { terms: ['yeast','candida'], value: 'Yeast Infection' }
      ],
      form: [
        { terms: ['gummi','gummies','gummy'], value: 'Gummy' },
        { terms: ['softgel','soft gel'], value: 'Softgel' },
        { terms: ['capsule','caps'], value: 'Capsule' },
        { terms: ['tablet','tab'], value: 'Tablet' },
        { terms: ['chewable'], value: 'Chewable' },
        { terms: ['gelcap','gel cap'], value: 'Gelcap' },
        { terms: ['caplet'], value: 'Caplet' },
        { terms: ['powder','powdered'], value: 'Powder' },
        { terms: ['liquid','syrup','drops','tincture'], value: 'Liquid' },
        { terms: ['cream','topical'], value: 'Cream' },
        { terms: ['patch'], value: 'Patch' },
        { terms: ['lozenge'], value: 'Lozenge' },
        { terms: ['tea'], value: 'Tea' },
        { terms: ['bar'], value: 'Bar' },
        { terms: ['gum'], value: 'Gum' },
        { terms: ['lollipop'], value: 'Lollipop' }
      ]
    };

    // --- Helpers ---
    function norm(text) {
      return (text || '').toString().toLowerCase();
    }

    function pickValuesFromRules(text, rules, limit) {
      const found = [];
      const seen = new Set();
      const t = norm(text);
      for (const rule of rules) {
        for (const term of rule.terms) {
          if (t.includes(term) && !seen.has(rule.value)) {
            found.push(rule.value);
            seen.add(rule.value);
            break;
          }
        }
        if (found.length >= limit) break;
      }
      return found;
    }

    function analyzeRowText(text) {
      // returns object for each field
      const subtypeFound = pickValuesFromRules(text, kw.subtype, 3);
      const appFound = pickValuesFromRules(text, kw.application, 7);
      const formFound = pickValuesFromRules(text, kw.form, 1);
      const genderFound = pickValuesFromRules(text, kw.gender, 1);

      return {
        subtype: subtypeFound.length ? subtypeFound.join('|') : 'Dietary Supplements',
        application: appFound.length ? appFound.join('|') : 'Overall Health',
        form: formFound.length ? formFound[0] : 'Capsule',
        gender: genderFound.length ? genderFound[0] : 'Gender Neutral'
      };
    }

    // --- UI wiring ---
    const fileInput = document.getElementById('fileInput');
    const dropZone = document.getElementById('dropZone');
    const fileName = document.getElementById('fileName');
    const generateBtn = document.getElementById('generateBtn');
    const statusDiv = document.getElementById('status');

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', e => { e.preventDefault(); dropZone.classList.remove('dragover'); });
    dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('dragover'); if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]); });

    fileInput.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });

    let uploadedFile = null;
    function handleFile(f) {
      uploadedFile = f;
      fileName.textContent = `Selected: ${f.name}`;
      generateBtn.disabled = false;
      statusDiv.textContent = '';
    }

    function setStatus(msg, type = 'info') {
      statusDiv.textContent = msg;
      statusDiv.className = type === 'error' ? 'text-red-600 small text-center' : 'text-slate-600 small text-center';
    }

    // --- Main processing ---
    generateBtn.addEventListener('click', async () => {
      if (!uploadedFile) return setStatus('Please upload a file first.', 'error');

      const checked = Array.from(document.querySelectorAll('input[type=checkbox]:checked')).map(i => i.value);
      if (!checked.length) return setStatus('Please select at least one attribute to fill.', 'error');

      generateBtn.disabled = true; generateBtn.textContent = 'Processing...'; setStatus('Reading workbook...');

      try {
        const data = await uploadedFile.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });

        // find sheet named 'data' case-insensitive
        const sheetName = Object.keys(wb.Sheets).find(n => n.toLowerCase() === 'data');
        if (!sheetName) throw new Error('Sheet named "data" not found (case-insensitive).');
        const ws = wb.Sheets[sheetName];

        // ensure sheet ref exists
        let range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } };

        // prepare header labels
        const headerLabels = { subtype: 'Health and beauty subtype*', gender: 'Gender*', application: 'Health Application*', form: 'Product Form*' };

        // Start writing headers at column index 9 (Excel col 10)
        const startCol = 9;
        checked.forEach((key, i) => {
          const addr = XLSX.utils.encode_cell({ r: 0, c: startCol + i });
          ws[addr] = { v: headerLabels[key] || key, t: 's' };
        });

        // update range end col
        range.e.c = Math.max(range.e.c, startCol + checked.length - 1);

        // iterate rows starting from row index 1 (row 2 in excel) but we'll check until range.e.r; if file has data beyond !ref then it's unsafe, but this handles common cases
        for (let r = 1; r <= range.e.r; r++) {
          // read SKU and Title
          const skuAddr = XLSX.utils.encode_cell({ r, c: 0 });
          const titleAddr = XLSX.utils.encode_cell({ r, c: 1 });
          const sku = ws[skuAddr] ? (ws[skuAddr].v || '').toString().trim() : '';
          const title = ws[titleAddr] ? (ws[titleAddr].v || '').toString().trim() : '';
          if (!sku && !title) continue; // skip empty rows

          // collect text from Title (col1), bullets (cols 2..6), desc (col7)
          let textParts = [];
          for (let c = 1; c <= 7; c++) {
            const a = XLSX.utils.encode_cell({ r, c });
            if (ws[a] && ws[a].v) textParts.push(String(ws[a].v));
          }
          const blob = textParts.join(' \n ').toLowerCase();

          // analyze
          const result = analyzeRowText(blob);

          // write each selected attr in correct column starting at startCol
          checked.forEach((key, i) => {
            const val = result[key] !== undefined ? result[key] : '';
            const addr = XLSX.utils.encode_cell({ r, c: startCol + i });
            // ensure we don't write into column 9 (index 8) — startCol is 9 so safe
            ws[addr] = { v: val, t: 's' };
          });
        }

        // update worksheet ref
        ws['!ref'] = XLSX.utils.encode_range(range);

        // write out and download (preserve macro extension if present)
        const isXlsm = uploadedFile.name.toLowerCase().endsWith('.xlsm');
        const out = XLSX.write(wb, { bookType: isXlsm ? 'xlsm' : 'xlsx', type: 'array', bookVBA: isXlsm });
        const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const base = uploadedFile.name.replace(/\.xlsm$|\.xlsx$/i, '');
        a.download = `FILLED_${base}${isXlsm ? '.xlsm' : '.xlsx'}`;
        document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);

        setStatus('Done — file generated and download started.');
      } catch (err) {
        console.error(err);
        setStatus(`Error: ${err.message}`, 'error');
      } finally {
        generateBtn.disabled = false; generateBtn.textContent = 'Analyze & Fill';
      }
    });
  </script>
</body>
</html>
