var currentVCard = '';

function val(id) { return document.getElementById(id).value.trim(); }

function toUTF8(str) { return unescape(encodeURIComponent(str)); }

function liveUpdate() {
    var name = val('fullName'), role = val('jobTitle'), phone = val('phone');
    var email = val('email'), company = val('company'), website = val('website');
    document.getElementById('cardCompany').textContent = company || 'Toyota';
    document.getElementById('cardName').textContent = name || 'Tu nombre';
    document.getElementById('cardRole').textContent = role || 'Cargo';
    function sd(wid, tid, v) { document.getElementById(tid).textContent = v; document.getElementById(wid).style.display = v ? 'flex' : 'none'; }
    sd('cardPhone', 'cardPhoneTxt', phone);
    sd('cardEmail', 'cardEmailTxt', email);
    sd('cardWeb', 'cardWebTxt', website);
}

function buildVCard() {
    var name = val('fullName'), role = val('jobTitle'), phone = val('phone');
    var email = val('email'), company = val('company'), website = val('website');
    var parts = name.split(' '), first = parts[0] || '', last = parts.slice(1).join(' ') || '';
    var v = 'BEGIN:VCARD\nVERSION:3.0\n';
    v += 'N:' + last + ';' + first + ';\n';
    v += 'FN:' + name + '\n';
    if (role) v += 'TITLE:' + role + '\n';
    if (company) v += 'ORG:' + company + '\n';
    if (phone) v += 'TEL;TYPE=CELL:' + phone + '\n';
    if (email) v += 'EMAIL:' + email + '\n';
    if (website) v += 'URL:' + website + '\n';
    v += 'END:VCARD';
    return v;
}

function generateQR() {
    var name = val('fullName'), phone = val('phone'), email = val('email');
    if (!name && !phone && !email) { alert('Ingresa al menos nombre, teléfono o correo.'); return; }
    currentVCard = buildVCard();

    var container = document.getElementById('qrcode');
    document.getElementById('cardQrPlaceholder').style.display = 'none';
    container.innerHTML = '';

    var qr = qrcode(0, 'L');
    qr.addData(toUTF8(currentVCard), 'Byte');
    qr.make();

    var count = qr.getModuleCount();
    var cell = 4;
    var canvas = document.createElement('canvas');
    canvas.width = count * cell;
    canvas.height = count * cell;
    var ctx = canvas.getContext('2d');
    for (var r = 0; r < count; r++) {
        for (var c = 0; c < count; c++) {
            ctx.fillStyle = qr.isDark(r, c) ? '#000000' : '#ffffff';
            ctx.fillRect(c * cell, r * cell, cell, cell);
        }
    }
    container.appendChild(canvas);
    document.getElementById('downloadBtn').style.display = 'block';
}

function downloadQR() {
    var canvas = document.querySelector('#qrcode canvas');
    if (!canvas) return;
    var link = document.createElement('a');
    link.download = 'QR_Toyota_' + (val('fullName').replace(/\s+/g, '_') || 'contacto') + '.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
}

document.addEventListener('DOMContentLoaded', liveUpdate);

// ── LOTE ──
var batchContacts = [], batchMapped = {}, batchCanvases = [];
var BATCH_ALIASES = {
    nombre: ['nombre y apellido', 'nombre', 'name', 'full name', 'nombre completo'],
    cargo: ['puesto', 'cargo', 'titulo', 'title', 'job title', 'position', 'rol'],
    telefono: ['celular corporativo', 'celular', 'móvil', 'movil', 'cell', 'mobile', 'cel'],
    telefono2: ['teléfono', 'telefono', 'tel', 'phone', 'fijo', 'fono'],
    email: ['correo', 'email', 'e-mail', 'mail'],
    empresa: ['tipo tarjeta', 'empresa', 'company'],
    web: ['web', 'website', 'url', 'sitio']
};

function bNorm(s) {
    return String(s).toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9 ]/g, '');
}
function bDetect(col) {
    var nc = bNorm(col);
    for (var f in BATCH_ALIASES) { if (BATCH_ALIASES[f].some(function (a) { return bNorm(a) === nc; })) return f; }
    for (var f in BATCH_ALIASES) { if (BATCH_ALIASES[f].some(function (a) { return nc.includes(bNorm(a)); })) return f; }
    return null;
}
function toggleBatch() {
    document.getElementById('batchToggle').classList.toggle('open');
    document.getElementById('batchPanel').classList.toggle('open');
}

function bSafeStr(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return String(v);
    return '';
}

function bHandleFile(file) {
    if (!file) return;
    var r = new FileReader();
    r.onload = function (e) {
        var wb;
        try {
            wb = XLSX.read(e.target.result, { type: 'array', cellStyles: false, cellNF: false, cellDates: false });
        } catch (err1) {
            try {
                wb = XLSX.read(e.target.result, { type: 'array', cellStyles: false, cellNF: false, cellDates: false, WTF: false });
            } catch (err2) {
                alert('No se pudo leer el archivo. Intenta guardarlo como .xlsx nuevo desde Excel y vuelve a subirlo.');
                return;
            }
        }
        var ws = wb.Sheets[wb.SheetNames[0]];
        var all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
        var hi = 0, mx = 0;
        for (var i = 0; i < Math.min(5, all.length); i++) {
            var t = all[i].filter(function (v) { return typeof v === 'string' && v.trim(); }).length;
            if (t > mx) { mx = t; hi = i; }
        }
        var headers = all[hi];
        batchMapped = {};
        headers.forEach(function (c) { if (c) { var f = bDetect(String(c)); if (f && !batchMapped[f]) batchMapped[f] = c; } });
        var rows = all.slice(hi + 1);
        batchContacts = rows.map(function (row) {
            var obj = {}; headers.forEach(function (h, i) { if (h) obj[h] = bSafeStr(row[i]); });
            return {
                nombre: (batchMapped.nombre ? bSafeStr(obj[batchMapped.nombre]) : '').trim(),
                cargo: (batchMapped.cargo ? bSafeStr(obj[batchMapped.cargo]) : '').trim(),
                telefono: (batchMapped.telefono ? bSafeStr(obj[batchMapped.telefono]) : batchMapped.telefono2 ? bSafeStr(obj[batchMapped.telefono2]) : '').trim(),
                email: (batchMapped.email ? bSafeStr(obj[batchMapped.email]) : '').trim(),
                empresa: (batchMapped.empresa ? bSafeStr(obj[batchMapped.empresa]) : 'Toyota').trim() || 'Toyota',
                web: (batchMapped.web ? bSafeStr(obj[batchMapped.web]) : '').trim(),
                sede: bSafeStr(obj['Sede']).trim()
            };
        }).filter(function (c) { return c.nombre || c.email || c.telefono; });

        var fm = document.getElementById('batchFmap'); fm.innerHTML = '';
        var labels = { nombre: 'Nombre', cargo: 'Cargo', telefono: 'Teléfono (móvil)', email: 'Email', empresa: 'Empresa', web: 'Web' };
        for (var f in labels) {
            var found = batchMapped[f];
            var d = document.createElement('div'); d.className = 'batch-fmap-row';
            var s1 = document.createElement('span'); s1.style.color = 'var(--muted)'; s1.textContent = labels[f];
            var s2 = document.createElement('span'); s2.className = found ? 'batch-fmap-ok' : 'batch-fmap-no'; s2.textContent = found ? '✓ ' + found : '— no detectado';
            d.appendChild(s1); d.appendChild(s2); fm.appendChild(d);
        }
        document.getElementById('batchCl').textContent = batchContacts.length + ' contactos listos';
        document.getElementById('batchMapWrap').style.display = 'block';
    };
    r.readAsArrayBuffer(file);
}

function bCleanPhone(p) {
    p = String(p).trim();
    var digits = p.replace(/[^\d+]/g, '');
    if (!digits) return '';
    if (!digits.startsWith('+')) digits = '+51' + digits.replace(/^51/, '');
    // Solo aceptar móviles: +519XXXXXXXX
    if (!/^\+519\d+$/.test(digits)) return '';
    return digits;
}

function bBuildVCard(c) {
    var empresaOverride = (document.getElementById('batchEmpresa').value || '').trim();
    var empresa = empresaOverride || c.empresa;
    var v = 'BEGIN:VCARD\nVERSION:3.0\n';
    if (c.nombre) { var p = c.nombre.split(' '); v += 'N:' + p.slice(1).join(' ') + ';' + p[0] + ';\n'; v += 'FN:' + c.nombre + '\n'; }
    if (c.cargo) v += 'TITLE:' + c.cargo + '\n';
    if (empresa) v += 'ORG:' + empresa + '\n';
    if (c.telefono) v += 'TEL;TYPE=CELL:' + bCleanPhone(c.telefono) + '\n';
    if (c.email) v += 'EMAIL:' + c.email + '\n';
    if (c.web) v += 'URL:' + c.web + '\n';
    v += 'END:VCARD';
    return v;
}

function bMakeQR(text) {
    return new Promise(function (res) {
        var qr = qrcode(0, 'Q');
        qr.addData(unescape(encodeURIComponent(text)), 'Byte');
        qr.make();
        var count = qr.getModuleCount(), cell = 10;
        var cv = document.createElement('canvas');
        cv.width = count * cell; cv.height = count * cell;
        var ctx = cv.getContext('2d');
        for (var row = 0; row < count; row++) {
            for (var col = 0; col < count; col++) {
                ctx.fillStyle = qr.isDark(row, col) ? '#000' : '#fff';
                ctx.fillRect(col * cell, row * cell, cell, cell);
            }
        }
        res(cv);
    });
}

async function batchGenerate() {
    document.getElementById('batchGenBtn').disabled = true;
    document.getElementById('batchPbWrap').style.display = 'block';
    batchCanvases = [];
    var grid = document.getElementById('batchGrid'); grid.innerHTML = '';
    for (var i = 0; i < batchContacts.length; i++) {
        var c = batchContacts[i];
        document.getElementById('batchPf').style.width = Math.round((i + 1) / batchContacts.length * 100) + '%';
        document.getElementById('batchSm').textContent = 'Generando ' + (i + 1) + ' de ' + batchContacts.length + '...';
        var cv = await bMakeQR(bBuildVCard(c));
        batchCanvases.push({ canvas: cv, name: c.nombre || ('contacto_' + (i + 1)) });
        var card = document.createElement('div'); card.className = 'batch-card';
        var sede = document.createElement('div'); sede.className = 'batch-card-sede';
        var empresaOverride = (document.getElementById('batchEmpresa').value || '').trim();
        sede.textContent = (empresaOverride || c.empresa || 'Toyota') + (c.sede ? ' · ' + c.sede : '');
        var nm = document.createElement('div'); nm.className = 'batch-card-name'; nm.textContent = c.nombre || '—';
        var mt = document.createElement('div'); mt.className = 'batch-card-meta';
        mt.textContent = (c.cargo ? c.cargo + ' · ' : '') + (c.email || '') + (c.telefono ? ' · ' + c.telefono : '');
        card.appendChild(sede); card.appendChild(nm); card.appendChild(mt);
        grid.appendChild(card);
        await new Promise(function (r) { setTimeout(r, 15); });
    }
    document.getElementById('batchSm').textContent = '✓ ' + batchContacts.length + ' QRs generados';
    document.getElementById('batchDlBtn').style.display = 'inline-block';
    document.getElementById('batchCl').textContent = batchContacts.length + ' QRs listos';
}

async function batchDownload() {
    var btn = document.getElementById('batchDlBtn');
    btn.disabled = true; btn.textContent = 'Preparando ZIP...';
    var zip = new JSZip();
    for (var i = 0; i < batchCanvases.length; i++) {
        var item = batchCanvases[i];
        var blob = await new Promise(function (resolve) { item.canvas.toBlob(resolve, 'image/png'); });
        zip.file(item.name.replace(/[^\wáéíóúüñÁÉÍÓÚÜÑ ]/g, '').trim() + '.png', blob);
    }
    var content = await zip.generateAsync({ type: 'blob' });
    var a = document.createElement('a'); a.href = URL.createObjectURL(content);
    a.download = 'Toyota_QR_lote.zip'; a.click();
    btn.disabled = false; btn.textContent = 'Descargar ZIP';
}

function batchReset() {
    batchContacts = []; batchCanvases = [];
    document.getElementById('batchFi').value = '';
    document.getElementById('batchMapWrap').style.display = 'none';
    document.getElementById('batchGrid').innerHTML = '';
    document.getElementById('batchPbWrap').style.display = 'none';
    document.getElementById('batchGenBtn').disabled = false;
    document.getElementById('batchDlBtn').style.display = 'none';
    document.getElementById('batchCl').textContent = '';
}

document.addEventListener('DOMContentLoaded', function () {
    var dz = document.getElementById('batchDz');
    dz.addEventListener('dragover', function (e) { e.preventDefault(); dz.classList.add('over'); });
    dz.addEventListener('dragleave', function () { dz.classList.remove('over'); });
    dz.addEventListener('drop', function (e) { e.preventDefault(); dz.classList.remove('over'); bHandleFile(e.dataTransfer.files[0]); });
    document.getElementById('batchFi').addEventListener('change', function (e) { bHandleFile(e.target.files[0]); });
});

