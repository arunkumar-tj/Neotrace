/* === Traceability System Application === */

(function () {
    'use strict';

    var STORAGE_KEY = 'traceability_records';
    var ADMIN_PW_KEY = 'traceability_admin_pw';
    var PRODUCTS_KEY = 'traceability_products';
    var SHEET_WEBHOOK_KEY = 'traceability_sheet_webhook';
    var QC_RECORDS_KEY = 'neotrace_qc_records';
    var SHEET_ID = '1RAzOiM_NX9wYqrkRdVBo0Ohm5sKkMGtnTE8ibwjcc3E';
    var SHEET_GID = '212071646';
    var DEFAULT_ADMIN_PW = 'admin123';

    var DEFAULT_PRODUCTS = [
        { id: 'neofly', name: 'Neofly', fields: [] },
        { id: 'neobolt', name: 'Neobolt', fields: [
            { key: 'batteryNo', label: 'Battery Serial No' },
            { key: 'chargerNo', label: 'Battery Charger Serial No' },
            { key: 'motorNo', label: 'Motor Serial No' }
        ]}
    ];

    // ===========================
    // Data Layer
    // ===========================
    function getRecords() {
        try { var d = localStorage.getItem(STORAGE_KEY); return d ? JSON.parse(d) : []; }
        catch (e) { return []; }
    }
    function saveRecords(records) { localStorage.setItem(STORAGE_KEY, JSON.stringify(records)); }

    function addRecord(record) {
        var records = getRecords();
        record.id = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
        record.createdAt = new Date().toISOString();
        records.push(record);
        saveRecords(records);
        syncRecordToGoogleSheet('create', record);
        return record;
    }

    function updateRecord(id, updates) {
        var records = getRecords();
        for (var i = 0; i < records.length; i++) {
            if (records[i].id === id) {
                for (var key in updates) {
                    if (updates.hasOwnProperty(key)) records[i][key] = updates[key];
                }
                records[i].updatedAt = new Date().toISOString();
                saveRecords(records);
                syncRecordToGoogleSheet('update', records[i]);
                return records[i];
            }
        }
        return null;
    }

    function deleteRecord(id) {
        var records = getRecords();
        var removed = null;
        records.forEach(function (r) { if (r.id === id) removed = r; });
        saveRecords(records.filter(function (r) { return r.id !== id; }));
        if (removed) syncRecordToGoogleSheet('delete', removed);
    }

    // QC Records local storage
    function getQCRecords() {
        try { var d = localStorage.getItem(QC_RECORDS_KEY); return d ? JSON.parse(d) : []; }
        catch (e) { return []; }
    }
    function saveQCRecord(payload) {
        var records = getQCRecords();
        records.push({
            id: payload.id || ('qc_' + Date.now()),
            template: payload.template || '',
            templateTitle: payload.templateTitle || '',
            orderNo: payload.orderNo || '',
            inspector: payload.inspector || '',
            qcDate: payload.qcDate || '',
            submittedAt: payload.submittedAt || new Date().toISOString()
        });
        localStorage.setItem(QC_RECORDS_KEY, JSON.stringify(records));
    }

    function getAdminPassword() { return localStorage.getItem(ADMIN_PW_KEY) || DEFAULT_ADMIN_PW; }
    function setAdminPassword(pw) { localStorage.setItem(ADMIN_PW_KEY, pw); }

    function getSheetWebhookUrl() { return localStorage.getItem(SHEET_WEBHOOK_KEY) || ''; }
    function setSheetWebhookUrl(url) {
        var u = String(url || '').trim();
        if (u) localStorage.setItem(SHEET_WEBHOOK_KEY, u);
        else localStorage.removeItem(SHEET_WEBHOOK_KEY);
    }

    function postSheetPayload(payload) {
        var url = getSheetWebhookUrl();
        if (!url || typeof fetch !== 'function') return Promise.resolve(false);

        // Use urlencoded payload so Apps Script reliably receives data on all browsers.
        return fetch(url, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
            body: 'payload=' + encodeURIComponent(JSON.stringify(payload))
        }).then(function () {
            return true;
        }).catch(function (err) {
            console.warn('Google Sheet sync failed:', err);
            return false;
        });
    }

    function syncRecordToGoogleSheet(action, record) {
        var payload = {
            action: action,
            sheetId: SHEET_ID,
            gid: SHEET_GID,
            syncedAt: new Date().toISOString(),
            record: record || null
        };
        postSheetPayload(payload);
    }

    function syncAllRecordsToGoogleSheet() {
        var records = getRecords();
        if (records.length === 0) return Promise.resolve({ total: 0, sent: 0 });

        var jobs = records.map(function (record) {
            var payload = {
                action: 'create',
                sheetId: SHEET_ID,
                gid: SHEET_GID,
                syncedAt: new Date().toISOString(),
                record: record
            };
            return postSheetPayload(payload);
        });

        return Promise.all(jobs).then(function (results) {
            var sent = results.filter(function (ok) { return !!ok; }).length;
            return { total: records.length, sent: sent };
        });
    }

    // ===========================
    // Products Data Layer
    // ===========================
    function migrateProduct(p) {
        if (p.fields && p.fields.length > 0 && typeof p.fields[0] === 'string') {
            var L = { batteryNo: 'Battery Serial No', chargerNo: 'Battery Charger Serial No', motorNo: 'Motor Serial No' };
            p.fields = p.fields.map(function (f) { return { key: f, label: L[f] || f }; });
        }
        return p;
    }

    function getProducts() {
        try {
            var d = localStorage.getItem(PRODUCTS_KEY);
            if (d) {
                var products = JSON.parse(d).map(migrateProduct);
                saveProducts(products);
                return products;
            }
        } catch (e) {}
        saveProducts(DEFAULT_PRODUCTS);
        return DEFAULT_PRODUCTS.slice();
    }

    function saveProducts(products) { localStorage.setItem(PRODUCTS_KEY, JSON.stringify(products)); }

    function addProductToStore(product) {
        var products = getProducts();
        product.id = product.name.toLowerCase().replace(/[^a-z0-9]+/g, '_');
        var base = product.id, c = 1;
        while (products.some(function (p) { return p.id === product.id; })) { product.id = base + '_' + c++; }
        products.push(product);
        saveProducts(products);
        return product;
    }

    function updateProduct(pid, updates) {
        var products = getProducts();
        for (var i = 0; i < products.length; i++) {
            if (products[i].id === pid) {
                for (var k in updates) { if (updates.hasOwnProperty(k)) products[i][k] = updates[k]; }
                saveProducts(products);
                return products[i];
            }
        }
        return null;
    }

    function deleteProduct(pid) {
        saveProducts(getProducts().filter(function (p) { return p.id !== pid; }));
    }

    function getProductById(pid) {
        var products = getProducts();
        for (var i = 0; i < products.length; i++) { if (products[i].id === pid) return products[i]; }
        return null;
    }

    // ===========================
    // Utility
    // ===========================
    function formatDate(isoStr) {
        if (!isoStr) return '-';
        return new Date(isoStr).toLocaleString();
    }
    function dateOnly(isoStr) { return isoStr ? isoStr.slice(0, 10) : ''; }
    function formatDateDDMMYY(dateStr) {
        if (!dateStr) return '';
        var parts = String(dateStr).split('-');
        if (parts.length !== 3) return dateStr;
        return parts[2] + '-' + parts[1] + '-' + parts[0].slice(2);
    }

    function getUniqueInspectors() {
        var set = {};
        getRecords().forEach(function (r) { if (r.inspector) set[r.inspector] = true; });
        return Object.keys(set).sort();
    }

    function escapeCSV(val) {
        val = String(val || '');
        if (/[,"\n]/.test(val)) return '"' + val.replace(/"/g, '""') + '"';
        return val;
    }

    function esc(str) {
        var d = document.createElement('div');
        d.textContent = str;
        return d.innerHTML;
    }

    function buildDynamicHeaders() {
        var all = [], seen = {};
        getProducts().forEach(function (p) {
            (p.fields || []).forEach(function (f) {
                if (!seen[f.key]) { seen[f.key] = true; all.push(f); }
            });
        });
        return all;
    }

    // QC Template Data — derived from actual PDF QC sheets
    // Item format: { n:sno, p:'parameter', t:'type', o:[options], ot:hasOtherText }
    // Types: s=select, so=select+other, t=text, y=yes, yn=yes/no, d=dual, ok=okCheck
    var QC_TEMPLATE_DATA = {
        NF: {
            title: 'NeoFly QC Sheet',
            docId: 'NM-SCM-26-01-V01-R01',
            headerFields: [],
            sections: [
                { name: 'NeoFly Selection', items: [
                    { n:1, p:'Frame colour', t:'so', o:['Matt Black'] },
                    { n:2, p:'Rear wheel size', t:'s', o:['24x1','24x1-3/8','24x1.75'] },
                    { n:3, p:'Rear wheel tyre', t:'s', o:['Solid','Pneumatic'] },
                    { n:4, p:'Castor wheel (CW)', t:'s', o:['4"','5.5"','6"'] },
                    { n:5, p:'CW hole position (from top)', t:'t' },
                    { n:6, p:'Pushrim', t:'s', o:['Standard','Ergonomic'] },
                    { n:7, p:'Wheel locks', t:'s', o:['Standard Scissor','Composite Scissor'] },
                    { n:8, p:'Armrest type', t:'s', o:['Flat','Round'] },
                    { n:9, p:'Camber angle (in degree)', t:'s', o:['0','2.5'] },
                    { n:10, p:'Cushion', t:'so', o:['NMC1'] },
                    { n:11, p:'Hub', t:'so', o:['Black Hub-1','Black Hub-2'] },
                    { n:12, p:'Fork', t:'s', o:['Black-2H','Black-5H'] },
                    { n:13, p:'BR angle plate colour', t:'so', o:['Black'] },
                    { n:14, p:'FR plate colour', t:'so', o:['Black'] }
                ]},
                { name: 'Spares / Accessories', items: [
                    { n:15, p:'Spare 1', t:'t', label:'Spare / Accessory' },
                    { n:'15b', p:'Spare 2', t:'t', label:'Spare / Accessory' },
                    { n:'15c', p:'Spare 3', t:'t', label:'Spare / Accessory' },
                    { n:'15d', p:'Spare 4', t:'t', label:'Spare / Accessory' }
                ]},
                { name: 'NeoFly Adjustment', items: [
                    { n:16, p:'Seat width + Increment', t:'d', o:['Width (mm)','Increment (mm)'] },
                    { n:17, p:'Seat depth', t:'t', u:'mm' },
                    { n:18, p:'Backrest position', t:'s', o:['Back','Front'] },
                    { n:19, p:'Backrest height', t:'t', u:'mm' },
                    { n:20, p:'Backrest angle', t:'t', u:'degree' },
                    { n:21, p:'Armrest height', t:'d', o:['UL (mm)','UR (mm)'] },
                    { n:22, p:'Footrest height', t:'t', u:'mm' },
                    { n:23, p:'Wedge', t:'s', o:['Present','Absent'] },
                    { n:24, p:'Wedge position (if wedge present)', t:'t' },
                    { n:25, p:'Spacer', t:'s', o:['Present','Absent'] },
                    { n:26, p:'Footrest fore-aft position', t:'s', o:['1','2','3'] },
                    { n:27, p:'Footrest angle', t:'t', u:'degree' },
                    { n:28, p:'Rear-wheel position', t:'t', u:'mm' },
                    { n:29, p:'Push rim position', t:'s', o:['In','Out'] }
                ]},
                { name: 'NeoFly Regular Checklist', items: [
                    { n:30, p:'Is leg guard assembled?', t:'y' },
                    { n:31, p:'Number of 1.5" straps', t:'t' },
                    { n:32, p:'PR weld plate cover placed?', t:'y' },
                    { n:33, p:'Is strap adjustment completed?', t:'y' },
                    { n:34, p:'AR alignment', t:'s', o:['Both Outward','Straight','Both Inward'] },
                    { n:35, p:'AR removal', t:'ok', v:'Easy to remove' },
                    { n:36, p:'Cushion removal', t:'ok', v:'Sticks to seat sling & able to remove' },
                    { n:37, p:'PH folding', t:'ok', v:'Identical folding' },
                    { n:38, p:'SG removal (SGC2)', t:'s', o:['Easy','Difficult'] },
                    { n:39, p:'SG folding (Play in SGC1)', t:'ok', v:'SG plate holds in air' },
                    { n:40, p:'Backrest unlocking', t:'s', o:['Easy','Difficult'] },
                    { n:41, p:'Backrest locking (folded)', t:'ok', v:'Lift and check in folded position' },
                    { n:42, p:'Is side reflector assembled in rear wheels?', t:'y' },
                    { n:43, p:'Tyre tread direction matching?', t:'y' },
                    { n:44, p:'Wheel axle play (on both side)', t:'s', o:['Minimal','More'] },
                    { n:45, p:'Wheel removal', t:'s', o:['Easy','Difficult'] },
                    { n:46, p:'Wheel interchangeability performed', t:'y' },
                    { n:47, p:'Wheel insertion', t:'s', o:['Easy','Difficult'] },
                    { n:48, p:'Pushrim lateral deviation', t:'ok', v:'Within visual limits on both sides' },
                    { n:49, p:'Pushrim radial deviation', t:'ok', v:'Within visual limits on both sides' },
                    { n:50, p:'Wheel lateral deviation', t:'ok', v:'Nominal' },
                    { n:51, p:'Wheel radial deviation', t:'ok', v:'Nominal' },
                    { n:52, p:'Is tube valve straight on both wheels', t:'y' },
                    { n:53, p:'Is both the tyres inflated to optimal', t:'y' },
                    { n:54, p:'Wheel locking', t:'ok', v:'OK' },
                    { n:55, p:'AT removal', t:'s', o:['Easy','Difficult'] },
                    { n:56, p:'Backrest unlocking & locking', t:'ok', v:'OK' },
                    { n:57, p:'Backrest Play - RubberHeadBolt adjusted', t:'y' },
                    { n:58, p:'Open PH & check locking', t:'ok', v:'OK' },
                    { n:59, p:'AR putting back', t:'s', o:['Easy','Difficult'] },
                    { n:60, p:'AR play measurement', t:'d', o:['UL (mm)','UR (mm)'] },
                    { n:61, p:'AT insertion', t:'s', o:['Easy','Difficult'] },
                    { n:62, p:'AT interchangeability performed', t:'y' },
                    { n:63, p:'SG- tyre clearance', t:'d', o:['UL (mm)','UR (mm)'] },
                    { n:64, p:'Four-point contact achieved?', t:'y' },
                    { n:65, p:'Toe', t:'d', o:['F (mm)','R (mm)'] },
                    { n:66, p:'Both AT touching together?', t:'y' },
                    { n:67, p:'AT hole position (from top)', t:'t' },
                    { n:68, p:'CW Lift when AT touches ground', t:'t', u:'mm' },
                    { n:69, p:'Caster swivel rotation', t:'ok', v:'Nominal' },
                    { n:70, p:'Caster hub play', t:'ok', v:'Nominal' },
                    { n:71, p:'Caster wheel hub play', t:'ok', v:'Nominal' },
                    { n:72, p:'Ground clearance', t:'ok', v:'50mm' },
                    { n:73, p:'ATC bolt tightened?', t:'y' },
                    { n:74, p:'Wheel clamp bolts tightened?', t:'y' },
                    { n:75, p:'Sharp corners deburred from AR bush top', t:'y' },
                    { n:76, p:'Sharp corners deburred from AR bush bottom', t:'y' },
                    { n:77, p:'Sharp corners removed from calf strap', t:'y' },
                    { n:78, p:'Is wheelchair going straight?', t:'y' },
                    { n:79, p:'Wheelie performed and AT tested', t:'y' },
                    { n:80, p:'COC assembled properly?', t:'y' },
                    { n:81, p:'Did WD 40 wide applied and caster caps are closed?', t:'y' }
                ]}
            ]
        },
        NB: {
            title: 'NeoBolt QC Sheet',
            docId: 'NM-SCM-26-02-V01-R01',
            headerFields: [
                { key:'motorNo', label:'Motor S.No' },
                { key:'batteryNo', label:'Battery S.No' },
                { key:'displayNo', label:'Display S.No' },
                { key:'ecuNo', label:'ECU S.No' },
                { key:'chargerNo', label:'Charger S.No' },
                { key:'batteryRange', label:'Battery Range', type:'select', options:['25 Km','50 Km'] },
                { key:'frameSize', label:'Frame Size' }
            ],
            sections: [
                { name: 'NeoBolt QC Checklist', items: [
                    { n:1, p:'Frame colour', t:'ok', v:'Matt Black' },
                    { n:2, p:'AM Adapter Alignment Plate assembled?', t:'yn' },
                    { n:3, p:'AM Adapter Alignment Plate Spacer', t:'d', o:['UL (mm)','UR (mm)'] },
                    { n:4, p:'Fork assembled properly with locknut', t:'yn' },
                    { n:5, p:'Fork bearing check by rotation?', t:'y' },
                    { n:6, p:'Is handle angle limiter hitting on handle angle stopper?', t:'y' },
                    { n:7, p:'Handlebar to wheel assembly straightness', t:'ok', v:'Within visual limits' },
                    { n:8, p:'Is HB-Fork Tube cap assembled?', t:'y' },
                    { n:9, p:'HB Mount Bolt Tightened?', t:'y' },
                    { n:10, p:'HB Top & Bottom Clamp Tightened?', t:'y' },
                    { n:11, p:'HB Top & Bottom Clamp assembled with rubber gasket & metal washer?', t:'y' },
                    { n:12, p:'HB Bottom Clamp Nut Tightened?', t:'y' },
                    { n:13, p:'FP Support tube cap assembled?', t:'y' },
                    { n:14, p:'AT-Straight Tube Cap assembled?', t:'y' },
                    { n:15, p:'AM-Insert fasteners tightened?', t:'y' },
                    { n:16, p:'AM-Stopper E-clips assembled?', t:'y' },
                    { n:17, p:'AM-Spring buckling bolt assembled?', t:'y' },
                    { n:18, p:'AM Spring Engagement/Disengagement by AM Lever?', t:'ok', v:'Stopper rotates equidistant on both sides' },
                    { n:19, p:'AM Bottom Cable tied properly?', t:'y' },
                    { n:20, p:'Is Tire fitted properly - direction of rotation?', t:'y' },
                    { n:21, p:'AM testing Completed?', t:'y' },
                    { n:22, p:'HB Mount Play checked', t:'y' },
                    { n:23, p:'AT-Handle tightened at proper orientation on both sides', t:'y' },
                    { n:24, p:'AT-Spring Engagement/Disengagement checked', t:'y' },
                    { n:25, p:'AT-Pin Locking with AT-MountHub', t:'y' },
                    { n:26, p:'Toggle Clamp assembled?', t:'y' },
                    { n:27, p:'Is toggle bolt hitting the NF frame top tube double bend?', t:'s', o:['No','Yes'] },
                    { n:28, p:'Toggle clamp rubber head bolt tightened?', t:'y' },
                    { n:29, p:'Brake lever left & right tightened at proper orientation?', t:'y' },
                    { n:30, p:'Mirror Mount tightened at proper orientation?', t:'y' },
                    { n:31, p:'Horn & Headlight switch tightened at proper orientation?', t:'y' },
                    { n:32, p:'Display clamp tightened to rotatable condition?', t:'y' },
                    { n:33, p:'Reverse switch tightened at proper orientation?', t:'y' },
                    { n:34, p:'Throttle tightened at proper orientation?', t:'y' },
                    { n:35, p:'AT-Wheel height adjustment completed & tightened?', t:'y' },
                    { n:36, p:'AT-Wheel bolt tightened?', t:'y' },
                    { n:37, p:'AT-Wheel Tube cap assembled?', t:'y' },
                    { n:38, p:'AM-NF Clamp tightened?', t:'y' },
                    { n:39, p:'Brake Cable adjustment done for left & right?', t:'y' },
                    { n:40, p:'Battery Locking on battery slide?', t:'y' },
                    { n:41, p:'Batman colour?', t:'so', o:['Silver'] },
                    { n:42, p:'CH-TT Tube Guard Assembled on NF?', t:'y' },
                    { n:43, p:'NeoMotion Sticker on HB Mount Plate?', t:'ok', v:'Sticked properly' },
                    { n:44, p:'Is Rubber Sleeve assembled on batman to avoid brake cable rubbing?', t:'y' },
                    { n:45, p:'Wireharness clubbed properly?', t:'y' },
                    { n:46, p:'Drum Brake - WheelMountPlate Bolt assembled? M6 x35 - SH', t:'y' },
                    { n:47, p:'Wheel Hub M14 Nut tightened?', t:'y' },
                    { n:48, p:'M14 Nut Rubber cover assembled?', t:'y' },
                    { n:49, p:'Wheel Lock Plate M6x10 bolt tightened?', t:'y' },
                    { n:50, p:'AM Lever Bolt M5x20-SH assembled?', t:'y' },
                    { n:51, p:'AM Yoke aligned straight and tightened?', t:'y' },
                    { n:52, p:'Fork to Suspension bolt tightened?', t:'y' },
                    { n:53, p:'AM Adapter - Stopper Side Touch up done? (If surface is grinded)', t:'y' },
                    { n:54, p:'Battery Mount Perpendicular to Central Hub?', t:'y' },
                    { n:55, p:'Seat Belt assembled on NF?', t:'y' },
                    { n:56, p:'Tube valve having plastic cap?', t:'y' },
                    { n:57, p:'Is the front wheel inflated to optimal pressure?', t:'y' }
                ]},
                { name: 'Electronics', items: [
                    { n:58, p:'Horn switch and sound checked', t:'y' },
                    { n:59, p:'Headlight On/Off switch and functionality checked', t:'y' },
                    { n:60, p:'Reverse switch, reverse direction of rotation of wheel and speed 6 Kmph checked', t:'y' },
                    { n:61, p:'Parking switch is working?', t:'y' },
                    { n:62, p:'Throttle check for all 3 PAS level?', t:'ok', v:'6 / 14 / 24 Kmph' },
                    { n:63, p:'Is text printed on LCD Display screen protector in English?', t:'y' }
                ]}
            ]
        },
        NEOSTAND: {
            title: 'NeoStand QC Sheet',
            docId: 'NM-SCM-26-03-V01-R02',
            headerFields: [
                { key:'linearActuatorNo', label:'Linear Actuator S.No' },
                { key:'joystickNo', label:'Joystick S.No' },
                { key:'batteryNo', label:'Battery S.No' },
                { key:'chargerNo', label:'Charger S.No' },
                { key:'frameSize', label:'Frame Size' }
            ],
            sections: [
                { name: 'NeoStand Selection', items: [
                    { n:1, p:'Frame colour', t:'so', o:['Matt Black'] },
                    { n:2, p:'Rear wheel size', t:'so', o:['24x1-3/8'] },
                    { n:3, p:'Rear wheel tyre', t:'s', o:['Solid','Pneumatic'] },
                    { n:4, p:'Castor wheel (CW)', t:'s', o:['4"','5.5"','6"'] },
                    { n:5, p:'CW hole position (from top)', t:'t' },
                    { n:6, p:'Pushrim', t:'s', o:['Standard','Ergonomic'] },
                    { n:7, p:'Wheel locks', t:'s', o:['Standard Scissor'] },
                    { n:8, p:'Armrest type', t:'s', o:['Flat'] },
                    { n:9, p:'Camber angle (in degree)', t:'s', o:['0','2.5'] },
                    { n:10, p:'Cushion', t:'so', o:['NMC1'] },
                    { n:11, p:'Hub', t:'s', o:['Black Hub-1'] },
                    { n:12, p:'Fork', t:'s', o:['Black-2H'] },
                    { n:13, p:'BR angle plate colour', t:'so', o:['Black'] },
                    { n:14, p:'FR plate colour', t:'so', o:['Black'] },
                    { n:15, p:'Joystick Type', t:'s', o:['Standard'] },
                    { n:16, p:'Joystick Position', t:'s', o:['Left','Right'] }
                ]},
                { name: 'Spares / Accessories', items: [
                    { n:17, p:'Spare 1', t:'t', label:'Spare / Accessory' },
                    { n:18, p:'Spare 2', t:'t', label:'Spare / Accessory' },
                    { n:19, p:'Spare 3', t:'t', label:'Spare / Accessory' },
                    { n:20, p:'Spare 4', t:'t', label:'Spare / Accessory' }
                ]},
                { name: 'NeoStand Adjustment', items: [
                    { n:22, p:'Seat width', t:'t', u:'mm' },
                    { n:23, p:'Seat depth', t:'t', u:'mm' },
                    { n:24, p:'Backrest height', t:'t', u:'mm' },
                    { n:25, p:'Backrest angle', t:'t', u:'degree' },
                    { n:26, p:'Armrest height', t:'d', o:['UL (mm)','UR (mm)'] },
                    { n:27, p:'Footrest height', t:'t', u:'mm' },
                    { n:28, p:'Wedge', t:'s', o:['Present','Absent'] },
                    { n:29, p:'Wedge position (if wedge present)', t:'s', o:['Front','Back'] },
                    { n:30, p:'Spacer', t:'s', o:['Present','Absent'] },
                    { n:31, p:'Rear-wheel position', t:'t', u:'mm' },
                    { n:32, p:'Push rim position', t:'s', o:['In','Out'] }
                ]},
                { name: 'NeoStand Regular Checklist', items: [
                    { n:33, p:'Standing Angle', t:'ok', v:'75 Degree' },
                    { n:34, p:'Seat Tube having End Caps?', t:'y' },
                    { n:35, p:'Footrest Tube having End Caps?', t:'y' },
                    { n:36, p:'Seat sitting on End Cap properly?', t:'yn' },
                    { n:37, p:'Footrest touching ground when at full standing?', t:'y' },
                    { n:38, p:'Wires clubbed properly?', t:'y' },
                    { n:39, p:'Knee Block assembled?', t:'y' },
                    { n:40, p:'Chest Harness assembled?', t:'y' },
                    { n:41, p:'Is leg guard assembled?', t:'y' },
                    { n:42, p:'Number of 1.5" straps', t:'t' },
                    { n:43, p:'PR weld plate cover placed?', t:'y' },
                    { n:44, p:'Is strap adjustment completed?', t:'y' },
                    { n:45, p:'AR alignment', t:'s', o:['Both Outward','Straight','Both Inward'] },
                    { n:46, p:'Cushion removal', t:'ok', v:'Sticks to seat sling & able to remove' },
                    { n:47, p:'Clearance between Top clamp and Actuator Red Knob', t:'t', u:'mm' },
                    { n:48, p:'PH folding', t:'ok', v:'Identical folding' },
                    { n:49, p:'SG Engagement', t:'s', o:['Easy','Difficult'] },
                    { n:50, p:'Backrest unlocking', t:'s', o:['Easy','Difficult'] },
                    { n:51, p:'Backrest locking (folded)', t:'ok', v:'Lift and check in folded position' },
                    { n:52, p:'Tyre tread direction matching?', t:'y' },
                    { n:53, p:'Wheel axle play (on both side)', t:'s', o:['Minimal','More'] },
                    { n:54, p:'Wheel removal', t:'s', o:['Easy','Difficult'] },
                    { n:55, p:'Wheel interchangeability performed', t:'y' },
                    { n:56, p:'Wheel insertion', t:'s', o:['Easy','Difficult'] },
                    { n:57, p:'Pushrim lateral deviation', t:'ok', v:'Within visual limits on both sides' },
                    { n:58, p:'Pushrim radial deviation', t:'ok', v:'Within visual limits on both sides' },
                    { n:59, p:'Wheel lateral deviation', t:'ok', v:'Nominal' },
                    { n:60, p:'Wheel radial deviation', t:'ok', v:'Nominal' },
                    { n:61, p:'Is tube valve straight on both wheels', t:'y' },
                    { n:62, p:'Is both the tyres inflated to optimal pressure', t:'y' },
                    { n:63, p:'Wheel locking', t:'ok', v:'OK' },
                    { n:64, p:'AT removal', t:'s', o:['Easy','Difficult'] },
                    { n:65, p:'Backrest unlocking & locking', t:'ok', v:'OK' },
                    { n:66, p:'Open PH & check locking', t:'ok', v:'OK' },
                    { n:67, p:'AT insertion', t:'s', o:['Easy','Difficult'] },
                    { n:68, p:'AT interchangeability performed', t:'y' },
                    { n:69, p:'SG- tyre clearance', t:'d', o:['UL (mm)','UR (mm)'] },
                    { n:70, p:'Four-point contact achieved?', t:'y' },
                    { n:71, p:'Toe', t:'d', o:['F (mm)','R (mm)'] },
                    { n:72, p:'Both AT touching together?', t:'y' },
                    { n:73, p:'AT hole position (from top)', t:'t' },
                    { n:74, p:'CW Lift when AT touches ground', t:'t', u:'mm' },
                    { n:75, p:'Caster swivel rotation', t:'ok', v:'Nominal' },
                    { n:76, p:'Caster hub play', t:'ok', v:'Nominal' },
                    { n:77, p:'Caster wheel hub play', t:'ok', v:'Nominal' },
                    { n:78, p:'Ground clearance', t:'ok', v:'50mm' },
                    { n:79, p:'QR Adapter bolts tightened?', t:'y' },
                    { n:80, p:'Locktite applied on M22 Nut?', t:'y' },
                    { n:81, p:'Sharp corners removed from calf strap', t:'y' },
                    { n:82, p:'Is wheelchair going straight?', t:'y' },
                    { n:83, p:'Wheelie performed and AT tested', t:'y' },
                    { n:84, p:'COC assembled properly?', t:'y' },
                    { n:85, p:'Did WD 40 wide applied and caster caps are closed?', t:'y' },
                    { n:86, p:'E Clips assembled?', t:'s', o:['Linear Actuator Top','Linear Actuator Bottom','F-Bar Top Tube Left side','F-Bar Top Tube Right side'], multi:true }
                ]},
                { name: 'Electronics', items: [
                    { n:87, p:'Perform Sit to Stand Function', t:'ok', v:'OK' },
                    { n:88, p:'Perform Stand to Sit Function', t:'ok', v:'OK' },
                    { n:89, p:'Horn working checked', t:'y' },
                    { n:90, p:'Joystick Forward/Reverse working checked?', t:'y' },
                    { n:91, p:'Audio sound from joystick in complete sit and stand position', t:'y' },
                    { n:92, p:'Charger Working Verified?', t:'y' }
                ]},
                { name: 'NeoStand Dispatch Checklist', packing: true, items: [
                    { n:1, p:'QR Axle - 2 No', t:'y', box:'Document Box' },
                    { n:2, p:'Toolkit', t:'y', box:'Document Box' },
                    { n:3, p:'Spanner 10-13 - 1 No.', t:'y', box:'Document Box' },
                    { n:4, p:'Box Spanner 18-19 - 1 No.', t:'y', box:'Document Box' },
                    { n:5, p:'Allen Key 3 - 1 No', t:'y', box:'Document Box' },
                    { n:6, p:'Allen Key 4 - 1 No', t:'y', box:'Document Box' },
                    { n:7, p:'Allen Key 5 - 1 No', t:'y', box:'Document Box' },
                    { n:8, p:'Invoice - 1 No', t:'y', box:'Documents' },
                    { n:9, p:'Warranty Card - 1 No', t:'y', box:'Documents' },
                    { n:10, p:'User Manual - 1 No', t:'y', box:'Documents' },
                    { n:11, p:'Cushion - 1 No', t:'y', box:'Utility Box' },
                    { n:12, p:'Knee Block - 1 No', t:'y', box:'Utility Box' },
                    { n:13, p:'Battery in Box - 1 No', t:'y', box:'Utility Box' },
                    { n:14, p:'Charger in Box - 1 No', t:'y', box:'Utility Box' },
                    { n:15, p:'Joystick with Mount in box - 1 No', t:'y', box:'Utility Box' },
                    { n:16, p:'Anti Tipper - 1 Set', t:'y', box:'Utility Box' },
                    { n:17, p:'Document Box - 1 Set', t:'y', box:'Utility Box' },
                    { n:18, p:'Spares, if any', t:'t', box:'Utility Box' },
                    { n:19, p:'Wheels - 1 Set', t:'y', box:'Wheel Box' },
                    { n:20, p:'Wheel Box - 1 No', t:'y', box:'Main Box' },
                    { n:21, p:'Folded NeoStand Frame - 1 No', t:'y', box:'Main Box' },
                    { n:22, p:'Utility Box - 1 No', t:'y', box:'Main Box' }
                ]}
            ]
        }
    };

    // ===========================
    // QR Code
    // ===========================
    function createQRPlaceholder(size, label) {
        var s = size || 80;
        var canvas = document.createElement('canvas');
        canvas.width = s; canvas.height = s;
        var ctx = canvas.getContext('2d');
        ctx.fillStyle = '#eee'; ctx.fillRect(0, 0, s, s);
        ctx.fillStyle = '#999'; ctx.font = '10px sans-serif'; ctx.textAlign = 'center';
        ctx.fillText(label || 'QR N/A', s / 2, s / 2 + 3);
        return canvas;
    }

    function generateQR(text, size, callback) {
        var s = size || 80;
        var value = String(text || '');

        // Preferred path: local QRCode library (from script include)
        if (typeof QRCode !== 'undefined' && QRCode && typeof QRCode.toDataURL === 'function') {
            QRCode.toDataURL(value, { width: s, margin: 1, errorCorrectionLevel: 'M' }, function (err, url) {
                if (!err && url) {
                    var img = document.createElement('img');
                    img.width = s; img.height = s;
                    img.src = url;
                    if (callback) callback(img);
                    return;
                }

                // Fallback: use hosted QR image so reports still show scannable QR
                var fallbackImg = document.createElement('img');
                fallbackImg.width = s; fallbackImg.height = s;
                fallbackImg.alt = 'QR Code';
                fallbackImg.onerror = function () {
                    if (callback) callback(createQRPlaceholder(s, 'QR Error'));
                };
                fallbackImg.onload = function () {
                    if (callback) callback(fallbackImg);
                };
                fallbackImg.src = 'https://quickchart.io/qr?size=' + s + '&text=' + encodeURIComponent(value);
            });
            return;
        }

        // Fallback when QRCode library is unavailable
        var externalImg = document.createElement('img');
        externalImg.width = s; externalImg.height = s;
        externalImg.alt = 'QR Code';
        externalImg.onerror = function () {
            if (callback) callback(createQRPlaceholder(s, 'QR N/A'));
        };
        externalImg.onload = function () {
            if (callback) callback(externalImg);
        };
        externalImg.src = 'https://quickchart.io/qr?size=' + s + '&text=' + encodeURIComponent(value);
    }

    // QR Modal
    var qrModal = document.getElementById('qrModal');
    var qrModalBody = document.getElementById('qrModalBody');
    document.getElementById('qrModalClose').addEventListener('click', function () { qrModal.classList.add('hidden'); });
    qrModal.addEventListener('click', function (e) { if (e.target === qrModal) qrModal.classList.add('hidden'); });

    function showQRModal(record) {
        var product = getProductById(record.type);
        var typeName = product ? product.name : record.type;
        var html = '<div id="qrModalPrintArea">';
        html += '<div class="qr-modal-qr" id="qrModalQRImage"></div>';
        html += '<div class="qr-detail-row"><span class="qr-detail-label">Product</span><span class="qr-detail-value">' + esc(typeName) + '</span></div>';
        html += '<div class="qr-detail-row"><span class="qr-detail-label">Order No</span><span class="qr-detail-value">' + esc(record.orderNo) + '</span></div>';
        html += '<div class="qr-detail-row"><span class="qr-detail-label">Frame/Chassis No</span><span class="qr-detail-value">' + esc(record.frameNo) + '</span></div>';
        if (product && product.fields) {
            product.fields.forEach(function (f) {
                html += '<div class="qr-detail-row"><span class="qr-detail-label">' + esc(f.label) + '</span><span class="qr-detail-value">' + esc(record[f.key] || '-') + '</span></div>';
            });
        }
        html += '<div class="qr-detail-row"><span class="qr-detail-label">Produced by</span><span class="qr-detail-value">' + esc(record.inspector) + '</span></div>';
        html += '<div class="qr-detail-row"><span class="qr-detail-label">Timestamp</span><span class="qr-detail-value">' + formatDate(record.timestamp) + '</span></div>';
        html += '</div>';
        html += '<button class="btn btn-primary qr-print-btn" id="qrPrintBtn">Print Order Summary</button>';
        qrModalBody.innerHTML = html;

        // Generate larger QR for the modal
        var qrContainer = document.getElementById('qrModalQRImage');
        generateQR(record.orderNo, 160, function (el) {
            qrContainer.appendChild(el);
        });

        // Print handler
        document.getElementById('qrPrintBtn').addEventListener('click', function () {
            var printArea = document.getElementById('qrModalPrintArea');
            var printWin = window.open('', '_blank', 'width=400,height=600');
            printWin.document.write('<html><head><title>QR Code - ' + esc(record.orderNo) + '</title>');
            printWin.document.write('<style>body{font-family:Arial,sans-serif;padding:20px;text-align:center;}');
            printWin.document.write('.qr-detail-row{display:flex;padding:6px 0;border-bottom:1px solid #ddd;font-size:14px;text-align:left;}');
            printWin.document.write('.qr-detail-label{font-weight:600;min-width:140px;color:#555;}');
            printWin.document.write('.qr-detail-value{flex:1;}.qr-modal-qr{margin:0 auto 16px;text-align:center;}');
            printWin.document.write('.qr-modal-qr img,.qr-modal-qr canvas{width:200px;height:200px;}</style></head><body>');
            printWin.document.write('<h2>Neotrace - Order QR</h2>');
            printWin.document.write(printArea.innerHTML);
            printWin.document.write('</body></html>');
            printWin.document.close();
            printWin.focus();
            printWin.print();
        });

        qrModal.classList.remove('hidden');
    }

    function printLabelOnlyQR(record) {
        var printWin = window.open('', '_blank', 'width=320,height=420');
        if (!printWin) return;

        printWin.document.write('<html><head><title>QR Label - ' + esc(record.orderNo) + '</title>');
        printWin.document.write('<style>body{font-family:Arial,sans-serif;padding:12px;text-align:center;}');
        printWin.document.write('.qr-wrap img,.qr-wrap canvas{width:220px;height:220px;}');
        printWin.document.write('.order-text{margin-top:8px;font-size:14px;font-weight:600;}</style></head><body>');
        printWin.document.write('<div class=\"qr-wrap\" id=\"labelQrWrap\"></div>');
        printWin.document.write('<div class=\"order-text\">Order: ' + esc(record.orderNo) + '</div>');
        printWin.document.write('</body></html>');
        printWin.document.close();

        generateQR(record.orderNo, 220, function (el) {
            var wrap = printWin.document.getElementById('labelQrWrap');
            if (wrap) {
                wrap.appendChild(el);
                setTimeout(function () {
                    printWin.focus();
                    printWin.print();
                }, 200);
            }
        });
    }

    // ===========================
    // Navigation
    // ===========================
    var navLinks = document.querySelectorAll('.nav-link');
    var pages = document.querySelectorAll('.page');
    var menuToggle = document.getElementById('menuToggle');
    var mainNav = document.getElementById('mainNav');

    function showPage(pageId) {
        pages.forEach(function (p) { p.classList.remove('active'); });
        navLinks.forEach(function (l) { l.classList.remove('active'); });
        var target = document.getElementById('page-' + pageId);
        if (target) target.classList.add('active');
        navLinks.forEach(function (l) { if (l.getAttribute('data-page') === pageId) l.classList.add('active'); });
        if (mainNav) mainNav.classList.remove('open');

        if (pageId === 'entry') { populateEntryProductDropdown(); startLiveClock('entry-timestamp'); } else { stopLiveClock(); }
        if (pageId === 'products') renderProductsList();
        if (pageId === 'records') renderRecords();
        if (pageId === 'reports') renderReport(getRecords());
        if (pageId === 'qc') initQCForm();
        if (pageId === 'dashboard') { populateInspectorDropdowns(); clearDashboardView(); }
    }

    navLinks.forEach(function (link) {
        link.addEventListener('click', function (e) { e.preventDefault(); showPage(this.getAttribute('data-page')); });
    });
    if (menuToggle) { menuToggle.addEventListener('click', function () { mainNav.classList.toggle('open'); }); }

    var _liveClockInterval = null;
    var _liveClockInputId = null;
    var _liveClockUserEdited = false;

    function toLocalDatetimeValue(date) {
        var d = new Date(date);
        d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
        return d.toISOString().slice(0, 16);
    }

    function startLiveClock(inputId) {
        stopLiveClock();
        _liveClockInputId = inputId;
        _liveClockUserEdited = false;
        var input = document.getElementById(inputId);
        if (!input) return;

        // Set immediately
        input.value = toLocalDatetimeValue(new Date());

        // Mark as user-edited if they manually change the field
        input.addEventListener('change', function () {
            _liveClockUserEdited = true;
            stopLiveClock();
        });

        _liveClockInterval = setInterval(function () {
            if (_liveClockUserEdited) { stopLiveClock(); return; }
            var el = document.getElementById(_liveClockInputId);
            if (el) el.value = toLocalDatetimeValue(new Date());
        }, 1000);
    }

    function stopLiveClock() {
        if (_liveClockInterval) {
            clearInterval(_liveClockInterval);
            _liveClockInterval = null;
        }
    }

    function setDefaultTimestamp(inputId) {
        startLiveClock(inputId);
    }

    // ===========================
    // Products Page
    // ===========================
    var pendingFields = [];

    function fieldKeyFromLabel(label) {
        return label.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
    }

    function getAllKnownFields() {
        var all = [], seen = {};
        getProducts().forEach(function (p) {
            (p.fields || []).forEach(function (f) {
                if (!seen[f.key]) { seen[f.key] = true; all.push({ key: f.key, label: f.label }); }
            });
        });
        return all;
    }

    function renderPendingFields() {
        var container = document.getElementById('customFieldsList');
        container.innerHTML = '';

        // Always show checkboxes for all known fields (in both add and edit mode)
        var editId = document.getElementById('edit-product-id').value;
        var knownFields = getAllKnownFields();
        if (knownFields.length > 0) {
            var cbGroup = document.createElement('div');
            cbGroup.className = 'checkbox-group existing-fields-checkboxes';
            var heading = document.createElement('div');
            heading.style.fontSize = '0.8125rem';
            heading.style.color = '#444';
            heading.style.marginBottom = '0.25rem';
            heading.textContent = 'Select from existing fields:';
            cbGroup.appendChild(heading);

            knownFields.forEach(function (f) {
                var isChecked = pendingFields.some(function (pf) { return pf.key === f.key; });
                var lbl = document.createElement('label');
                lbl.className = 'checkbox-label';
                var cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = isChecked;
                cb.setAttribute('data-key', f.key);
                cb.setAttribute('data-label', f.label);
                cb.addEventListener('change', function () {
                    var key = this.getAttribute('data-key');
                    var label = this.getAttribute('data-label');
                    if (this.checked) {
                        if (!pendingFields.some(function (pf) { return pf.key === key; })) {
                            pendingFields.push({ key: key, label: label });
                        }
                    } else {
                        pendingFields = pendingFields.filter(function (pf) { return pf.key !== key; });
                    }
                    renderCustomFieldItems();
                });
                lbl.appendChild(cb);
                lbl.appendChild(document.createTextNode(f.label));
                cbGroup.appendChild(lbl);
            });
            container.appendChild(cbGroup);
        }

        var itemsContainer = document.createElement('div');
        itemsContainer.id = 'customFieldItems';
        container.appendChild(itemsContainer);
        renderCustomFieldItems();
    }

    function renderCustomFieldItems() {
        var container = document.getElementById('customFieldItems');
        if (!container) return;
        container.innerHTML = '';
        pendingFields.forEach(function (f, idx) {
            var item = document.createElement('div');
            item.className = 'custom-field-item';
            item.innerHTML = '<span>' + esc(f.label) + '</span><button type="button" class="remove-field-btn" data-idx="' + idx + '">&times;</button>';
            container.appendChild(item);
        });
        container.querySelectorAll('.remove-field-btn').forEach(function (btn) {
            btn.addEventListener('click', function () {
                var idx = parseInt(this.getAttribute('data-idx'));
                var removed = pendingFields.splice(idx, 1)[0];
                var checkboxes = document.querySelectorAll('.existing-fields-checkboxes input[type="checkbox"]');
                checkboxes.forEach(function (cb) {
                    if (cb.getAttribute('data-key') === removed.key) cb.checked = false;
                });
                renderCustomFieldItems();
            });
        });
    }

    document.getElementById('addFieldBtn').addEventListener('click', function () {
        var input = document.getElementById('new-field-name');
        var label = input.value.trim();
        if (!label) return;
        var key = fieldKeyFromLabel(label);
        if (!key) return;
        // Block if key or label already exists in pendingFields
        if (pendingFields.some(function (f) { return f.key === key || f.label.toLowerCase() === label.toLowerCase(); })) {
            alert('A field with this name already exists.');
            return;
        }
        // Block if label matches any known global field (case-insensitive)
        var allKnown = getAllKnownFields();
        var existingMatch = allKnown.find(function (f) { return f.label.toLowerCase() === label.toLowerCase() || f.key === key; });
        if (existingMatch) {
            // Auto-add the existing field instead of creating a duplicate
            if (!pendingFields.some(function (f) { return f.key === existingMatch.key; })) {
                pendingFields.push({ key: existingMatch.key, label: existingMatch.label });
                input.value = '';
                renderPendingFields();
            } else {
                alert('Field "' + existingMatch.label + '" is already added.');
            }
            return;
        }
        pendingFields.push({ key: key, label: label });
        input.value = '';
        renderPendingFields();
    });

    // Live validation: warn if typed name matches existing field
    document.getElementById('new-field-name').addEventListener('input', function () {
        var label = this.value.trim();
        var key = fieldKeyFromLabel(label);
        var allKnown = getAllKnownFields();
        var existingEl = document.getElementById('new-field-hint');
        if (!existingEl) {
            existingEl = document.createElement('small');
            existingEl.id = 'new-field-hint';
            existingEl.style.color = '#cc7700';
            existingEl.style.fontSize = '0.8rem';
            this.parentNode.appendChild(existingEl);
        }
        var match = label && allKnown.find(function (f) { return f.label.toLowerCase() === label.toLowerCase() || f.key === key; });
        var inPending = label && pendingFields.some(function (f) { return f.label.toLowerCase() === label.toLowerCase() || f.key === key; });
        if (inPending) {
            existingEl.textContent = '⚠ This field is already added.';
        } else if (match) {
            existingEl.textContent = '↩ Matches existing field "' + match.label + '" — clicking Add will reuse it.';
        } else {
            existingEl.textContent = '';
        }
    });

    function resetProductForm() {
        document.getElementById('edit-product-id').value = '';
        document.getElementById('new-product-name').value = '';
        document.getElementById('new-field-name').value = '';
        document.getElementById('productFormTitle').textContent = 'Add New Product';
        document.getElementById('productSubmitBtn').textContent = 'Add Product';
        document.getElementById('cancelProductEdit').classList.add('hidden');
        pendingFields = [];
        renderPendingFields();
    }

    function renderProductsList() {
        var products = getProducts();
        var container = document.getElementById('productsList');
        container.innerHTML = '';
        if (products.length === 0) { container.innerHTML = '<p>No products configured.</p>'; return; }

        products.forEach(function (p) {
            var card = document.createElement('div');
            card.className = 'product-card';
            var fieldNames = p.fields.length > 0 ? p.fields.map(function (f) { return f.label; }).join(', ') : 'No extra fields';
            card.innerHTML =
                '<div class="product-info"><strong>' + esc(p.name) + '</strong>' +
                '<span class="product-fields">Fields: Frame/Chassis No' + (p.fields.length > 0 ? ', ' + esc(fieldNames) : '') + ', Produced by</span></div>' +
                '<div class="admin-record-actions">' +
                '<button class="btn btn-secondary btn-sm product-modify-btn" data-pid="' + esc(p.id) + '">Modify</button>' +
                '<button class="btn btn-danger btn-sm product-delete-btn" data-pid="' + esc(p.id) + '">Delete</button></div>';
            container.appendChild(card);
        });

        container.querySelectorAll('.product-delete-btn').forEach(function (btn) {
            btn.addEventListener('click', function () {
                var p = getProductById(this.getAttribute('data-pid'));
                if (!p || !confirm('Delete product "' + p.name + '"?')) return;
                deleteProduct(p.id);
                renderProductsList();
                populateEntryProductDropdown();
            });
        });

        container.querySelectorAll('.product-modify-btn').forEach(function (btn) {
            btn.addEventListener('click', function () {
                var p = getProductById(this.getAttribute('data-pid'));
                if (!p) return;
                document.getElementById('edit-product-id').value = p.id;
                document.getElementById('new-product-name').value = p.name;
                document.getElementById('productFormTitle').textContent = 'Modify Product';
                document.getElementById('productSubmitBtn').textContent = 'Save Changes';
                document.getElementById('cancelProductEdit').classList.remove('hidden');
                pendingFields = p.fields.map(function (f) { return { key: f.key, label: f.label }; });
                renderPendingFields();
                document.getElementById('addProductForm').scrollIntoView({ behavior: 'smooth' });
            });
        });
    }

    document.getElementById('cancelProductEdit').addEventListener('click', resetProductForm);

    document.getElementById('addProductForm').addEventListener('submit', function (e) {
        e.preventDefault();
        var name = document.getElementById('new-product-name').value.trim();
        if (!name) return;
        var editId = document.getElementById('edit-product-id').value;
        var conf = document.getElementById('product-confirmation');

        if (editId) {
            updateProduct(editId, { name: name, fields: pendingFields.slice() });
            conf.textContent = 'Product "' + name + '" updated successfully.';
        } else {
            addProductToStore({ name: name, fields: pendingFields.slice() });
            conf.textContent = 'Product "' + name + '" added successfully.';
        }
        conf.className = 'confirmation confirmation-success';
        resetProductForm();
        renderProductsList();
        populateEntryProductDropdown();
        setTimeout(function () { conf.className = 'confirmation hidden'; }, 4000);
    });

    // ===========================
    // Entry Form
    // ===========================
    function populateEntryProductDropdown() {
        var products = getProducts();
        var select = document.getElementById('entry-product');
        var cur = select.value;
        select.innerHTML = '<option value="">Select product</option>';
        products.forEach(function (p) {
            var opt = document.createElement('option');
            opt.value = p.id; opt.textContent = p.name;
            select.appendChild(opt);
        });
        select.value = cur;
        updateEntryFields();
    }

    function updateEntryFields() {
        var pid = document.getElementById('entry-product').value;
        var product = pid ? getProductById(pid) : null;
        var container = document.getElementById('entryDynamicFields');
        container.innerHTML = '';
        if (!product || !product.fields) return;
        product.fields.forEach(function (f) {
            var div = document.createElement('div');
            div.className = 'form-group';
            div.innerHTML = '<label for="entry-' + esc(f.key) + '">' + esc(f.label) + '</label>' +
                '<input type="text" id="entry-' + esc(f.key) + '" required placeholder="Scan or enter ' + esc(f.label.toLowerCase()) + '">';
            container.appendChild(div);
        });
    }

    document.getElementById('entry-product').addEventListener('change', updateEntryFields);

    function showEntryMessage(text, type) {
        var conf = document.getElementById('entry-confirmation');
        conf.textContent = text;
        conf.className = 'confirmation ' + (type === 'success' ? 'confirmation-success' : 'confirmation-error');
        setTimeout(function () { conf.className = 'confirmation hidden'; }, 5000);
    }

    function checkDuplicate(record, product) {
        var records = getRecords();
        for (var i = 0; i < records.length; i++) {
            var r = records[i];
            if (r.orderNo === record.orderNo && r.frameNo === record.frameNo && r.type === record.type)
                return 'Duplicate: Order ' + record.orderNo + ' with Frame ' + record.frameNo + ' already exists.';
            if (record.frameNo && r.frameNo === record.frameNo && r.type === record.type)
                return 'Duplicate Frame/Chassis No: ' + record.frameNo + ' already exists in Order ' + r.orderNo + '.';
            if (product && product.fields) {
                for (var j = 0; j < product.fields.length; j++) {
                    var fk = product.fields[j].key;
                    if (record[fk] && r[fk] && record[fk] === r[fk])
                        return 'Duplicate ' + product.fields[j].label + ': ' + record[fk] + ' already exists in Order ' + r.orderNo + '.';
                }
            }
        }
        return null;
    }

    document.getElementById('entryForm').addEventListener('submit', function (e) {
        e.preventDefault();
        var pid = document.getElementById('entry-product').value;
        var product = getProductById(pid);
        if (!product) { alert('Please select a product.'); return; }

        var record = {
            type: product.id,
            orderNo: document.getElementById('entry-orderNo').value.trim(),
            frameNo: document.getElementById('entry-frameNo').value.trim(),
            inspector: document.getElementById('entry-inspector').value.trim(),
            timestamp: document.getElementById('entry-timestamp').value
        };
        if (product.fields) {
            product.fields.forEach(function (f) {
                var inp = document.getElementById('entry-' + f.key);
                record[f.key] = inp ? inp.value.trim() : '';
            });
        }

        var dup = checkDuplicate(record, product);
        if (dup) { showEntryMessage(dup, 'error'); return; }

        addRecord(record);
        showEntryMessage(product.name + ' entry saved successfully for Order: ' + record.orderNo, 'success');
        this.reset();
        setDefaultTimestamp('entry-timestamp');
        populateEntryProductDropdown();
    });

    // ===========================
    // Records Page — Master Records
    // ===========================
    var mrSort = { col: null, dir: 'asc' };
    var mrActiveFilters = { orderNo: [], status: [] };

    // Group records by orderNo, collecting all product type names per order
    function buildMasterGroups() {
        var records = getRecords();
        var map = {};
        var order = [];
        records.forEach(function (r) {
            var product = getProductById(r.type);
            var typeName = product ? product.name : (r.type || 'Unknown');
            if (!map[r.orderNo]) {
                map[r.orderNo] = { orderNo: r.orderNo, products: [] };
                order.push(r.orderNo);
            }
            if (map[r.orderNo].products.indexOf(typeName) === -1) {
                map[r.orderNo].products.push(typeName);
            }
        });
        return order.map(function (k) { return map[k]; });
    }

    function renderRecords() {
        var tbody = document.getElementById('recordsBody');
        var noMsg = document.getElementById('noRecordsMsg');
        if (!tbody) return;
        tbody.innerHTML = '';

        var groups = buildMasterGroups();

        // Apply filters
        if (mrActiveFilters.orderNo.length > 0) {
            groups = groups.filter(function (g) {
                return mrActiveFilters.orderNo.indexOf(g.orderNo) !== -1;
            });
        }
        if (mrActiveFilters.status.length > 0) {
            groups = groups.filter(function (g) {
                return g.products.some(function (p) {
                    return mrActiveFilters.status.indexOf(p) !== -1;
                });
            });
        }

        // Apply sort
        if (mrSort.col) {
            groups.sort(function (a, b) {
                var va = mrSort.col === 'orderNo' ? a.orderNo : a.products.join(', ');
                var vb = mrSort.col === 'orderNo' ? b.orderNo : b.products.join(', ');
                return mrSort.dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
            });
        }

        if (groups.length === 0) {
            noMsg.textContent = 'No records found.';
            noMsg.classList.remove('hidden');
            return;
        }
        noMsg.classList.add('hidden');

        groups.forEach(function (g) {
            var tr = document.createElement('tr');
            var badgesHtml = g.products.map(function (p) {
                return '<span class="status-badge status-produced">' + esc(p) + '</span>';
            }).join(' ');
            tr.innerHTML =
                '<td>' + esc(g.orderNo) + '</td>' +
                '<td class="status-cell">' + badgesHtml + '</td>';
            tbody.appendChild(tr);
        });

        updateMasterSortIcons();
    }

    function updateMasterSortIcons() {
        document.querySelectorAll('.mr-sort-btn').forEach(function (btn) {
            var col = btn.getAttribute('data-col');
            btn.classList.remove('sort-asc', 'sort-desc');
            if (mrSort.col === col) btn.classList.add('sort-' + mrSort.dir);
        });
        // Highlight active filter buttons
        document.querySelectorAll('.mr-filter-btn').forEach(function (btn) {
            var col = btn.getAttribute('data-col');
            btn.classList.toggle('filter-active', mrActiveFilters[col] && mrActiveFilters[col].length > 0);
        });
    }

    function openMasterFilterDropdown(col) {
        var dropdown = document.getElementById('mr-filter-' + col);
        if (!dropdown) return;
        // Close other dropdowns first
        document.querySelectorAll('.mr-filter-dropdown').forEach(function (d) {
            if (d !== dropdown) d.classList.add('hidden');
        });
        if (!dropdown.classList.contains('hidden')) {
            dropdown.classList.add('hidden');
            return;
        }
        // Build unique values
        var groups = buildMasterGroups();
        var values = [];
        groups.forEach(function (g) {
            var candidates = col === 'orderNo' ? [g.orderNo] : g.products;
            candidates.forEach(function (v) {
                if (values.indexOf(v) === -1) values.push(v);
            });
        });
        values.sort();
        dropdown.innerHTML = '';
        var activeList = mrActiveFilters[col];
        // Select-all checkbox
        var allChecked = activeList.length === 0;
        var allLabel = document.createElement('label');
        allLabel.className = 'mr-filter-item';
        allLabel.innerHTML = '<input type="checkbox" class="mr-filter-all" data-col="' + col + '"' + (allChecked ? ' checked' : '') + '> <em>All</em>';
        dropdown.appendChild(allLabel);
        // Value checkboxes
        values.forEach(function (v) {
            var checked = activeList.length === 0 || activeList.indexOf(v) !== -1;
            var lbl = document.createElement('label');
            lbl.className = 'mr-filter-item';
            lbl.innerHTML = '<input type="checkbox" class="mr-filter-val" data-col="' + col + '" value="' + esc(v) + '"' + (checked ? ' checked' : '') + '> ' + esc(v);
            dropdown.appendChild(lbl);
        });
        // Apply button
        var applyBtn = document.createElement('button');
        applyBtn.className = 'btn btn-primary mr-filter-apply';
        applyBtn.textContent = 'Apply';
        applyBtn.addEventListener('click', function () {
            var allBox = dropdown.querySelector('.mr-filter-all');
            if (allBox && allBox.checked) {
                mrActiveFilters[col] = [];
            } else {
                mrActiveFilters[col] = [];
                dropdown.querySelectorAll('.mr-filter-val:checked').forEach(function (cb) {
                    mrActiveFilters[col].push(cb.value);
                });
            }
            dropdown.classList.add('hidden');
            renderRecords();
        });
        dropdown.appendChild(applyBtn);
        // Prevent clicks inside the dropdown from bubbling to document (which closes it)
        dropdown.addEventListener('click', function (e) { e.stopPropagation(); });
        dropdown.classList.remove('hidden');
    }

    // Sort button clicks
    document.querySelectorAll('.mr-sort-btn').forEach(function (btn) {
        btn.addEventListener('click', function () {
            var col = btn.getAttribute('data-col');
            if (mrSort.col === col) {
                mrSort.dir = mrSort.dir === 'asc' ? 'desc' : 'asc';
            } else {
                mrSort.col = col;
                mrSort.dir = 'asc';
            }
            renderRecords();
        });
    });

    // Filter button clicks
    document.querySelectorAll('.mr-filter-btn').forEach(function (btn) {
        btn.addEventListener('click', function (e) {
            e.stopPropagation();
            openMasterFilterDropdown(btn.getAttribute('data-col'));
        });
    });

    // "All" checkbox toggles individual boxes
    document.getElementById('recordsTable').addEventListener('change', function (e) {
        if (e.target.classList.contains('mr-filter-all')) {
            var col = e.target.getAttribute('data-col');
            var dropdown = document.getElementById('mr-filter-' + col);
            dropdown.querySelectorAll('.mr-filter-val').forEach(function (cb) {
                cb.checked = e.target.checked;
            });
        }
        if (e.target.classList.contains('mr-filter-val')) {
            var col2 = e.target.getAttribute('data-col');
            var dropdown2 = document.getElementById('mr-filter-' + col2);
            var allBox = dropdown2.querySelector('.mr-filter-all');
            var allChecked2 = Array.from(dropdown2.querySelectorAll('.mr-filter-val')).every(function (cb) { return cb.checked; });
            if (allBox) allBox.checked = allChecked2;
        }
    });

    // Close dropdowns when clicking outside
    document.addEventListener('click', function () {
        document.querySelectorAll('.mr-filter-dropdown').forEach(function (d) {
            d.classList.add('hidden');
        });
    });

    // ===========================
    // Chatbot
    // ===========================
    (function () {
        var KB = [
            { keys: ['hello','hi','hey','greet','good morning','good afternoon'], answer: "Hello! I'm the Neotrace Assistant. I can answer questions about your data or help you navigate the app. What would you like to know?" },
            { keys: ['what is neotrace','about neotrace','neotrace app','what does neotrace do','purpose'], answer: "Neotrace is a production traceability web app for tracking wheelchair and mobility products (NeoFly, NeoBolt, NeoStand). It logs production entries, QC inspections, reports, and syncs data to Google Sheets." },
            { keys: ['add entry','new entry','create entry','how to add','log production','production entry','entry tab'], answer: "To add a production entry:\n1. Go to the <b>Entry</b> tab\n2. Select the product type\n3. Fill in Order No, Frame/Chassis No, and required fields\n4. Click <b>Save Entry</b>" },
            { keys: ['master records','what is master'], answer: "<b>Master Records</b> shows all production entries grouped by Order No with sortable and filterable columns." },
            { keys: ['detailed records','what is detailed'], answer: "<b>Detailed Records</b> shows the full table with all fields, type filter, keyword search, date-range filter, and CSV export." },
            { keys: ['qc report','quality check','quality control','submit qc','qc form'], answer: "To file a QC Report:\n1. Go to <b>QC Reports</b> tab\n2. Select the template (NeoFly, NeoBolt, NeoStand)\n3. Complete the checklist\n4. Click <b>Download PDF</b> for a local copy or <b>Submit QC Report</b> to send to Google Drive." },
            { keys: ['download pdf','pdf download','save pdf'], answer: "On the <b>QC Reports</b> tab, click <b>Download PDF</b> (above the Submit button) to generate and save a formatted PDF to your device." },
            { keys: ['download csv','export csv','csv'], answer: "In <b>Detailed Records</b>, apply filters then click <b>Download CSV</b> to export." },
            { keys: ['sort','sorting'], answer: "In <b>Master Records</b>, click the ⇅ icon next to any column header to sort ascending/descending." },
            { keys: ['filter','filtering'], answer: "In <b>Master Records</b>, click the ▽ icon on any column header, select the values you want, and press <b>Apply</b>." },
            { keys: ['dashboard','charts','performance'], answer: "The <b>Dashboard</b> tab shows production and QC charts. Use the date/inspector filters and click <b>Load</b>. Click <b>Go to Dashboard</b> below." },
            { keys: ['reports tab','report page'], answer: "The <b>Reports</b> tab lets you search all records with keyword and date filters, and download CSV." },
            { keys: ['sync','google sheet','google sheets','webhook'], answer: "To sync to Google Sheets:\n1. Deploy the Google Apps Script as a Web App\n2. In <b>Admin</b> tab, paste the Web App URL\n3. Click <b>Sync All Records Now</b>" },
            { keys: ['admin','admin panel','password'], answer: "The <b>Admin</b> tab (default password: admin123) lets you edit/delete records and configure Google Sheets sync." },
            { keys: ['product','products tab','neofly','neobolt','neostand'], answer: "The <b>Products</b> tab lists all product types and their tracked fields (Battery No, Charger No, Motor No, etc.)." },
            { keys: ['offline','pwa','install app','mobile'], answer: "Neotrace is a PWA. Tap <b>Add to Home Screen</b> in your browser to install it. It works offline and syncs when back online." },
            { keys: ['delete','remove record','edit record'], answer: "To edit or delete a record, go to the <b>Admin</b> tab." },
            { keys: ['duplicate','duplicate order','duplicate frame'], answer: "Neotrace blocks duplicate Order Nos and Frame/Chassis Nos when saving an entry." },
            { keys: ['help','what can you do','topics'], answer: "I can help with:\n• Live data — orders, records, QC counts\n• Navigating to any page\n• How-to guidance on any feature\n\nTry: <i>\"How many orders?\"</i>, <i>\"Show orders\"</i>, <i>\"Find order ORD-001\"</i>, or <i>\"Go to QC Reports\"</i>" }
        ];

        var PAGE_NAV = [
            { keys: ['go to entry','open entry','entry tab','navigate entry'], page: 'entry', label: 'Entry' },
            { keys: ['go to records','open records','records tab','navigate records'], page: 'records', label: 'Records' },
            { keys: ['go to reports','open reports','reports tab','navigate reports'], page: 'reports', label: 'Reports' },
            { keys: ['go to qc','open qc','qc tab','navigate qc','go to qc reports','open qc reports'], page: 'qc', label: 'QC Reports' },
            { keys: ['go to dashboard','open dashboard','dashboard tab','navigate dashboard'], page: 'dashboard', label: 'Dashboard' },
            { keys: ['go to products','open products','products tab','navigate products'], page: 'products', label: 'Products' },
            { keys: ['go to admin','open admin','admin tab','navigate admin'], page: 'admin', label: 'Admin' }
        ];

        var QUICK = [
            'How many orders do we have?',
            'Show all orders',
            'How many QC reports?',
            'Go to QC Reports',
            'Go to Dashboard'
        ];

        var fab      = document.getElementById('chatbot-fab');
        var win      = document.getElementById('chatbot-window');
        var closeBtn = document.getElementById('chatbot-close');
        var msgBox   = document.getElementById('chatbot-messages');
        var qrBox    = document.getElementById('chatbot-quick-replies');
        var input    = document.getElementById('chatbot-input');
        var sendBtn  = document.getElementById('chatbot-send');
        var opened   = false;

        function escHtml(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

        function appendMsg(text, role) {
            var div = document.createElement('div');
            div.className = 'cb-msg cb-' + role;
            div.innerHTML = role === 'bot' ? text.replace(/\n/g, '<br>') : escHtml(text);
            msgBox.appendChild(div);
            msgBox.scrollTop = msgBox.scrollHeight;
        }

        function showTyping() {
            var div = document.createElement('div');
            div.className = 'cb-msg cb-bot cb-typing';
            div.id = 'cb-typing-indicator';
            div.innerHTML = '<span></span><span></span><span></span>';
            msgBox.appendChild(div);
            msgBox.scrollTop = msgBox.scrollHeight;
        }

        function removeTyping() {
            var el = document.getElementById('cb-typing-indicator');
            if (el) el.remove();
        }

        // --- Data-aware query handlers ---
        function handleDataQuery(lq) {
            var records = getRecords();
            var qcRecs = getQCRecords();

            // Total order/record count
            if (/how many (orders|records|entries|total)/.test(lq) || lq === 'total orders' || lq === 'count orders') {
                var groups = {};
                records.forEach(function(r){ groups[r.orderNo] = true; });
                var orderCount = Object.keys(groups).length;
                return 'There are <b>' + records.length + ' production entries</b> across <b>' + orderCount + ' unique orders</b> in the system.';
            }

            // Show/list all orders
            if (/^(show|list|all) ?(all ?)?(orders?|records?)/.test(lq) || lq === 'show orders' || lq === 'list orders') {
                if (records.length === 0) return 'No production records found yet.';
                var map = {};
                records.forEach(function(r){
                    if (!map[r.orderNo]) map[r.orderNo] = [];
                    var p = getProductById(r.type);
                    var name = p ? p.name : r.type;
                    if (map[r.orderNo].indexOf(name) === -1) map[r.orderNo].push(name);
                });
                var keys = Object.keys(map).sort();
                var lines = keys.map(function(k){ return '• <b>' + escHtml(k) + '</b> — ' + escHtml(map[k].join(', ')); });
                return 'Found <b>' + keys.length + ' orders</b>:\n' + lines.join('\n');
            }

            // Find specific order
            var orderMatch = lq.match(/(?:order|find|show|lookup|look up)\s+([\w\-\/]+)/);
            if (orderMatch) {
                var target = orderMatch[1].toUpperCase();
                var hits = records.filter(function(r){ return r.orderNo.toUpperCase() === target; });
                if (hits.length === 0) {
                    // fuzzy — partial match
                    hits = records.filter(function(r){ return r.orderNo.toUpperCase().indexOf(target) !== -1; });
                }
                if (hits.length === 0) return 'No records found for order <b>' + escHtml(target) + '</b>.';
                var lines2 = hits.map(function(r){
                    var p = getProductById(r.type);
                    return '• Order <b>' + escHtml(r.orderNo) + '</b> | ' + escHtml(p ? p.name : r.type) + ' | Frame: ' + escHtml(r.frameNo) + ' | By: ' + escHtml(r.inspector) + ' | ' + escHtml(r.timestamp ? r.timestamp.slice(0,10) : '');
                });
                return lines2.join('\n');
            }

            // QC count
            if (/how many qc|qc count|total qc|number of qc/.test(lq)) {
                return 'There are <b>' + qcRecs.length + ' QC reports</b> recorded locally.';
            }

            // Show QC records
            if (/show qc|list qc|all qc/.test(lq)) {
                if (qcRecs.length === 0) return 'No QC reports recorded locally yet. QC reports are saved when you Download PDF or Submit a QC Report.';
                var lines3 = qcRecs.slice(-10).reverse().map(function(q){
                    return '• <b>' + escHtml(q.orderNo || '—') + '</b> | ' + escHtml(q.templateTitle || q.template) + ' | Inspector: ' + escHtml(q.inspector) + ' | Date: ' + escHtml(q.qcDate || q.submittedAt.slice(0,10));
                });
                return 'Last ' + Math.min(qcRecs.length, 10) + ' QC reports (most recent first):\n' + lines3.join('\n');
            }

            // Recent entries
            if (/recent|latest|last (entry|entries|record|records)/.test(lq)) {
                if (records.length === 0) return 'No records yet.';
                var recent = records.slice(-5).reverse();
                var lines4 = recent.map(function(r){
                    var p = getProductById(r.type);
                    return '• <b>' + escHtml(r.orderNo) + '</b> | ' + escHtml(p ? p.name : r.type) + ' | ' + escHtml(r.timestamp ? r.timestamp.slice(0,10) : '');
                });
                return 'Last ' + recent.length + ' entries:\n' + lines4.join('\n');
            }

            return null;
        }

        function getBotAnswer(q) {
            var lq = q.toLowerCase().trim();

            // Navigation intent
            for (var n = 0; n < PAGE_NAV.length; n++) {
                for (var k = 0; k < PAGE_NAV[n].keys.length; k++) {
                    if (lq.indexOf(PAGE_NAV[n].keys[k]) !== -1) {
                        var pageId = PAGE_NAV[n].page;
                        var label = PAGE_NAV[n].label;
                        setTimeout(function(){ showPage(pageId); }, 800);
                        return 'Navigating you to the <b>' + label + '</b> tab now…';
                    }
                }
            }

            // Data queries
            var dataAnswer = handleDataQuery(lq);
            if (dataAnswer) return dataAnswer;

            // FAQ knowledge base
            for (var i = 0; i < KB.length; i++) {
                for (var j = 0; j < KB[i].keys.length; j++) {
                    if (lq.indexOf(KB[i].keys[j]) !== -1) return KB[i].answer;
                }
            }

            return "I'm not sure about that. Try asking:<br>• <i>\"How many orders?\"</i><br>• <i>\"Show all orders\"</i><br>• <i>\"Find order [order no]\"</i><br>• <i>\"Go to QC Reports\"</i><br>• Or type <b>help</b> for all topics.";
        }

        function botReply(q) {
            showTyping();
            setTimeout(function () {
                removeTyping();
                appendMsg(getBotAnswer(q), 'bot');
            }, 650);
        }

        function renderQuickReplies() {
            qrBox.innerHTML = '';
            QUICK.forEach(function (q) {
                var btn = document.createElement('button');
                btn.className = 'cb-qr-btn';
                btn.textContent = q;
                btn.addEventListener('click', function () {
                    appendMsg(q, 'user');
                    qrBox.innerHTML = '';
                    botReply(q);
                });
                qrBox.appendChild(btn);
            });
        }

        function openChat() {
            win.classList.remove('hidden');
            fab.classList.add('fab-open');
            if (!opened) {
                opened = true;
                setTimeout(function () {
                    appendMsg("Hi! I'm the <b>Neotrace Assistant</b>. I can look up your live order data or help you navigate the app.", 'bot');
                    renderQuickReplies();
                }, 200);
            }
            input.focus();
        }

        function closeChat() {
            win.classList.add('hidden');
            fab.classList.remove('fab-open');
        }

        function sendMessage() {
            var text = input.value.trim();
            if (!text) return;
            appendMsg(text, 'user');
            input.value = '';
            qrBox.innerHTML = '';
            botReply(text);
        }

        fab.addEventListener('click', function () {
            win.classList.contains('hidden') ? openChat() : closeChat();
        });
        closeBtn.addEventListener('click', closeChat);
        sendBtn.addEventListener('click', sendMessage);
        input.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') sendMessage();
        });
    })();

    // ===========================
    // Records Sub-tabs
    // ===========================
    document.querySelectorAll('.sub-tab-btn').forEach(function (btn) {
        btn.addEventListener('click', function () {
            var target = btn.getAttribute('data-subtab');
            // Update button active states
            document.querySelectorAll('.sub-tab-btn').forEach(function (b) { b.classList.remove('active'); });
            btn.classList.add('active');
            // Show correct panel
            document.querySelectorAll('.sub-tab-panel').forEach(function (p) { p.classList.remove('active'); });
            var panel = document.getElementById('subtab-' + target);
            if (panel) panel.classList.add('active');
            // If switching to detailed records, render it
            if (target === 'detailed-records') renderDetailedRecords();
        });
    });

    // ===========================
    // Detailed Records Sub-tab
    // ===========================
    var detTypeFilter = document.getElementById('det-type-filter');
    var detSearch = document.getElementById('det-search');
    var detDateFrom = document.getElementById('det-date-from');
    var detDateTo = document.getElementById('det-date-to');
    var detSearchBtn = document.getElementById('det-search-btn');
    var detClearBtn = document.getElementById('det-clear-btn');
    var detDownloadCsvBtn = document.getElementById('det-download-csv');
    var detSearchTriggered = false;

    function renderDetailedRecords() {
        var records = getRecords();
        var filterType = detTypeFilter ? detTypeFilter.value : 'all';
        var search = detSearch ? detSearch.value.trim().toLowerCase() : '';
        var dateFrom = detDateFrom && detDateFrom.value ? new Date(detDateFrom.value) : null;
        var dateTo = detDateTo && detDateTo.value ? new Date(detDateTo.value) : null;
        if (dateTo) dateTo.setHours(23, 59, 59, 999);

        var dynFields = [];
        if (filterType === 'all') {
            dynFields = buildDynamicHeaders();
        } else {
            var selProd = getProductById(filterType);
            dynFields = (selProd && selProd.fields) ? selProd.fields.slice() : [];
        }

        var tbody = document.getElementById('detailedRecordsBody');
        var noMsg = document.getElementById('noDetailedRecordsMsg');
        if (!tbody) return;

        // Rebuild thead
        var thead = document.querySelector('#detailedRecordsTable thead tr');
        if (thead) {
            thead.innerHTML = '<th>Print Label</th><th>QR Code</th><th>Type</th><th>Order No</th><th>Frame/Chassis No</th>';
            dynFields.forEach(function (f) { thead.innerHTML += '<th>' + esc(f.label) + '</th>'; });
            thead.innerHTML += '<th>Produced By</th><th>Timestamp</th>';
        }

        tbody.innerHTML = '';

        if (!detSearchTriggered && !search && !dateFrom && !dateTo) {
            noMsg.textContent = 'Enter a search term or date range to view detailed records.';
            noMsg.classList.remove('hidden');
            return;
        }

        var filtered = records.filter(function (r) {
            if (filterType !== 'all' && r.type !== filterType) return false;
            if (search) {
                var parts = [r.orderNo, r.frameNo, r.inspector];
                dynFields.forEach(function (f) { parts.push(r[f.key] || ''); });
                if (parts.join(' ').toLowerCase().indexOf(search) === -1) return false;
            }
            if (dateFrom || dateTo) {
                var ts = new Date(r.timestamp);
                if (dateFrom && ts < dateFrom) return false;
                if (dateTo && ts > dateTo) return false;
            }
            return true;
        });

        if (filtered.length === 0) {
            noMsg.textContent = 'No records found.';
            noMsg.classList.remove('hidden');
            return;
        }
        noMsg.classList.add('hidden');

        filtered.forEach(function (r) {
            var product = getProductById(r.type);
            var typeName = product ? product.name : r.type;
            var tr = document.createElement('tr');
            var html = '<td class="qr-print-cell"></td><td class="qr-cell"></td><td>' + esc(typeName) + '</td><td>' + esc(r.orderNo) + '</td><td>' + esc(r.frameNo) + '</td>';
            dynFields.forEach(function (f) { html += '<td>' + esc(r[f.key] || '-') + '</td>'; });
            html += '<td>' + esc(r.inspector) + '</td><td>' + formatDate(r.timestamp) + '</td>';
            tr.innerHTML = html;

            var printCell = tr.querySelector('.qr-print-cell');
            var qrCell = tr.querySelector('.qr-cell');
            (function (rec, cell) {
                var btn = document.createElement('button');
                btn.className = 'btn btn-secondary btn-sm';
                btn.textContent = 'Print QR';
                btn.addEventListener('click', function () { printLabelOnlyQR(rec); });
                cell.appendChild(btn);
            })(r, printCell);
            (function (rec, cell) {
                generateQR(rec.orderNo, 64, function (el) {
                    el.style.cursor = 'pointer';
                    el.title = 'Click to view order summary';
                    el.addEventListener('click', function () { showQRModal(rec); });
                    cell.appendChild(el);
                });
            })(r, qrCell);
            tbody.appendChild(tr);
        });
    }

    if (detSearchBtn) detSearchBtn.addEventListener('click', function () {
        detSearchTriggered = true;
        renderDetailedRecords();
    });
    if (detSearch) detSearch.addEventListener('input', function () {
        detSearchTriggered = false;
        renderDetailedRecords();
    });
    if (detTypeFilter) detTypeFilter.addEventListener('change', function () {
        if (detSearchTriggered) renderDetailedRecords();
    });
    if (detDateFrom) detDateFrom.addEventListener('change', function () {
        detSearchTriggered = true;
        renderDetailedRecords();
    });
    if (detDateTo) detDateTo.addEventListener('change', function () {
        detSearchTriggered = true;
        renderDetailedRecords();
    });
    if (detClearBtn) detClearBtn.addEventListener('click', function () {
        if (detSearch) detSearch.value = '';
        if (detDateFrom) detDateFrom.value = '';
        if (detDateTo) detDateTo.value = '';
        if (detTypeFilter) detTypeFilter.value = 'all';
        detSearchTriggered = false;
        renderDetailedRecords();
    });
    if (detDownloadCsvBtn) detDownloadCsvBtn.addEventListener('click', function () {
        var records = getRecords();
        var filterType = detTypeFilter ? detTypeFilter.value : 'all';
        var dynFields = filterType === 'all' ? buildDynamicHeaders() : ((getProductById(filterType) || {}).fields || []);
        var dateFrom = detDateFrom && detDateFrom.value ? new Date(detDateFrom.value) : null;
        var dateTo = detDateTo && detDateTo.value ? new Date(detDateTo.value) : null;
        if (dateTo) dateTo.setHours(23, 59, 59, 999);
        var search = detSearch ? detSearch.value.trim().toLowerCase() : '';

        var filtered = records.filter(function (r) {
            if (filterType !== 'all' && r.type !== filterType) return false;
            if (search) {
                var parts = [r.orderNo, r.frameNo, r.inspector];
                dynFields.forEach(function (f) { parts.push(r[f.key] || ''); });
                if (parts.join(' ').toLowerCase().indexOf(search) === -1) return false;
            }
            if (dateFrom) { if (new Date(r.timestamp) < dateFrom) return false; }
            if (dateTo) { if (new Date(r.timestamp) > dateTo) return false; }
            return true;
        });

        var headers = ['Type', 'Order No', 'Frame/Chassis No'];
        dynFields.forEach(function (f) { headers.push(f.label); });
        headers.push('Produced By', 'Timestamp');

        var csvRows = [headers.join(',')];
        filtered.forEach(function (r) {
            var product = getProductById(r.type);
            var row = [product ? product.name : r.type, r.orderNo, r.frameNo];
            dynFields.forEach(function (f) { row.push(r[f.key] || ''); });
            row.push(r.inspector, r.timestamp);
            csvRows.push(row.map(function (v) { return '"' + String(v || '').replace(/"/g, '""') + '"'; }).join(','));
        });

        var blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'detailed_records_' + new Date().toISOString().slice(0, 10) + '.csv';
        a.click();
        URL.revokeObjectURL(url);
    });

    // ===========================
    // Reports Page
    // ===========================
    function populateInspectorDropdowns() {
        var inspectors = getUniqueInspectors();
        var sel = document.getElementById('dash-inspector');
        if (!sel) return;
        var cur = sel.value;
        sel.innerHTML = '<option value="">All Inspectors</option>';
        inspectors.forEach(function (n) {
            var opt = document.createElement('option');
            opt.value = n; opt.textContent = n;
            sel.appendChild(opt);
        });
        sel.value = cur;
    }

    var currentReportData = [];

    function renderReport(records) {
        var dynFields = buildDynamicHeaders();
        var thead = document.querySelector('#reportTable thead tr');
        thead.innerHTML = '<th>Print Label</th><th>QR Code</th><th>Type</th><th>Order No</th><th>Frame/Chassis No</th>';
        dynFields.forEach(function (f) { thead.innerHTML += '<th>' + esc(f.label) + '</th>'; });
        thead.innerHTML += '<th>Produced By</th><th>Timestamp</th>';

        var tbody = document.getElementById('reportBody');
        var noMsg = document.getElementById('noReportMsg');
        tbody.innerHTML = '';

        if (records.length === 0) { noMsg.classList.remove('hidden'); return; }
        noMsg.classList.add('hidden');

        records.forEach(function (r) {
            var product = getProductById(r.type);
            var typeName = product ? product.name : r.type;
            var tr = document.createElement('tr');
            var html = '<td class="qr-print-cell"></td><td class="qr-cell"></td><td>' + esc(typeName) + '</td><td>' + esc(r.orderNo) + '</td><td>' + esc(r.frameNo) + '</td>';
            dynFields.forEach(function (f) { html += '<td>' + esc(r[f.key] || '-') + '</td>'; });
            html += '<td>' + esc(r.inspector) + '</td><td>' + formatDate(r.timestamp) + '</td>';
            tr.innerHTML = html;

            var printCell = tr.querySelector('.qr-print-cell');
            var qrCell = tr.querySelector('.qr-cell');
            (function (rec, cell) {
                var btn = document.createElement('button');
                btn.className = 'btn btn-secondary btn-sm';
                btn.textContent = 'Print QR';
                btn.addEventListener('click', function () { printLabelOnlyQR(rec); });
                cell.appendChild(btn);
            })(r, printCell);
            (function (rec, cell) {
                generateQR(rec.orderNo, 64, function (el) {
                    el.style.cursor = 'pointer';
                    el.title = 'Click to view order summary';
                    el.addEventListener('click', function () { showQRModal(rec); });
                    cell.appendChild(el);
                });
            })(r, qrCell);
            tbody.appendChild(tr);
        });
        currentReportData = records;
    }

    function applyReportFilters() {
        var search = document.getElementById('rpt-search').value.trim().toLowerCase();
        var dateFrom = document.getElementById('rpt-dateFrom').value;
        var dateTo = document.getElementById('rpt-dateTo').value;
        var records = getRecords();
        var dynFields = buildDynamicHeaders();

        var filtered = records.filter(function (r) {
            if (search) {
                var parts = [r.orderNo, r.frameNo, r.inspector];
                dynFields.forEach(function (f) { parts.push(r[f.key] || ''); });
                if (parts.join(' ').toLowerCase().indexOf(search) === -1) return false;
            }
            if (dateFrom || dateTo) {
                var d = dateOnly(r.timestamp);
                if (dateFrom && d < dateFrom) return false;
                if (dateTo && d > dateTo) return false;
            }
            return true;
        });
        renderReport(filtered);
    }

    document.getElementById('reportSearchBtn').addEventListener('click', applyReportFilters);
    document.getElementById('reportSearchClearBtn').addEventListener('click', function () {
        document.getElementById('rpt-search').value = '';
        applyReportFilters();
    });
    document.getElementById('rpt-search').addEventListener('keydown', function (e) { if (e.key === 'Enter') applyReportFilters(); });
    document.getElementById('rpt-dateFrom').addEventListener('change', applyReportFilters);
    document.getElementById('rpt-dateTo').addEventListener('change', applyReportFilters);

    document.getElementById('clearFilters').addEventListener('click', function () {
        document.getElementById('rpt-search').value = '';
        document.getElementById('rpt-dateFrom').value = '';
        document.getElementById('rpt-dateTo').value = '';
        renderReport(getRecords());
    });

    // CSV Download
    document.getElementById('downloadCSV').addEventListener('click', function () {
        var data = currentReportData.length > 0 ? currentReportData : getRecords();
        if (data.length === 0) { alert('No data to export.'); return; }
        var dynFields = buildDynamicHeaders();
        var headers = ['Type', 'Order No', 'Frame/Chassis No'];
        dynFields.forEach(function (f) { headers.push(f.label); });
        headers.push('Produced By', 'Timestamp');

        var rows = [headers.map(escapeCSV).join(',')];
        data.forEach(function (r) {
            var product = getProductById(r.type);
            var row = [escapeCSV(product ? product.name : r.type), escapeCSV(r.orderNo), escapeCSV(r.frameNo)];
            dynFields.forEach(function (f) { row.push(escapeCSV(r[f.key] || '')); });
            row.push(escapeCSV(r.inspector), escapeCSV(r.timestamp));
            rows.push(row.join(','));
        });

        var blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'traceability_report_' + new Date().toISOString().slice(0, 10) + '.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });

    // ===========================
    // QC Reports
    // ===========================

    function renderQCForm(templateKey) {
        var tplData = QC_TEMPLATE_DATA[templateKey];
        if (!tplData) {
            document.getElementById('qcSectionsContainer').innerHTML = '';
            document.getElementById('qcDynamicHeaders').innerHTML = '';
            document.getElementById('qc-docId').textContent = '—';
            return;
        }

        // Update document ID
        document.getElementById('qc-docId').textContent = tplData.docId || '—';

        // Render dynamic header fields
        var dynHdr = document.getElementById('qcDynamicHeaders');
        dynHdr.innerHTML = '';
        (tplData.headerFields || []).forEach(function (f) {
            var div = document.createElement('div');
            div.className = 'form-group';
            if (f.type === 'select' && f.options) {
                div.innerHTML = '<label for="qc-hdr-' + f.key + '">' + esc(f.label) + '</label>' +
                    '<select id="qc-hdr-' + f.key + '">' +
                    '<option value="">Select</option>' +
                    f.options.map(function (o) { return '<option value="' + esc(o) + '">' + esc(o) + '</option>'; }).join('') +
                    '</select>';
            } else {
                div.innerHTML = '<label for="qc-hdr-' + f.key + '">' + esc(f.label) + '</label>' +
                    '<input type="text" id="qc-hdr-' + f.key + '" placeholder="Enter ' + esc(f.label) + '">';
            }
            dynHdr.appendChild(div);
        });

        // Render checklist sections
        var container = document.getElementById('qcSectionsContainer');
        container.innerHTML = '';
        var globalIdx = 0;

        tplData.sections.forEach(function (section) {
            var card = document.createElement('div');
            card.className = 'qc-section-card';

            var hdr = document.createElement('h3');
            hdr.className = 'qc-section-title';
            hdr.textContent = section.name;
            card.appendChild(hdr);

            var tbl = document.createElement('table');
            tbl.className = 'qc-checklist-table';

            // Header row
            var isPacking = section.packing;
            var thead = document.createElement('thead');
            if (isPacking) {
                thead.innerHTML = '<tr><th>S.No</th><th>Box</th><th>Item</th><th>Checked</th><th>Remark</th></tr>';
            } else {
                thead.innerHTML = '<tr><th>S.No</th><th>Checking Parameter</th><th>Value / Selection</th><th>Remark</th></tr>';
            }
            tbl.appendChild(thead);

            var tbody = document.createElement('tbody');
            section.items.forEach(function (item) {
                var tr = document.createElement('tr');
                var id = 'qc-item-' + globalIdx;

                if (isPacking) {
                    tr.innerHTML =
                        '<td class="qc-sno">' + esc(String(item.n)) + '</td>' +
                        '<td class="qc-box">' + esc(item.box || '') + '</td>' +
                        '<td class="qc-param">' + esc(item.p) + '</td>' +
                        '<td class="qc-input">' + renderQCInput(item, id) + '</td>' +
                        '<td class="qc-remark"><input type="text" id="' + id + '-remark" placeholder="Remark"></td>';
                } else {
                    tr.innerHTML =
                        '<td class="qc-sno">' + esc(String(item.n)) + '</td>' +
                        '<td class="qc-param">' + esc(item.p) + '</td>' +
                        '<td class="qc-input">' + renderQCInput(item, id) + '</td>' +
                        '<td class="qc-remark"><input type="text" id="' + id + '-remark" placeholder="Remark"></td>';
                }

                tr.setAttribute('data-qc-idx', globalIdx);
                tr.setAttribute('data-qc-type', item.t);
                tr.setAttribute('data-qc-param', item.p);
                tbody.appendChild(tr);
                globalIdx++;
            });

            tbl.appendChild(tbody);
            var tw = document.createElement('div');
            tw.className = 'table-wrapper';
            tw.appendChild(tbl);
            card.appendChild(tw);
            container.appendChild(card);
        });
    }

    function renderQCInput(item, id) {
        var html = '';
        switch (item.t) {
            case 's': // select (radio)
                (item.o || []).forEach(function (opt, i) {
                    html += '<label class="qc-radio-label"><input type="radio" name="' + id + '" value="' + esc(opt) + '"> ' + esc(opt) + '</label>';
                });
                break;
            case 'so': // select + "Others" text
                (item.o || []).forEach(function (opt) {
                    html += '<label class="qc-radio-label"><input type="radio" name="' + id + '" value="' + esc(opt) + '"> ' + esc(opt) + '</label>';
                });
                html += '<label class="qc-radio-label"><input type="radio" name="' + id + '" value="__OTHER__"> Others:</label>';
                html += '<input type="text" id="' + id + '-other" class="qc-inline-text" placeholder="Specify">';
                break;
            case 't': // text
                html += '<input type="text" id="' + id + '-val" class="qc-text-input" placeholder="' + esc(item.u ? 'Enter value (' + item.u + ')' : 'Enter value') + '">';
                if (item.u) html += '<span class="qc-unit">' + esc(item.u) + '</span>';
                break;
            case 'd': // dual text
                (item.o || []).forEach(function (lbl, i) {
                    html += '<div class="qc-dual-field"><label class="qc-dual-label">' + esc(lbl) + '</label><input type="text" id="' + id + '-d' + i + '" class="qc-text-input qc-dual-input" placeholder="Value"></div>';
                });
                break;
            case 'y': // yes (single checkbox)
                html += '<label class="qc-check-label"><input type="checkbox" id="' + id + '-chk" value="Yes"> Yes</label>';
                break;
            case 'yn': // yes/no radio
                html += '<label class="qc-radio-label"><input type="radio" name="' + id + '" value="Yes"> Yes</label>';
                html += '<label class="qc-radio-label"><input type="radio" name="' + id + '" value="No"> No</label>';
                break;
            case 'ok': // ok check with standard display
                html += '<span class="qc-standard-text">' + esc(item.v || '') + '</span>';
                html += '<label class="qc-check-label qc-ok-label"><input type="checkbox" id="' + id + '-chk" value="OK"> OK</label>';
                break;
            default:
                html += '<input type="text" id="' + id + '-val" class="qc-text-input">';
        }
        return html;
    }

    function initQCForm() {
        var qcDate = document.getElementById('qc-date');
        if (qcDate && !qcDate.value) qcDate.value = new Date().toISOString().slice(0, 10);
        var tpl = document.getElementById('qc-template');
        if (tpl && tpl.value) renderQCForm(tpl.value);
    }

    function collectQCData() {
        var templateKey = document.getElementById('qc-template').value;
        var tplData = QC_TEMPLATE_DATA[templateKey];
        if (!tplData) return null;

        // Collect header fields
        var headerData = {};
        (tplData.headerFields || []).forEach(function (f) {
            var el = document.getElementById('qc-hdr-' + f.key);
            headerData[f.key] = el ? el.value.trim() : '';
        });

        // Collect checklist items
        var globalIdx = 0;
        var sections = [];
        tplData.sections.forEach(function (section) {
            var sectionItems = [];
            section.items.forEach(function (item) {
                var id = 'qc-item-' + globalIdx;
                var value = '';
                var checked = false;

                switch (item.t) {
                    case 's':
                    case 'yn':
                        var radio = document.querySelector('input[name="' + id + '"]:checked');
                        value = radio ? radio.value : '';
                        checked = !!value;
                        break;
                    case 'so':
                        var radio2 = document.querySelector('input[name="' + id + '"]:checked');
                        if (radio2) {
                            if (radio2.value === '__OTHER__') {
                                var otherEl = document.getElementById(id + '-other');
                                value = 'Others: ' + (otherEl ? otherEl.value.trim() : '');
                            } else {
                                value = radio2.value;
                            }
                            checked = true;
                        }
                        break;
                    case 't':
                        var txtEl = document.getElementById(id + '-val');
                        value = txtEl ? txtEl.value.trim() : '';
                        checked = !!value;
                        break;
                    case 'd':
                        var parts = [];
                        (item.o || []).forEach(function (lbl, i) {
                            var dEl = document.getElementById(id + '-d' + i);
                            parts.push(lbl + ': ' + (dEl ? dEl.value.trim() : ''));
                        });
                        value = parts.join(' | ');
                        checked = parts.some(function (p) { return p.indexOf(': ') < p.length - 2; });
                        break;
                    case 'y':
                    case 'ok':
                        var chkEl = document.getElementById(id + '-chk');
                        checked = chkEl ? chkEl.checked : false;
                        value = checked ? (item.t === 'ok' ? 'OK' : 'Yes') : '';
                        break;
                }

                var remarkEl = document.getElementById(id + '-remark');
                sectionItems.push({
                    no: item.n,
                    param: item.p,
                    type: item.t,
                    value: value,
                    checked: checked,
                    standard: item.v || (item.o ? item.o.join(' / ') : ''),
                    remark: remarkEl ? remarkEl.value.trim() : '',
                    box: item.box || ''
                });
                globalIdx++;
            });
            sections.push({ name: section.name, packing: !!section.packing, items: sectionItems });
        });

        return {
            id: 'qc_' + Date.now(),
            template: templateKey,
            templateTitle: tplData.title,
            docId: tplData.docId,
            orderNo: document.getElementById('qc-orderNo').value.trim(),
            customerName: (document.getElementById('qc-customerName') || { value: '' }).value.trim(),
            frameNo: document.getElementById('qc-frameNo').value.trim(),
            inspector: document.getElementById('qc-inspector').value.trim(),
            qcDate: document.getElementById('qc-date').value,
            qcStartTime: (document.getElementById('qc-startTime') || { value: '' }).value,
            qcEndTime: (document.getElementById('qc-endTime') || { value: '' }).value,
            headerData: headerData,
            sections: sections,
            notes: (document.getElementById('qc-notes') || { value: '' }).value.trim(),
            submittedAt: new Date().toISOString()
        };
    }

    function submitQCToGoogleDrive(payload) {
        payload = payload || {};
        payload.action = 'qc_submit';
        payload.sheetId = SHEET_ID;
        payload.gid = SHEET_GID;
        return postSheetPayload(payload);
    }

    // ---- Download QC Report as PDF (client-side) ----
    function downloadQCPdf() {
        var payload = collectQCData();
        var msg = document.getElementById('qc-confirmation');
        if (!payload) {
            msg.textContent = 'Please select a QC template first.';
            msg.className = 'confirmation confirmation-error';
            return;
        }

        if (!window.jspdf || !window.jspdf.jsPDF) {
            msg.textContent = 'PDF library not loaded. Please reload the page and try again.';
            msg.className = 'confirmation confirmation-error';
            return;
        }

        var jsPDF = window.jspdf.jsPDF;
        var doc = new jsPDF('p', 'mm', 'a4');
        var pageWidth = doc.internal.pageSize.getWidth();
        var margin = 14;
        var y = 15;

        // Title
        doc.setFontSize(16);
        doc.setFont('helvetica', 'bold');
        doc.text(payload.templateTitle || 'Neotrace QC Report', pageWidth / 2, y, { align: 'center' });
        y += 7;

        // Document ID
        doc.setFontSize(9);
        doc.setFont('helvetica', 'italic');
        doc.setTextColor(100);
        doc.text('Document ID: ' + (payload.docId || '-'), pageWidth / 2, y, { align: 'center' });
        doc.setTextColor(0);
        y += 8;

        // Header info table
        doc.setFont('helvetica', 'normal');
        var hdrRows = [
            ['Order No:', payload.orderNo || '-', 'Customer Name:', payload.customerName || '-'],
            ['Date:', payload.qcDate || '-', 'QC Person:', payload.inspector || '-'],
            ['Frame/Chassis No:', payload.frameNo || '-', 'QC Start Time:', payload.qcStartTime || '-'],
            ['', '', 'QC End Time:', payload.qcEndTime || '-']
        ];

        // Add template-specific header data
        var hdrKeys = Object.keys(payload.headerData || {});
        for (var h = 0; h < hdrKeys.length; h += 2) {
            var row = [hdrKeys[h] + ':', payload.headerData[hdrKeys[h]] || '-'];
            if (h + 1 < hdrKeys.length) {
                row.push(hdrKeys[h + 1] + ':');
                row.push(payload.headerData[hdrKeys[h + 1]] || '-');
            } else {
                row.push('');
                row.push('');
            }
            hdrRows.push(row);
        }

        doc.autoTable({
            startY: y,
            body: hdrRows,
            theme: 'grid',
            margin: { left: margin, right: margin },
            styles: { fontSize: 9, cellPadding: 2 },
            columnStyles: {
                0: { fontStyle: 'bold', cellWidth: 38 },
                1: { cellWidth: 48 },
                2: { fontStyle: 'bold', cellWidth: 38 },
                3: { cellWidth: 48 }
            },
            didParseCell: function (data) {
                if (data.column.index === 0 || data.column.index === 2) {
                    data.cell.styles.fillColor = [245, 245, 245];
                }
            }
        });
        y = doc.lastAutoTable.finalY + 6;

        // Sections
        (payload.sections || []).forEach(function (section) {
            // Check if we need a new page
            if (y > doc.internal.pageSize.getHeight() - 30) {
                doc.addPage();
                y = 15;
            }

            // Section heading
            doc.setFontSize(12);
            doc.setFont('helvetica', 'bold');
            doc.text(section.name, margin, y);
            y += 2;

            if (section.packing) {
                // Packing checklist
                var packHead = [['S.No', 'Box', 'Item', 'Checked', 'Remark']];
                var packBody = section.items.map(function (item) {
                    return [
                        String(item.no || ''),
                        item.box || '',
                        item.param || '',
                        item.checked ? 'YES' : (item.value || '-'),
                        item.remark || ''
                    ];
                });
                doc.autoTable({
                    startY: y,
                    head: packHead,
                    body: packBody,
                    theme: 'grid',
                    margin: { left: margin, right: margin },
                    styles: { fontSize: 8, cellPadding: 2 },
                    headStyles: { fillColor: [204, 0, 0], textColor: 255, fontStyle: 'bold' }
                });
            } else {
                // QC checklist
                var qcHead = [['S.No', 'Checking Parameter', 'Value / Status', 'Remark']];
                var qcBody = section.items.map(function (item) {
                    var displayValue = item.value || '-';
                    if (item.type === 'ok' && item.checked) {
                        displayValue = 'OK (' + (item.standard || '') + ')';
                    } else if (item.type === 'y' && item.checked) {
                        displayValue = 'Yes';
                    } else if (item.type === 'yn') {
                        displayValue = item.value || '-';
                    }
                    return [
                        String(item.no || ''),
                        item.param || '',
                        displayValue,
                        item.remark || ''
                    ];
                });
                doc.autoTable({
                    startY: y,
                    head: qcHead,
                    body: qcBody,
                    theme: 'grid',
                    margin: { left: margin, right: margin },
                    styles: { fontSize: 8, cellPadding: 2 },
                    headStyles: { fillColor: [204, 0, 0], textColor: 255, fontStyle: 'bold' },
                    columnStyles: {
                        0: { cellWidth: 14 },
                        1: { cellWidth: 60 },
                        3: { cellWidth: 35 }
                    }
                });
            }
            y = doc.lastAutoTable.finalY + 6;
        });

        // Notes
        if (payload.notes) {
            if (y > doc.internal.pageSize.getHeight() - 30) {
                doc.addPage();
                y = 15;
            }
            doc.setFontSize(12);
            doc.setFont('helvetica', 'bold');
            doc.text('Overall Notes', margin, y);
            y += 5;
            doc.setFontSize(9);
            doc.setFont('helvetica', 'normal');
            var noteLines = doc.splitTextToSize(payload.notes, pageWidth - margin * 2);
            doc.text(noteLines, margin, y);
            y += noteLines.length * 4 + 4;
        }

        // Footer
        if (y > doc.internal.pageSize.getHeight() - 15) {
            doc.addPage();
            y = 15;
        }
        doc.setFontSize(8);
        doc.setFont('helvetica', 'italic');
        doc.setTextColor(136);
        doc.text('Generated: ' + new Date().toISOString() + '  |  Neotrace', pageWidth / 2, doc.internal.pageSize.getHeight() - 8, { align: 'center' });
        doc.setTextColor(0);

        // Download via blob so the browser shows a native Save dialog
        var fileName = 'QC_' + (payload.template || 'UNKNOWN') + '_' + (payload.orderNo || 'NO_ORDER') + '_' + (payload.qcDate || new Date().toISOString().slice(0, 10)) + '.pdf';
        var blob = doc.output('blob');
        var blobUrl = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = blobUrl;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        setTimeout(function () {
            document.body.removeChild(a);
            URL.revokeObjectURL(blobUrl);
        }, 1000);

        // Save QC record locally for dashboard charts
        saveQCRecord(payload);

        msg.textContent = 'PDF saved: ' + fileName;
        msg.className = 'confirmation confirmation-success';
        setTimeout(function () { msg.className = 'confirmation hidden'; }, 4000);
    }

    var downloadBtn = document.getElementById('downloadQcPdf');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', downloadQCPdf);
    }

    var qcTemplateEl = document.getElementById('qc-template');
    if (qcTemplateEl) {
        qcTemplateEl.addEventListener('change', function () {
            renderQCForm(this.value);
        });
    }

    var qcForm = document.getElementById('qcForm');
    if (qcForm) {
        qcForm.addEventListener('submit', function (e) {
            e.preventDefault();
            var msg = document.getElementById('qc-confirmation');
            var payload = collectQCData();
            if (!payload) {
                msg.textContent = 'Please select a QC template.';
                msg.className = 'confirmation confirmation-error';
                return;
            }

            // Validate: check that at least one section has responses
            var totalItems = 0;
            var filledItems = 0;
            payload.sections.forEach(function (sec) {
                sec.items.forEach(function (item) {
                    totalItems++;
                    if (item.checked || item.value) filledItems++;
                });
            });

            if (filledItems < totalItems * 0.5) {
                msg.textContent = 'Please complete at least 50% of the checklist items before submitting.';
                msg.className = 'confirmation confirmation-error';
                return;
            }

            msg.textContent = 'Submitting QC report...';
            msg.className = 'confirmation confirmation-success';

            saveQCRecord(payload);
            submitQCToGoogleDrive(payload).then(function () {
                msg.textContent = 'QC report submitted successfully! PDF is being generated and saved to Google Drive.';
                msg.className = 'confirmation confirmation-success';
                setTimeout(function () { msg.className = 'confirmation hidden'; }, 6000);
                qcForm.reset();
                document.getElementById('qcSectionsContainer').innerHTML = '';
                document.getElementById('qcDynamicHeaders').innerHTML = '';
                document.getElementById('qc-docId').textContent = '—';
                initQCForm();
            }).catch(function () {
                msg.textContent = 'Unable to submit QC report. Check Google webhook URL in Admin settings.';
                msg.className = 'confirmation confirmation-error';
            });
        });
    }

    // ===========================
    // Dashboard with Charts
    // ===========================
    var pieChartInstance = null;
    var barChartInstance = null;
    var qcPieChartInstance = null;
    var qcBarChartInstance = null;
    var CHART_COLORS = ['#cc0000', '#1a7d2f', '#2563eb', '#d97706', '#7c3aed', '#0891b2', '#be185d', '#65a30d', '#ea580c', '#6366f1'];

    function refreshDashboard() {
        var records = getRecords();
        var products = getProducts();
        var inspFilter = document.getElementById('dash-inspector').value;
        var dateFrom = document.getElementById('dash-dateFrom').value;
        var dateTo = document.getElementById('dash-dateTo').value;

        var filtered = records.filter(function (r) {
            if (inspFilter && r.inspector !== inspFilter) return false;
            var d = dateOnly(r.timestamp);
            if (dateFrom && d < dateFrom) return false;
            if (dateTo && d > dateTo) return false;
            return true;
        });

        // Summary cards
        var productCounts = {};
        products.forEach(function (p) { productCounts[p.id] = 0; });
        var uniqueInsp = {};
        filtered.forEach(function (r) {
            if (productCounts.hasOwnProperty(r.type)) productCounts[r.type]++;
            uniqueInsp[r.inspector] = true;
        });

        var cardsContainer = document.getElementById('dashboardCards');
        cardsContainer.innerHTML = '';
        var cards = [{ value: filtered.length, label: 'Total Entries' }];
        products.forEach(function (p) { cards.push({ value: productCounts[p.id] || 0, label: p.name + ' Entries' }); });
        cards.push({ value: Object.keys(uniqueInsp).length, label: 'Active Producers' });
        cards.forEach(function (c) {
            var div = document.createElement('div');
            div.className = 'dash-card';
            div.innerHTML = '<div class="card-value">' + c.value + '</div><div class="card-label">' + c.label + '</div>';
            cardsContainer.appendChild(div);
        });

        // Pie Chart - Inspector Contribution
        var inspCounts = {};
        filtered.forEach(function (r) { inspCounts[r.inspector] = (inspCounts[r.inspector] || 0) + 1; });
        var pieLabels = Object.keys(inspCounts).sort();
        var pieData = pieLabels.map(function (l) { return inspCounts[l]; });

        if (pieChartInstance) pieChartInstance.destroy();
        var pieCtx = document.getElementById('pieChart').getContext('2d');
        if (pieLabels.length > 0) {
            pieChartInstance = new Chart(pieCtx, {
                type: 'pie',
                data: { labels: pieLabels, datasets: [{ data: pieData, backgroundColor: pieLabels.map(function (_, i) { return CHART_COLORS[i % CHART_COLORS.length]; }) }] },
                options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, padding: 10 } } } }
            });
        } else {
            pieChartInstance = new Chart(pieCtx, {
                type: 'pie',
                data: { labels: ['No data'], datasets: [{ data: [1], backgroundColor: ['#e0e0e0'] }] },
                options: { responsive: true, plugins: { legend: { display: false } } }
            });
        }

        // Bar Chart - Day-wise Production
        var dayCounts = {};
        filtered.forEach(function (r) { var d = dateOnly(r.timestamp); if (d) dayCounts[d] = (dayCounts[d] || 0) + 1; });
        var dayKeys = Object.keys(dayCounts).sort();
        var barLabels = dayKeys.map(function (d) { return formatDateDDMMYY(d); });
        var barData = dayKeys.map(function (d) { return dayCounts[d]; });

        if (barChartInstance) barChartInstance.destroy();
        var barCtx = document.getElementById('barChart').getContext('2d');
        barChartInstance = new Chart(barCtx, {
            type: 'bar',
            data: { labels: barLabels, datasets: [{ label: 'Entries', data: barData, backgroundColor: '#2563eb', borderRadius: 4 }] },
            options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } }, x: { ticks: { maxRotation: 45 } } } }
        });

        // Daily breakdown table
        var dashHead = document.querySelector('#dashboardTable thead tr');
        dashHead.innerHTML = '<th>Produced By</th><th>Date</th>';
        products.forEach(function (p) { dashHead.innerHTML += '<th>' + esc(p.name) + ' Count</th>'; });
        dashHead.innerHTML += '<th>Total Produced</th>';

        var breakdown = {};
        filtered.forEach(function (r) {
            var d = dateOnly(r.timestamp);
            var key = r.inspector + '|' + d;
            if (!breakdown[key]) {
                breakdown[key] = { inspector: r.inspector, date: d, counts: {}, total: 0 };
                products.forEach(function (p) { breakdown[key].counts[p.id] = 0; });
            }
            if (breakdown[key].counts.hasOwnProperty(r.type)) breakdown[key].counts[r.type]++;
            breakdown[key].total++;
        });

        var bRows = Object.keys(breakdown).map(function (k) { return breakdown[k]; });
        bRows.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : a.inspector.localeCompare(b.inspector); });

        var tbody = document.getElementById('dashboardBody');
        tbody.innerHTML = '';
        bRows.forEach(function (row) {
            var tr = document.createElement('tr');
            var html = '<td>' + esc(row.inspector) + '</td><td>' + esc(formatDateDDMMYY(row.date)) + '</td>';
            products.forEach(function (p) { html += '<td>' + (row.counts[p.id] || 0) + '</td>'; });
            html += '<td>' + row.total + '</td>';
            tr.innerHTML = html;
            tbody.appendChild(tr);
        });

        // QC Charts — filter QC records by same date range
        var qcRecords = getQCRecords().filter(function (q) {
            var d = (q.qcDate || q.submittedAt || '').slice(0, 10);
            if (dateFrom && d < dateFrom) return false;
            if (dateTo && d > dateTo) return false;
            return true;
        });

        // QC Pie — count by inspector
        var qcInspCounts = {};
        qcRecords.forEach(function (q) {
            var k = q.inspector || 'Unknown';
            qcInspCounts[k] = (qcInspCounts[k] || 0) + 1;
        });
        var qcPieLabels = Object.keys(qcInspCounts).sort();
        var qcPieData = qcPieLabels.map(function (l) { return qcInspCounts[l]; });

        if (qcPieChartInstance) qcPieChartInstance.destroy();
        var qcPieCtx = document.getElementById('qcPieChart').getContext('2d');
        if (qcPieLabels.length > 0) {
            qcPieChartInstance = new Chart(qcPieCtx, {
                type: 'pie',
                data: { labels: qcPieLabels, datasets: [{ data: qcPieData, backgroundColor: qcPieLabels.map(function (_, i) { return CHART_COLORS[i % CHART_COLORS.length]; }) }] },
                options: { responsive: true, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, padding: 10 } } } }
            });
        } else {
            qcPieChartInstance = new Chart(qcPieCtx, {
                type: 'pie',
                data: { labels: ['No QC data'], datasets: [{ data: [1], backgroundColor: ['#e0e0e0'] }] },
                options: { responsive: true, plugins: { legend: { display: false } } }
            });
        }

        // QC Bar — count by day
        var qcDayCounts = {};
        qcRecords.forEach(function (q) {
            var d = (q.qcDate || q.submittedAt || '').slice(0, 10);
            if (d) qcDayCounts[d] = (qcDayCounts[d] || 0) + 1;
        });
        var qcDayKeys = Object.keys(qcDayCounts).sort();

        if (qcBarChartInstance) qcBarChartInstance.destroy();
        var qcBarCtx = document.getElementById('qcBarChart').getContext('2d');
        qcBarChartInstance = new Chart(qcBarCtx, {
            type: 'bar',
            data: { labels: qcDayKeys.map(function (d) { return formatDateDDMMYY(d); }), datasets: [{ label: 'QC Reports', data: qcDayKeys.map(function (d) { return qcDayCounts[d]; }), backgroundColor: '#1a7d2f', borderRadius: 4 }] },
            options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } }, x: { ticks: { maxRotation: 45 } } } }
        });
    }

    function clearDashboardView() {
        document.getElementById('dashboardCards').innerHTML = '';
        document.getElementById('dashboardBody').innerHTML = '';
        if (pieChartInstance) { pieChartInstance.destroy(); pieChartInstance = null; }
        if (barChartInstance) { barChartInstance.destroy(); barChartInstance = null; }
        if (qcPieChartInstance) { qcPieChartInstance.destroy(); qcPieChartInstance = null; }
        if (qcBarChartInstance) { qcBarChartInstance.destroy(); qcBarChartInstance = null; }
    }

    function clearDashboardFiltersAndView() {
        document.getElementById('dash-inspector').value = '';
        document.getElementById('dash-dateFrom').value = '';
        document.getElementById('dash-dateTo').value = '';
        clearDashboardView();
    }

    document.getElementById('loadDashboard').addEventListener('click', refreshDashboard);
    document.getElementById('clearDashboard').addEventListener('click', clearDashboardFiltersAndView);

    // ===========================
    // Admin Page
    // ===========================
    document.getElementById('adminLoginBtn').addEventListener('click', function () {
        var pw = document.getElementById('adminPassword').value;
        if (pw === getAdminPassword()) {
            document.getElementById('adminLoginSection').classList.add('hidden');
            document.getElementById('adminPanel').classList.remove('hidden');
            document.getElementById('adminError').classList.add('hidden');
            document.getElementById('adminPassword').value = '';
            loadSheetWebhookConfig();
        } else {
            document.getElementById('adminError').classList.remove('hidden');
        }
    });

    document.getElementById('adminPassword').addEventListener('keydown', function (e) {
        if (e.key === 'Enter') document.getElementById('adminLoginBtn').click();
    });

    document.getElementById('adminLogout').addEventListener('click', function () {
        document.getElementById('adminPanel').classList.add('hidden');
        document.getElementById('adminLoginSection').classList.remove('hidden');
        document.getElementById('adminResults').innerHTML = '';
        document.getElementById('editFormContainer').classList.add('hidden');
        document.getElementById('changePasswordSection').classList.add('hidden');
    });

    document.getElementById('changePasswordBtn').addEventListener('click', function () {
        document.getElementById('changePasswordSection').classList.toggle('hidden');
    });

    document.getElementById('saveNewPassword').addEventListener('click', function () {
        var cur = document.getElementById('currentPassword').value;
        var newPw = document.getElementById('newPassword').value;
        var confirmPw = document.getElementById('confirmNewPassword').value;
        var msg = document.getElementById('passwordMsg');

        if (cur !== getAdminPassword()) { msg.textContent = 'Current password is incorrect.'; msg.className = 'error-text'; msg.classList.remove('hidden'); return; }
        if (!newPw || newPw.length < 4) { msg.textContent = 'New password must be at least 4 characters.'; msg.className = 'error-text'; msg.classList.remove('hidden'); return; }
        if (newPw !== confirmPw) { msg.textContent = 'Passwords do not match.'; msg.className = 'error-text'; msg.classList.remove('hidden'); return; }

        setAdminPassword(newPw);
        msg.textContent = 'Password changed successfully.';
        msg.className = 'success-text';
        msg.classList.remove('hidden');
        document.getElementById('currentPassword').value = '';
        document.getElementById('newPassword').value = '';
        document.getElementById('confirmNewPassword').value = '';
        setTimeout(function () { msg.classList.add('hidden'); }, 3000);
    });


    function loadSheetWebhookConfig() {
        var input = document.getElementById('sheetWebhookUrl');
        if (input) input.value = getSheetWebhookUrl();
    }

    var saveSheetWebhookBtn = document.getElementById('saveSheetWebhook');
    if (saveSheetWebhookBtn) {
        saveSheetWebhookBtn.addEventListener('click', function () {
            var input = document.getElementById('sheetWebhookUrl');
            var msg = document.getElementById('sheetWebhookMsg');
            var value = input ? input.value.trim() : '';
            setSheetWebhookUrl(value);
            if (msg) {
                msg.textContent = value ? 'Google Sheet webhook URL saved.' : 'Google Sheet webhook URL cleared.';
                msg.className = 'success-text';
                msg.classList.remove('hidden');
                setTimeout(function () { msg.classList.add('hidden'); }, 2500);
            }
        });
    }


    var syncAllBtn = document.getElementById('syncAllToSheet');
    if (syncAllBtn) {
        syncAllBtn.addEventListener('click', function () {
            var msg = document.getElementById('sheetWebhookMsg');
            var url = getSheetWebhookUrl();
            if (!url) {
                if (msg) {
                    msg.textContent = 'Please save a valid Apps Script Web App URL first.';
                    msg.className = 'error-text';
                    msg.classList.remove('hidden');
                }
                return;
            }

            syncAllBtn.disabled = true;
            syncAllBtn.textContent = 'Syncing...';
            syncAllRecordsToGoogleSheet().then(function (res) {
                if (msg) {
                    msg.textContent = 'Sync request sent for ' + res.sent + ' of ' + res.total + ' records.';
                    msg.className = 'success-text';
                    msg.classList.remove('hidden');
                }
            }).finally(function () {
                syncAllBtn.disabled = false;
                syncAllBtn.textContent = 'Sync All Records Now';
            });
        });
    }

    // Admin search
    document.getElementById('adminSearchBtn').addEventListener('click', adminSearch);
    document.getElementById('admin-search').addEventListener('keydown', function (e) { if (e.key === 'Enter') adminSearch(); });

    function adminSearch() {
        var term = document.getElementById('admin-search').value.trim().toLowerCase();
        var records = getRecords();
        var container = document.getElementById('adminResults');
        container.innerHTML = '';
        document.getElementById('editFormContainer').classList.add('hidden');

        if (!term) { container.innerHTML = '<p>Enter a search term.</p>'; return; }

        var matches = records.filter(function (r) {
            return r.orderNo.toLowerCase().indexOf(term) !== -1 || r.frameNo.toLowerCase().indexOf(term) !== -1;
        });

        if (matches.length === 0) { container.innerHTML = '<p>No records found.</p>'; return; }

        matches.forEach(function (r) {
            var product = getProductById(r.type);
            var typeName = product ? product.name : r.type;
            var card = document.createElement('div');
            card.className = 'admin-record-card';

            var info = document.createElement('div');
            info.className = 'admin-record-info';
            info.innerHTML = '<strong>' + esc(typeName.toUpperCase()) + '</strong> | Order: ' + esc(r.orderNo) +
                ' | Frame: ' + esc(r.frameNo) + ' | Produced by: ' + esc(r.inspector) + ' | ' + formatDate(r.timestamp);

            var actions = document.createElement('div');
            actions.className = 'admin-record-actions';

            var editBtn = document.createElement('button');
            editBtn.className = 'btn btn-secondary btn-sm';
            editBtn.textContent = 'Edit';
            editBtn.addEventListener('click', function () { loadEditForm(r); });

            var delBtn = document.createElement('button');
            delBtn.className = 'btn btn-danger btn-sm';
            delBtn.textContent = 'Delete';
            delBtn.addEventListener('click', function () {
                if (!confirm('Delete record for Order: ' + r.orderNo + '? This cannot be undone.')) return;
                deleteRecord(r.id);
                adminSearch();
            });

            actions.appendChild(editBtn);
            actions.appendChild(delBtn);
            card.appendChild(info);
            card.appendChild(actions);
            container.appendChild(card);
        });
    }

    function loadEditForm(record) {
        document.getElementById('editFormContainer').classList.remove('hidden');
        document.getElementById('edit-id').value = record.id;
        document.getElementById('edit-orderNo').value = record.orderNo;
        document.getElementById('edit-frameNo').value = record.frameNo;
        document.getElementById('edit-inspector').value = record.inspector;
        if (record.timestamp) document.getElementById('edit-timestamp').value = record.timestamp;

        var product = getProductById(record.type);
        var container = document.getElementById('editDynamicFields');
        container.innerHTML = '';
        if (product && product.fields) {
            product.fields.forEach(function (f) {
                var div = document.createElement('div');
                div.className = 'form-group';
                div.innerHTML = '<label for="edit-' + esc(f.key) + '">' + esc(f.label) + '</label>' +
                    '<input type="text" id="edit-' + esc(f.key) + '" value="' + esc(record[f.key] || '') + '">';
                container.appendChild(div);
            });
        }
        document.getElementById('editFormContainer').scrollIntoView({ behavior: 'smooth' });
    }

    document.getElementById('editRecordForm').addEventListener('submit', function (e) {
        e.preventDefault();
        var id = document.getElementById('edit-id').value;
        var updates = {
            orderNo: document.getElementById('edit-orderNo').value.trim(),
            frameNo: document.getElementById('edit-frameNo').value.trim(),
            inspector: document.getElementById('edit-inspector').value.trim(),
            timestamp: document.getElementById('edit-timestamp').value
        };

        var record = getRecords().filter(function (r) { return r.id === id; })[0];
        if (record) {
            var product = getProductById(record.type);
            if (product && product.fields) {
                product.fields.forEach(function (f) {
                    var inp = document.getElementById('edit-' + f.key);
                    updates[f.key] = inp ? inp.value.trim() : '';
                });
            }
        }

        var result = updateRecord(id, updates);
        if (result) {
            alert('Record updated successfully.');
            document.getElementById('editFormContainer').classList.add('hidden');
            adminSearch();
        } else {
            alert('Error: Record not found.');
        }
    });

    document.getElementById('cancelEdit').addEventListener('click', function () {
        document.getElementById('editFormContainer').classList.add('hidden');
    });

    // ===========================
    // Init
    // ===========================
    setDefaultTimestamp('entry-timestamp');
    populateEntryProductDropdown();
    renderRecords();
    loadSheetWebhookConfig();

})();
