const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fetch = require('node-fetch');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── CONFIG ──────────────────────────────────────────────────────
const CONFIG = {
    baseUrl: 'https://thesleepcompany.unicommerce.co.in',
    token: '229e9495-4148-4dda-8f03-7fb996f49aa8',
    delayMs: 100
};

// ─── Helpers ─────────────────────────────────────────────────────
function excelDateToString(val) {
    if (!val) return '';

    let date;
    if (typeof val === 'string') {
        const cleaned = val.replace(/\//g, '-').trim();
        const parts = cleaned.split(/[- :]/);
        if (parts.length >= 3) {
            const day = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1;
            const year = parseInt(parts[2]);
            const hour = parseInt(parts[3]) || 0;
            const min = parseInt(parts[4]) || 0;
            const sec = parseInt(parts[5]) || 0;
            date = new Date(year, month, day, hour, min, sec);
        } else {
            date = new Date(val);
        }
    } else {
        date = new Date(Math.round((val - 25569) * 86400 * 1000));
    }

    if (isNaN(date.getTime())) return '';

    const pad = n => String(n).padStart(2, '0');
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T` +
        `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

const num = v => parseFloat(v) || 0;
const str = v => (v === undefined || v === null) ? '' : String(v).trim();
const bool = v => v === true || v === 'true' || v === 1 || v === '1' || v === 'yes';

const pipes = v => str(v).split('|').map(s => s.trim());
const pipeNums = v => pipes(v).map(s => num(s));
const pipeBools = v => pipes(v).map(s => bool(s));

// ─── Parse uploaded file ────────────────────────────────────────
function parseUpload(filePath) {
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

// ─── Row → Unicommerce Payload (Final Fixed Version) ─────────────────────
function rowToPayload(row) {
    const displayOrderCode = str(row['Display Sales Order Code'] || row['Display Order Code'] || row['displayOrderCode']);

    if (!displayOrderCode) {
        throw new Error('Display Sales Order Code is required');
    }

    const orderCode = str(row['Sales Order Code*'] || row['code']) || displayOrderCode;

    const shippingId = str(row['Shipping Address Id']) || ('SHP_' + displayOrderCode);
    const billingId = str(row['Billing Address Id']) || ('BIL_' + displayOrderCode);

    // ==================== ADDRESSES ====================
    const shippingAddress = {
        id: shippingId,
        name: str(row['Shipping Address Name'] || row['customerName']),
        addressLine1: str(row['Shipping Address Line 1']),
        addressLine2: str(row['Shipping Address Line 2']),
        city: str(row['Shipping Address City']),
        state: str(row['Shipping Address State']),
        country: str(row['Shipping Address Country']) || 'IN',
        pincode: str(row['Shipping Address Pincode']),
        phone: str(row['Shipping Address Phone'])
    };

    const billingAddress = {
        id: billingId,
        name: str(row['Billing Address Name']) || shippingAddress.name,
        addressLine1: str(row['Billing Address Line 1']) || shippingAddress.addressLine1,
        addressLine2: str(row['Billing Address Line 2']) || shippingAddress.addressLine2,
        city: str(row['Billing Address City']) || shippingAddress.city,
        state: str(row['Billing Address State']) || shippingAddress.state,
        country: str(row['Billing Address Country']) || shippingAddress.country,
        pincode: str(row['Billing Address Pincode']) || shippingAddress.pincode,
        phone: str(row['Billing Address Phone']) || shippingAddress.phone
    };

    // ==================== ITEMS ====================
    const itemSkus = pipes(row['Item SKU Code*'] || row['itemSku']);
    if (itemSkus.length === 0 || !itemSkus[0]) {
        throw new Error('Item SKU Code* is required');
    }

    const itemCodes = pipes(row['Sale Order Item Code'] || row['code']);
    const facilityCodes = pipes(row['Facility Code']);
    const shippingMethods = pipes(row['Shipping Method*']);
    const channelProductIds = pipes(row['Channel Product Id']);
    const sellingPrices = pipeNums(row['Selling Price']);
    const totalPrices = pipeNums(row['Total Price']);
    const discounts = pipeNums(row['Discount']);
    const shippingChargesArr = pipeNums(row['Shipping Charges']);
    const giftWrapChargesArr = pipeNums(row['Gift Wrap Charges']);
    const prepaidAmounts = pipeNums(row['Prepaid Amount']);
    const packetNumbers = pipeNums(row['Packet Number']);
    const giftWraps = pipeBools(row['Gift Wrap']);
    const giftMessages = pipes(row['Gift Message']);
    const itemTags = pipes(row['Item Tag']);

    const pick = (arr, i) => (arr[i] !== undefined && arr[i] !== '') ? arr[i] : (arr[0] !== undefined ? arr[0] : '');
    const pickN = (arr, i) => (arr[i] !== undefined) ? arr[i] : (arr[0] !== undefined ? arr[0] : 0);

    const saleOrderItems = itemSkus.map((sku, i) => {
        const sellingPrice = pickN(sellingPrices, i);
        const totalPrice = pickN(totalPrices, i) || sellingPrice || 0;

        const item = {
            code: pick(itemCodes, i) || `${displayOrderCode}_${i + 1}`,
            itemSku: sku,
            channelProductId: pick(channelProductIds, i) || sku,
            totalPrice: totalPrice,
            sellingPrice: sellingPrice || 0
        };

        if (pick(facilityCodes, i)) item.facilityCode = pick(facilityCodes, i);
        if (pick(shippingMethods, i)) item.shippingMethodCode = pick(shippingMethods, i);
        if (pickN(discounts, i)) item.discount = pickN(discounts, i);
        if (pickN(shippingChargesArr, i)) item.shippingCharges = pickN(shippingChargesArr, i);
        if (pickN(giftWrapChargesArr, i)) item.giftWrapCharges = pickN(giftWrapChargesArr, i);
        if (pickN(prepaidAmounts, i)) item.prepaidAmount = pickN(prepaidAmounts, i);
        if (pickN(packetNumbers, i)) item.packetNumber = pickN(packetNumbers, i);
        if (pick(giftWraps, i) !== undefined) item.giftWrap = pick(giftWraps, i);
        if (pick(giftMessages, i)) item.giftMessage = pick(giftMessages, i);
        if (pick(itemTags, i)) item.itemTag = pick(itemTags, i);

        return item;
    });

    // ==================== CUSTOM FIELDS ====================
    const customFieldValues = [
        { name: 'collected_amount', value: str(row['collected_amount']) },
        { name: 'isFYND', value: str(row['isFYND']) || 'false' },
        { name: 'net_payable_amount', value: str(row['net_payable_amount']) },
        { name: 'order_source', value: str(row['order_source']) || 'POS' },
        { name: 'paymode1_name', value: str(row['paymode1_name']) },
        { name: 'paymode1_id', value: str(row['paymode1_id']) },
        { name: 'paymode1_amount', value: str(row['paymode1_amount']) },
        { name: 'paymode1_type', value: str(row['paymode1_type']) },
        { name: 'promotional_discount_name', value: str(row['promotional_discount_name'] || row['promotional-discount_name']) },
        { name: 'promotional_discount_value', value: str(row['promotional_discount_value'] || row['promotional-discount_value']) },
        { name: 'tsc_share_discount_name', value: str(row['tsc_share_discount_name'] || row['tsc_share-discount_name']) },
        { name: 'tsc_share_discount_value', value: str(row['tsc_share_discount_value'] || row['tsc_share-discount_value']) },
        { name: 'store_code', value: str(row['store_code']) },
        { name: 'total_discount', value: str(row['total_discount']) },
        { name: 'tags', value: str(row['tags']) },
        { name: 'delivery_preference', value: str(row['delivery_preference']) }
    ].filter(cf => cf.value !== '');

    return {
        saleOrder: {
            code: orderCode,
            displayOrderCode: displayOrderCode,
            channel: str(row['Channel']) || 'CUSTOM',
            cashOnDelivery: bool(row['COD*']),
            customerName: str(row['customerName'] || row['Shipping Address Name']),
            customerGSTIN: str(row['Customer GSTIN']),
            currencyCode: str(row['Currency Code']) || 'INR',
            addresses: [shippingAddress, billingAddress],
            billingAddress: { referenceId: billingId },
            shippingAddress: { referenceId: shippingId },
            saleOrderItems,
            customFieldValues,
            totalCashOnDeliveryCharges: 0
        }
    };
}

// ─── POST /api/place-orders ─────────────────────────────────────
app.post('/api/place-orders', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    let rows;
    try {
        rows = parseUpload(req.file.path);
    } catch (e) {
        return res.status(400).json({ error: 'Failed to parse Excel file: ' + e.message });
    } finally {
        fs.unlink(req.file.path, () => { });
    }

    if (!rows.length) return res.status(400).json({ error: 'Excel sheet is empty' });

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();

    const heartbeat = setInterval(() => res.write(': ping\n\n'), 15000);
    res.on('close', () => clearInterval(heartbeat));

    const send = data => res.write(`data: ${JSON.stringify(data)}\n\n`);
    const apiUrl = `${CONFIG.baseUrl.replace(/\/$/, '')}/services/rest/v1/oms/saleOrder/create`;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const orderCode = str(row['Display Sales Order Code']) || str(row['Sales Order Code*']) || `Row ${i + 1}`;

        let payload, success = false, ucCode = '', errors = [];
        try {
            payload = rowToPayload(row);
        } catch (e) {
            errors = [e.message];
            send({ index: i, total: rows.length, orderCode, success: false, ucCode, errors, payload: null });
            if (i < rows.length - 1) await new Promise(r => setTimeout(r, CONFIG.delayMs));
            continue;
        }

        try {
            const apiRes = await fetch(apiUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': CONFIG.token.startsWith('Bearer ') ? CONFIG.token : `Bearer ${CONFIG.token}`
                },
                body: JSON.stringify(payload)
            });

            let data;
            try { data = await apiRes.json(); } catch { data = null; }

            if (data && data.successful) {
                success = true;
                ucCode = data.saleOrderDetailDTO?.code || data.code || '';
            } else if (data) {
                errors = (data.errors || []).map(e => `[${e.fieldName || 'error'}] ${e.description || e.message || ''}`).filter(Boolean);
                if (!errors.length) errors = [data.message || `HTTP ${apiRes.status}`];
            } else {
                errors = [`HTTP ${apiRes.status}`];
            }
        } catch (e) {
            errors = [`Network error: ${e.message}`];
        }

        send({ index: i, total: rows.length, orderCode, success, ucCode, errors, payload });

        if (i < rows.length - 1) await new Promise(r => setTimeout(r, CONFIG.delayMs));
    }

    res.write('data: {"done":true}\n\n');
    clearInterval(heartbeat);
    res.end();
});

// ─── POST /api/parse-preview ──────────────────────────────────────────────────
app.post('/api/parse-preview', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    try {
        const rows = parseUpload(req.file.path);
        res.json({ total: rows.length });
    } catch (e) {
        res.status(400).json({ error: 'Failed to parse: ' + e.message });
    } finally {
        fs.unlink(req.file.path, () => { });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`✅ Server running at http://localhost:${PORT}`));