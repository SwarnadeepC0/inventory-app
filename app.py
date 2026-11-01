from flask import Flask, render_template, request, jsonify, send_file
import openpyxl
from datetime import datetime
import os

app = Flask(__name__)

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
IGST_RATE = 5
INVENTORY_FILE = os.path.join(DATA_DIR, 'inventory.xlsx')
ORDERS_FILE = os.path.join(DATA_DIR, 'orders.xlsx')
CUSTOMERS_FILE = os.path.join(DATA_DIR, 'customers.xlsx')

os.makedirs(DATA_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

def init_inventory():
    if not os.path.exists(INVENTORY_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Inventory"
        ws.append(['Item Code', 'Item Name', 'Category', 'Unit Price',
                    'Quantity Available', 'Unit', 'Last Updated'])
        samples = [
            ['ITM0001', 'Widget A', 'Widgets', 100, 500, 'pcs'],
            ['ITM0002', 'Widget B', 'Widgets', 250, 300, 'pcs'],
            ['ITM0003', 'Gadget X', 'Gadgets', 500, 150, 'pcs'],
        ]
        now = datetime.now().strftime('%Y-%m-%d')
        for s in samples:
            ws.append(s + [now])
        wb.save(INVENTORY_FILE)


def init_orders():
    if not os.path.exists(ORDERS_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Orders"
        ws.append(['Order ID', 'Date', 'Customer Name', 'Address',
                    'Item Code', 'Item Name', 'Quantity', 'Unit Price',
                    'Amount', 'Status', 'Payment Date', 'Ship Date', 'Notes',
                    'Source', 'Amount Paid'])
        wb.save(ORDERS_FILE)


CUSTOMER_COLS = [
    'Display Name', 'Customer Type', 'Gender', 'First Name', 'Last Name',
    'Company Name', 'Email', 'Phone 1', 'Phone 2',
    'GST Treatment', 'Place of Supply', 'Tax Preference',
    'Address Line', 'City', 'State', 'Country', 'Source', 'Created Date',
]


def init_customers():
    if not os.path.exists(CUSTOMERS_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Customers"
        ws.append(CUSTOMER_COLS)
        wb.save(CUSTOMERS_FILE)


def read_customers():
    init_customers()
    wb = openpyxl.load_workbook(CUSTOMERS_FILE)
    ws = wb.active
    customers = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        customers.append({
            'display_name': str(row[0] or ''),
            'customer_type': str(row[1] or ''),
            'gender': str(row[2] or ''),
            'first_name': str(row[3] or ''),
            'last_name': str(row[4] or ''),
            'company_name': str(row[5] or ''),
            'email': str(row[6] or ''),
            'phone1': str(row[7] or ''),
            'phone2': str(row[8] or ''),
            'gst_treatment': str(row[9] or ''),
            'place_of_supply': str(row[10] or ''),
            'tax_preference': str(row[11] or ''),
            'address_line': str(row[12] or ''),
            'city': str(row[13] or ''),
            'state': str(row[14] or ''),
            'country': str(row[15] or ''),
            'source': str(row[16] or '') if len(row) > 16 else '',
            'created_date': str(row[17] or '') if len(row) > 17 else '',
        })
    wb.close()
    return customers


def save_customer(c):
    wb = openpyxl.load_workbook(CUSTOMERS_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] and str(row[0]) == c['display_name']:
            wb.close()
            return False, 'Display Name already exists'
    ws.append([
        c['display_name'], c.get('customer_type', ''), c.get('gender', ''),
        c.get('first_name', ''), c.get('last_name', ''), c.get('company_name', ''),
        c.get('email', ''), c.get('phone1', ''), c.get('phone2', ''),
        c.get('gst_treatment', ''), c.get('place_of_supply', ''),
        c.get('tax_preference', ''),
        c.get('address_line', ''), c.get('city', ''), c.get('state', ''),
        c.get('country', ''), c.get('source', ''),
        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    ])
    wb.save(CUSTOMERS_FILE)
    wb.close()
    return True, 'Customer created'


def update_customer(display_name, updates):
    col_map = {
        'customer_type': 1, 'gender': 2, 'first_name': 3, 'last_name': 4,
        'company_name': 5, 'email': 6, 'phone1': 7, 'phone2': 8,
        'gst_treatment': 9, 'place_of_supply': 10, 'tax_preference': 11,
        'address_line': 12, 'city': 13, 'state': 14, 'country': 15, 'source': 16,
    }
    wb = openpyxl.load_workbook(CUSTOMERS_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == display_name:
            for key, val in updates.items():
                if key in col_map:
                    row[col_map[key]].value = val
            wb.save(CUSTOMERS_FILE)
            wb.close()
            return True
    wb.close()
    return False


def read_inventory():
    init_inventory()
    wb = openpyxl.load_workbook(INVENTORY_FILE)
    ws = wb.active
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        items.append({
            'item_code': str(row[0]),
            'item_name': str(row[1] or ''),
            'category': str(row[2] or ''),
            'unit_price': float(row[3] or 0),
            'quantity_available': int(row[4] or 0),
            'unit': str(row[5] or 'pcs'),
            'last_updated': str(row[6] or ''),
        })
    wb.close()
    return items


def update_inventory_qty(item_code, delta):
    """Add *delta* to item stock (negative to deduct). Returns (ok, new_qty)."""
    wb = openpyxl.load_workbook(INVENTORY_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == item_code:
            cur = int(row[4].value or 0)
            new = cur + delta
            if new < 0:
                wb.close()
                return False, cur
            row[4].value = new
            row[6].value = datetime.now().strftime('%Y-%m-%d %H:%M')
            wb.save(INVENTORY_FILE)
            wb.close()
            return True, new
    wb.close()
    return False, 0


def read_orders(status_filter=None):
    """Read orders, grouping rows with the same Order ID into one order with multiple items."""
    init_orders()
    wb = openpyxl.load_workbook(ORDERS_FILE)
    ws = wb.active
    orders_map = {}
    order_ids = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        oid = str(row[0])
        item = {
            'item_code': str(row[4] or ''),
            'item_name': str(row[5] or ''),
            'quantity': int(row[6] or 0),
            'unit_price': float(row[7] or 0),
            'amount': float(row[8] or 0),
        }
        if oid in orders_map:
            orders_map[oid]['items'].append(item)
            orders_map[oid]['total_amount'] += item['amount']
        else:
            try:
                amount_paid = float(row[14] or 0) if len(row) > 14 else 0
            except (ValueError, TypeError):
                amount_paid = 0
            o = {
                'order_id': oid,
                'date': str(row[1] or ''),
                'customer_name': str(row[2] or ''),
                'address': str(row[3] or ''),
                'status': str(row[9] or 'Pending'),
                'payment_date': str(row[10] or ''),
                'ship_date': str(row[11] or ''),
                'notes': str(row[12] or ''),
                'source': str(row[13] or '') if len(row) > 13 else '',
                'amount_paid': amount_paid,
                'items': [item],
                'total_amount': item['amount'],
            }
            orders_map[oid] = o
            order_ids.append(oid)
    wb.close()
    result = [orders_map[oid] for oid in order_ids]
    if status_filter:
        result = [o for o in result if o['status'] == status_filter]
    return result


def next_order_id():
    wb = openpyxl.load_workbook(ORDERS_FILE)
    ws = wb.active
    max_id = 0
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] and str(row[0]).startswith('ORD'):
            try:
                max_id = max(max_id, int(str(row[0])[3:]))
            except ValueError:
                pass
    wb.close()
    return f'ORD{max_id + 1:04d}'


def save_order(order):
    """Append rows for a new order (one row per item)."""
    wb = openpyxl.load_workbook(ORDERS_FILE)
    ws = wb.active
    for item in order['items']:
        ws.append([
            order['order_id'], order['date'], order['customer_name'],
            order['address'], item['item_code'], item['item_name'],
            item['quantity'], item['unit_price'], item['amount'],
            order['status'], order.get('payment_date', ''),
            order.get('ship_date', ''), order.get('notes', ''),
            order.get('source', ''), order.get('amount_paid', 0),
        ])
    wb.save(ORDERS_FILE)
    wb.close()


def update_order_fields(order_id, updates: dict):
    """Update order-level fields for all rows belonging to an order."""
    order_col_map = {
        'customer_name': 2, 'address': 3, 'status': 9,
        'payment_date': 10, 'ship_date': 11, 'notes': 12, 'source': 13,
        'amount_paid': 14,
    }
    wb = openpyxl.load_workbook(ORDERS_FILE)
    ws = wb.active
    found = False
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == order_id:
            for key, val in updates.items():
                if key in order_col_map:
                    row[order_col_map[key]].value = val
            found = True
    if found:
        wb.save(ORDERS_FILE)
    wb.close()
    return found


def replace_order(order_id, order):
    """Delete all rows for order_id and write new ones (for full order edits)."""
    wb = openpyxl.load_workbook(ORDERS_FILE)
    ws = wb.active
    rows_to_delete = []
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == order_id:
            rows_to_delete.append(row[0].row)
    for r in reversed(rows_to_delete):
        ws.delete_rows(r)
    for item in order['items']:
        ws.append([
            order['order_id'], order['date'], order['customer_name'],
            order['address'], item['item_code'], item['item_name'],
            item['quantity'], item['unit_price'], item['amount'],
            order['status'], order.get('payment_date', ''),
            order.get('ship_date', ''), order.get('notes', ''),
            order.get('source', ''), order.get('amount_paid', 0),
        ])
    wb.save(ORDERS_FILE)
    wb.close()


# ---------------------------------------------------------------------------
# Routes — Pages
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html')


def amount_to_words(amount):
    """Convert a numeric amount to Indian Rupee words."""
    ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven',
            'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen',
            'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
    tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty',
            'Sixty', 'Seventy', 'Eighty', 'Ninety']

    def _two_digits(n):
        if n < 20:
            return ones[n]
        return (tens[n // 10] + ('-' + ones[n % 10] if n % 10 else '')).strip()

    def _three_digits(n):
        if n >= 100:
            return ones[n // 100] + ' Hundred' + (' ' + _two_digits(n % 100) if n % 100 else '')
        return _two_digits(n)

    if amount == 0:
        return 'Zero'

    rupees = int(amount)
    paise = round((amount - rupees) * 100)

    parts = []
    if rupees >= 10000000:
        parts.append(_two_digits(rupees // 10000000) + ' Crore')
        rupees %= 10000000
    if rupees >= 100000:
        parts.append(_two_digits(rupees // 100000) + ' Lakh')
        rupees %= 100000
    if rupees >= 1000:
        parts.append(_two_digits(rupees // 1000) + ' Thousand')
        rupees %= 1000
    if rupees > 0:
        parts.append(_three_digits(rupees))

    result = 'Indian Rupee ' + ' '.join(parts)
    if paise:
        result += ' and ' + _two_digits(paise) + ' Paise'
    result += ' Only'
    return result


@app.route('/bill/<order_id>')
def bill_page(order_id):
    orders = read_orders()
    order = next((o for o in orders if o['order_id'] == order_id), None)
    if not order:
        return "Order not found", 404
    if order['status'] not in ('Bill Generated', 'Shipped'):
        update_order_fields(order_id, {'status': 'Bill Generated'})
        order['status'] = 'Bill Generated'
    amount_paid = order.get('amount_paid', 0)
    subtotal = order['total_amount']
    igst_amount = round(subtotal * IGST_RATE / 100, 2)
    total = round(subtotal + igst_amount, 2)
    total_in_words = amount_to_words(total)
    payment_made = min(amount_paid, total)
    balance_due = round(total - payment_made, 2)
    return render_template('bill.html', order=order,
                           subtotal=subtotal, igst_rate=IGST_RATE,
                           igst_amount=igst_amount, total=total,
                           total_in_words=total_in_words,
                           payment_made=payment_made,
                           balance_due=balance_due)

# ---------------------------------------------------------------------------
# Routes — API
# ---------------------------------------------------------------------------

@app.route('/api/dashboard')
def api_dashboard():
    orders = read_orders()
    inv = read_inventory()
    statuses = {}
    revenue = 0
    for o in orders:
        statuses[o['status']] = statuses.get(o['status'], 0) + 1
        if o['status'] in ('Payment Received', 'Bill Generated', 'Shipped'):
            revenue += round(o['total_amount'] * (1 + IGST_RATE / 100), 2)
        elif o['status'] == 'Partially Paid':
            revenue += o.get('amount_paid', 0)
    customers = read_customers()
    return jsonify({
        'total_orders': len(orders),
        'order_received': statuses.get('Order Received', 0),
        'order_rejected': statuses.get('Order Rejected', 0),
        'payment_received': statuses.get('Payment Received', 0),
        'bill_generated': statuses.get('Bill Generated', 0),
        'shipped': statuses.get('Shipped', 0),
        'partially_paid': statuses.get('Partially Paid', 0),
        'cancelled': statuses.get('Cancelled', 0),
        'total_revenue': revenue,
        'inventory_items': len(inv),
        'low_stock': len([i for i in inv if i['quantity_available'] < 10]),
        'total_customers': len(customers),
    })


@app.route('/api/inventory')
def api_inventory():
    return jsonify(read_inventory())


@app.route('/api/inventory', methods=['POST'])
def api_save_inventory_item():
    d = request.json
    wb = openpyxl.load_workbook(INVENTORY_FILE)
    ws = wb.active
    found = False
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == d['item_code']:
            row[1].value = d['item_name']
            row[2].value = d.get('category', '')
            row[3].value = float(d.get('unit_price', 0))
            row[4].value = int(d.get('quantity_available', 0))
            row[5].value = d.get('unit', 'pcs')
            row[6].value = datetime.now().strftime('%Y-%m-%d %H:%M')
            found = True
            break
    if not found:
        ws.append([
            d['item_code'], d['item_name'], d.get('category', ''),
            float(d.get('unit_price', 0)), int(d.get('quantity_available', 0)),
            d.get('unit', 'pcs'), datetime.now().strftime('%Y-%m-%d %H:%M'),
        ])
    wb.save(INVENTORY_FILE)
    wb.close()
    return jsonify({'success': True, 'updated': found})


@app.route('/api/inventory/delete', methods=['POST'])
def api_delete_inventory_item():
    code = request.json.get('item_code')
    wb = openpyxl.load_workbook(INVENTORY_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == code:
            ws.delete_rows(row[0].row)
            wb.save(INVENTORY_FILE)
            wb.close()
            return jsonify({'success': True})
    wb.close()
    return jsonify({'success': False, 'error': 'Not found'}), 404


@app.route('/api/inventory/upload', methods=['POST'])
def api_upload_inventory():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file'}), 400
    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': 'Upload .xlsx only'}), 400

    tmp = os.path.join(DATA_DIR, '_upload.xlsx')
    f.save(tmp)
    try:
        wb_up = openpyxl.load_workbook(tmp)
        ws_up = wb_up.active
        headers = [str(c.value or '').strip().lower() for c in ws_up[1]]

        hmap = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            if any(k in h for k in ('code', 'id', 'sku')):
                hmap['code'] = i
            elif any(k in h for k in ('name', 'item', 'product', 'desc')):
                hmap['name'] = i
            elif any(k in h for k in ('cat', 'type', 'group')):
                hmap['cat'] = i
            elif any(k in h for k in ('price', 'rate', 'cost', 'mrp')):
                hmap['price'] = i
            elif any(k in h for k in ('qty', 'quantity', 'stock', 'avail')):
                hmap['qty'] = i
            elif any(k in h for k in ('unit', 'uom')):
                hmap['unit'] = i

        if 'name' not in hmap:
            hmap['name'] = 0

        wb_inv = openpyxl.load_workbook(INVENTORY_FILE)
        ws_inv = wb_inv.active
        existing = {str(r[0].value) for r in ws_inv.iter_rows(min_row=2) if r[0].value}
        count = 0
        auto_id = len(existing)

        def _val(vals, key, default=None):
            idx = hmap.get(key)
            if idx is not None and idx < len(vals):
                return vals[idx]
            return default

        for row in ws_up.iter_rows(min_row=2, values_only=True):
            vals = list(row)
            if not any(vals):
                continue
            name = str(_val(vals, 'name', '') or '').strip()
            if not name:
                continue
            auto_id += 1
            code = str(_val(vals, 'code', '') or '').strip() or f'ITM{auto_id:04d}'
            cat = str(_val(vals, 'cat', '') or '')
            try:
                price = float(_val(vals, 'price', 0) or 0)
            except (ValueError, TypeError):
                price = 0
            try:
                qty = int(float(_val(vals, 'qty', 0) or 0))
            except (ValueError, TypeError):
                qty = 0
            unit = str(_val(vals, 'unit', 'pcs') or 'pcs')
            now = datetime.now().strftime('%Y-%m-%d %H:%M')

            updated = False
            for inv_row in ws_inv.iter_rows(min_row=2):
                if str(inv_row[0].value) == code:
                    inv_row[1].value = name
                    inv_row[2].value = cat
                    inv_row[3].value = price
                    inv_row[4].value = qty
                    inv_row[5].value = unit
                    inv_row[6].value = now
                    updated = True
                    break
            if not updated:
                ws_inv.append([code, name, cat, price, qty, unit, now])
            count += 1

        wb_inv.save(INVENTORY_FILE)
        wb_inv.close()
        wb_up.close()
        return jsonify({'success': True, 'items_imported': count})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


@app.route('/api/inventory/download')
def api_download_inventory():
    return send_file(INVENTORY_FILE, as_attachment=True,
                     download_name='inventory.xlsx')


@app.route('/api/orders')
def api_orders():
    status = request.args.get('status')
    return jsonify(read_orders(status))


@app.route('/api/orders', methods=['POST'])
def api_create_order():
    d = request.json
    items_data = d.get('items', [])
    if not items_data:
        return jsonify({'success': False, 'error': 'No items in order'}), 400

    inv = read_inventory()
    order_items = []
    total_amount = 0

    for item_d in items_data:
        item = next((i for i in inv if i['item_code'] == item_d['item_code']), None)
        if not item:
            return jsonify({'success': False, 'error': f'Item {item_d["item_code"]} not found'}), 400
        qty = int(item_d['quantity'])
        price = item['unit_price']
        amount = float(item_d.get('amount') or (price * qty))
        order_items.append({
            'item_code': item_d['item_code'],
            'item_name': item['item_name'],
            'quantity': qty,
            'unit_price': price,
            'amount': amount,
        })
        total_amount += amount

    # Aggregate qty needed per item_code to check stock
    qty_needed = {}
    for oi in order_items:
        qty_needed[oi['item_code']] = qty_needed.get(oi['item_code'], 0) + oi['quantity']

    rejected = False
    for code, needed in qty_needed.items():
        item = next((i for i in inv if i['item_code'] == code), None)
        if not item or item['quantity_available'] < needed:
            rejected = True
            break

    if not rejected:
        for code, needed in qty_needed.items():
            update_inventory_qty(code, -needed)
        status = 'Order Received'
    else:
        status = 'Order Rejected'

    oid = next_order_id()
    order = {
        'order_id': oid,
        'date': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'customer_name': d['customer_name'],
        'address': d['address'],
        'items': order_items,
        'total_amount': total_amount,
        'status': status,
        'notes': d.get('notes', ''),
        'source': d.get('source', ''),
    }
    save_order(order)

    msg = f'Order {oid}: {status}'
    if status == 'Order Rejected':
        msg += ' (insufficient stock for one or more items)'
    return jsonify({'success': True, 'order': order, 'message': msg})


@app.route('/api/orders/<order_id>/status', methods=['PUT'])
def api_update_status(order_id):
    d = request.json
    new_status = d['status']
    orders = read_orders()
    order = next((o for o in orders if o['order_id'] == order_id), None)
    if not order:
        return jsonify({'success': False, 'error': 'Not found'}), 404

    old = order['status']
    updates = {'status': new_status}

    # Restore stock on cancel
    if new_status == 'Cancelled' and old in ('Order Received', 'Payment Received', 'Bill Generated', 'Partially Paid'):
        for item in order['items']:
            update_inventory_qty(item['item_code'], item['quantity'])

    # Retry rejected order
    if new_status == 'Order Received' and old == 'Order Rejected':
        inv = read_inventory()
        qty_needed = {}
        for item in order['items']:
            qty_needed[item['item_code']] = qty_needed.get(item['item_code'], 0) + item['quantity']
        for code, needed in qty_needed.items():
            i = next((x for x in inv if x['item_code'] == code), None)
            avail = i['quantity_available'] if i else 0
            if avail < needed:
                return jsonify({
                    'success': False,
                    'error': f'Insufficient stock for {code} ({avail} available, {needed} needed)',
                }), 400
        for code, needed in qty_needed.items():
            update_inventory_qty(code, -needed)

    if new_status == 'Payment Received':
        updates['payment_date'] = datetime.now().strftime('%Y-%m-%d %H:%M')
        updates['amount_paid'] = round(order['total_amount'] * (1 + IGST_RATE / 100), 2)
    if new_status == 'Shipped':
        updates['ship_date'] = datetime.now().strftime('%Y-%m-%d %H:%M')

    update_order_fields(order_id, updates)
    return jsonify({'success': True, 'message': f'{order_id} -> {new_status}'})


@app.route('/api/orders/<order_id>/partial-pay', methods=['POST'])
def api_partial_pay(order_id):
    d = request.json
    pay_amount = float(d.get('amount', 0))
    if pay_amount <= 0:
        return jsonify({'success': False, 'error': 'Amount must be greater than 0'}), 400

    orders = read_orders()
    order = next((o for o in orders if o['order_id'] == order_id), None)
    if not order:
        return jsonify({'success': False, 'error': 'Not found'}), 404

    if order['status'] not in ('Order Received', 'Partially Paid', 'Bill Generated'):
        return jsonify({'success': False, 'error': f'Cannot accept payment for status: {order["status"]}'}), 400

    subtotal = order['total_amount']
    grand_total = round(subtotal * (1 + IGST_RATE / 100), 2)
    already_paid = order.get('amount_paid', 0)
    new_paid = round(already_paid + pay_amount, 2)
    balance = round(grand_total - new_paid, 2)

    if new_paid > grand_total:
        return jsonify({'success': False, 'error': f'Payment exceeds balance due. Balance: {grand_total - already_paid:.2f}'}), 400

    updates = {
        'amount_paid': new_paid,
        'payment_date': datetime.now().strftime('%Y-%m-%d %H:%M'),
    }
    if balance <= 0:
        if order['status'] == 'Bill Generated':
            pass  # keep Bill Generated status
        else:
            updates['status'] = 'Payment Received'
    else:
        if order['status'] != 'Bill Generated':
            updates['status'] = 'Partially Paid'

    update_order_fields(order_id, updates)
    return jsonify({
        'success': True,
        'message': f'{order_id}: Received {pay_amount:.2f}, Total paid: {new_paid:.2f}, Balance: {balance:.2f}',
        'amount_paid': new_paid,
        'balance': balance,
    })


@app.route('/api/orders/<order_id>', methods=['PUT'])
def api_update_order(order_id):
    d = request.json
    orders = read_orders()
    old_order = next((o for o in orders if o['order_id'] == order_id), None)
    if not old_order:
        return jsonify({'success': False, 'error': 'Not found'}), 404

    new_items = d.get('items')
    if new_items is not None:
        inv = read_inventory()

        # If order was accepted, restore old inventory first
        if old_order['status'] in ('Order Received', 'Payment Received', 'Bill Generated', 'Partially Paid', 'Shipped'):
            for item in old_order['items']:
                update_inventory_qty(item['item_code'], item['quantity'])

        # Build new items list
        processed_items = []
        total = 0
        for ni in new_items:
            inv_item = next((i for i in inv if i['item_code'] == ni['item_code']), None)
            qty = int(ni['quantity'])
            price = inv_item['unit_price'] if inv_item else float(ni.get('unit_price', 0))
            amt = float(ni.get('amount') or (price * qty))
            processed_items.append({
                'item_code': ni['item_code'],
                'item_name': inv_item['item_name'] if inv_item else ni.get('item_name', ''),
                'quantity': qty,
                'unit_price': price,
                'amount': amt,
            })
            total += amt

        # Check stock for new items if order was accepted
        if old_order['status'] in ('Order Received', 'Payment Received', 'Bill Generated', 'Partially Paid', 'Shipped'):
            inv = read_inventory()  # re-read after restoring
            qty_needed = {}
            for pi in processed_items:
                qty_needed[pi['item_code']] = qty_needed.get(pi['item_code'], 0) + pi['quantity']
            for code, needed in qty_needed.items():
                i = next((x for x in inv if x['item_code'] == code), None)
                if not i or i['quantity_available'] < needed:
                    # Rollback
                    for item in old_order['items']:
                        update_inventory_qty(item['item_code'], -item['quantity'])
                    return jsonify({'success': False, 'error': f'Insufficient stock for {code}'}), 400
            for code, needed in qty_needed.items():
                update_inventory_qty(code, -needed)

        new_amount_paid = float(d['amount_paid']) if 'amount_paid' in d else old_order.get('amount_paid', 0)
        grand_total = round(total * (1 + IGST_RATE / 100), 2)

        # Determine status based on payment
        new_status = old_order['status']
        if old_order['status'] not in ('Bill Generated', 'Shipped'):
            if new_amount_paid >= grand_total:
                new_status = 'Payment Received'
            elif new_amount_paid > 0:
                new_status = 'Partially Paid'
            elif old_order['status'] in ('Payment Received', 'Partially Paid'):
                new_status = 'Order Received'

        payment_date = old_order['payment_date']
        if new_amount_paid > 0 and (not payment_date or payment_date == 'None'):
            payment_date = datetime.now().strftime('%Y-%m-%d %H:%M')

        updated_order = {
            'order_id': order_id,
            'date': old_order['date'],
            'customer_name': d.get('customer_name', old_order['customer_name']),
            'address': d.get('address', old_order['address']),
            'items': processed_items,
            'status': new_status,
            'payment_date': payment_date,
            'ship_date': old_order['ship_date'],
            'notes': d.get('notes', old_order['notes']),
            'source': old_order.get('source', ''),
            'amount_paid': new_amount_paid,
        }
        replace_order(order_id, updated_order)
        return jsonify({'success': True})
    else:
        update_order_fields(order_id, d)
        return jsonify({'success': True})


@app.route('/api/orders/<order_id>/delete', methods=['POST'])
def api_delete_order(order_id):
    orders = read_orders()
    order = next((o for o in orders if o['order_id'] == order_id), None)
    if not order:
        return jsonify({'success': False, 'error': 'Not found'}), 404
    # Restore stock if order was accepted
    if order['status'] in ('Order Received', 'Payment Received', 'Bill Generated', 'Partially Paid'):
        for item in order['items']:
            update_inventory_qty(item['item_code'], item['quantity'])
    # Delete all rows for this order
    wb = openpyxl.load_workbook(ORDERS_FILE)
    ws = wb.active
    rows_to_del = []
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == order_id:
            rows_to_del.append(row[0].row)
    for r in reversed(rows_to_del):
        ws.delete_rows(r)
    wb.save(ORDERS_FILE)
    wb.close()
    return jsonify({'success': True})


@app.route('/api/orders/download')
def api_download_orders():
    return send_file(ORDERS_FILE, as_attachment=True,
                     download_name='orders.xlsx')


# ---------------------------------------------------------------------------
# Routes — Customers API
# ---------------------------------------------------------------------------

@app.route('/api/customers')
def api_customers():
    return jsonify(read_customers())


@app.route('/api/customers', methods=['POST'])
def api_create_customer():
    d = request.json
    if not d.get('display_name'):
        return jsonify({'success': False, 'error': 'Display Name is required'}), 400
    updated = update_customer(d['display_name'], d)
    if updated:
        return jsonify({'success': True, 'updated': True, 'message': 'Customer updated'})
    ok, msg = save_customer(d)
    if not ok:
        return jsonify({'success': False, 'error': msg}), 400
    return jsonify({'success': True, 'updated': False, 'message': msg})


@app.route('/api/customers/<display_name>', methods=['PUT'])
def api_update_customer(display_name):
    d = request.json
    ok = update_customer(display_name, d)
    if not ok:
        return jsonify({'success': False, 'error': 'Customer not found'}), 404
    return jsonify({'success': True})


@app.route('/api/customers/delete', methods=['POST'])
def api_delete_customer():
    name = request.json.get('display_name')
    wb = openpyxl.load_workbook(CUSTOMERS_FILE)
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == name:
            ws.delete_rows(row[0].row)
            wb.save(CUSTOMERS_FILE)
            wb.close()
            return jsonify({'success': True})
    wb.close()
    return jsonify({'success': False, 'error': 'Not found'}), 404


@app.route('/api/customers/download')
def api_download_customers():
    return send_file(CUSTOMERS_FILE, as_attachment=True,
                     download_name='customers.xlsx')


# ---------------------------------------------------------------------------
if __name__ == '__main__':
    init_inventory()
    init_orders()
    init_customers()
    app.run(host='0.0.0.0', port=5000, debug=True)
