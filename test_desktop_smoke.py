import os, tempfile, types
from pathlib import Path
import importlib.util

src=Path(__file__).resolve().parent / 'main.py'
spec=importlib.util.spec_from_file_location('medibill',src)
m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)

# Basic utility checks
assert m.parse_date('30-09-2026').isoformat()=='2026-09-30'
assert m.strong_pw('Admin@1234')

# Headless DB schema smoke test
with tempfile.TemporaryDirectory() as td:
    db=m.DB(os.path.join(td,'test.db'))
    # Owner exists and remains unique Admin.
    admins=db.q("SELECT email,role,active FROM users WHERE role='Admin'")
    assert len(admins)==1 and admins[0]['email'].lower()=='admin@medibillwb.in' and admins[0]['active']==1
    # Insert representative bill/customer/medicine data.
    med=db.one("SELECT id,name,price,mrp,stock,gst FROM medicines LIMIT 1")
    now=m.datetime.now(); phone='9876543210'
    cur=db.conn.cursor(); cur.execute("INSERT OR IGNORE INTO customers(name,phone,address,last_purchase,created_by,updated_by) VALUES(?,?,?,?,?,?)",('Test Customer',phone,'',now.isoformat(),'admin@medibillwb.in','admin@medibillwb.in'))
    cid=cur.execute('SELECT id FROM customers WHERE phone=?',(phone,)).fetchone()[0]
    inv='TEST-'+now.strftime('%Y%m%d%H%M%S%f')
    cur.execute("INSERT INTO bills(invoice,bill_date,bill_time,customer_id,customer_name,customer_phone,payment,subtotal,gst,total,discount,created_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",(inv,now.strftime('%Y-%m-%d'),now.strftime('%H:%M:%S'),cid,'Test Customer',phone,'Cash',100,5,105,0,'admin@medibillwb.in'))
    bid=cur.lastrowid
    cur.execute("INSERT INTO bill_items(bill_id,medicine_id,name,molecule,batch,expiry,mrp,qty,price,gst,discount,amount) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",(bid,med['id'],med['name'],'','','',med['mrp'],1,med['price'],med['gst'],0,med['price']*(1+med['gst']/100)))
    db.conn.commit()
    assert db.one('SELECT total FROM bills WHERE invoice=?',(inv,))['total']==105
    assert db.one('SELECT COUNT(*) c FROM bill_items WHERE bill_id=?',(bid,))['c']==1
    # Fake App object for PDF generation.
    app=object.__new__(m.App); app.company='MediBill Pharmacy'; app.sms=types.SimpleNamespace(s={'company_address':'Test Address','company_phone':'9876543210','company_reg':'REG','company_gstin':'GST'})
    app.last_bill=None
    b={'invoice':inv,'date':now.strftime('%Y-%m-%d'),'time':now.strftime('%H:%M:%S'),'name':'Test Customer','phone':phone,'address':'','doctor':'','doctor_reg':'','payment':'Cash','subtotal':100,'gst':5,'total':105,'discount':0,'round_off':0,'items':[{'name':med['name'],'batch':'','qty':1,'mrp':med['mrp'],'price':med['price'],'gst':med['gst'],'discount':0,'amount':105}]}
    outdir=Path(td)/'Bills';outdir.mkdir(); m.PDF_DIR=str(outdir)
    pdf=app.create_pdf(b); assert pdf and Path(pdf).exists() and Path(pdf).stat().st_size>0
    # Linux print path should fall back to opening; monkeypatch to avoid desktop launch.
    app.open_file=lambda p: None
    assert 'PDF' in app.print_pdf(pdf) or 'opened' in app.print_pdf(pdf)
    db.close()

# Verify requested UI wiring exists in source.
s=src.read_text()
for needle in ['Export Excel','self.export_inventory_excel','self.export_customers_excel','self.export_report_excel','def print_pdf','self.print_pdf(pdf_path)']:
    assert needle in s, needle
assert "self.pdf_btn" not in s
assert "self.button(foot,'Print'" not in s
print('v19 headless smoke tests: PASS')

# Export function smoke tests with fake UI dependencies.
from openpyxl import load_workbook
class V:
    def __init__(self,v=''): self.v=v
    def get(self): return self.v

with tempfile.TemporaryDirectory() as td:
    db=m.DB(os.path.join(td,'export.db'))
    # Ensure there is at least one customer and bill.
    med=db.one('SELECT * FROM medicines LIMIT 1')
    db.conn.execute("INSERT OR IGNORE INTO customers(name,phone,total_purchase,last_purchase,created_by,updated_by) VALUES(?,?,?,?,?,?)",('Export Customer','9000000000',123,'2026-09-26','admin','admin'))
    cid=db.one("SELECT id FROM customers WHERE phone='9000000000'")['id']
    inv='EXP-1'; db.conn.execute("INSERT INTO bills(invoice,bill_date,bill_time,customer_id,customer_name,customer_phone,payment,subtotal,gst,total,created_by) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(inv,'2026-09-26','10:00:00',cid,'Export Customer','9000000000','Cash',100,5,105,'admin'))
    bid=db.conn.execute('SELECT last_insert_rowid()').fetchone()[0]
    db.conn.execute("INSERT INTO bill_items(bill_id,medicine_id,name,qty,price,gst,amount) VALUES(?,?,?,?,?,?,?)",(bid,med['id'],med['name'],1,100,5,105));db.conn.commit()
    app=object.__new__(m.App);app.db=db;app.customer_search=V('')
    paths={
        'inv':os.path.join(td,'inv.xlsx'),'cust':os.path.join(td,'cust.xlsx'),'sales':os.path.join(td,'sales.xlsx'),'report':os.path.join(td,'report.xlsx')}
    orig_save=m.filedialog.asksaveasfilename; orig_info=m.messagebox.showinfo; orig_err=m.messagebox.showerror
    try:
        m.messagebox.showinfo=lambda *a,**k:None; m.messagebox.showerror=lambda *a,**k: (_ for _ in ()).throw(AssertionError(a))
        m.filedialog.asksaveasfilename=lambda **k: paths['inv']; app.export_inventory_excel(); assert os.path.exists(paths['inv'])
        m.filedialog.asksaveasfilename=lambda **k: paths['cust']; app.export_customers_excel(); assert os.path.exists(paths['cust'])
        # sales needs filter vars and method db
        app.sf=V('');app.sfrom=V('');app.sto=V('');app.spay=V('All')
        m.filedialog.asksaveasfilename=lambda **k: paths['sales']; app.export_excel(); assert os.path.exists(paths['sales'])
        # report filters
        app.rf=V('');app.rfrom=V('');app.rto=V('');app.rpay=V('All');app.rminprice=V('');app.rmaxprice=V('')
        m.filedialog.asksaveasfilename=lambda **k: paths['report']; app.export_report_excel(); assert os.path.exists(paths['report'])
        assert load_workbook(paths['inv']).sheetnames==['Inventory']
        assert load_workbook(paths['cust']).sheetnames==['Customers']
        assert load_workbook(paths['sales']).sheetnames==['Bills','Bill Items']
        assert load_workbook(paths['report']).sheetnames==['Filtered Bills','Top Medicines']
    finally:
        m.filedialog.asksaveasfilename=orig_save;m.messagebox.showinfo=orig_info;m.messagebox.showerror=orig_err
        db.close()
print('Excel export smoke tests: PASS')

# Full headless checkout test: transaction, customer auto-save, stock deduction,
# bill-item discount persistence, PDF generation + merged print invocation.
with tempfile.TemporaryDirectory() as td:
    db=m.DB(os.path.join(td,'checkout.db'))
    med=db.one('SELECT * FROM medicines WHERE stock>2 LIMIT 1')
    before=int(med['stock'])
    app=object.__new__(m.App);app.db=db;app.user={'email':'admin@medibillwb.in','role':'Admin'};app.company='MediBill Pharmacy'
    app.sms=types.SimpleNamespace(s={'company_address':'','company_phone':'','company_reg':'','company_gstin':''})
    app.b_name=V('Checkout Customer');app.b_phone=V('9123456789');app.b_address=V('Test Address');app.b_doctor=V('');app.b_doctor_reg=V('');app.b_payment=V('UPI')
    app.cart=[{'id':med['id'],'name':med['name'],'molecule':med['molecule'],'batch':med['batch'],'expiry':med['expiry'],'mrp':float(med['mrp']),'manufacturer':'','pack':'','qty':2,'price':float(med['price']),'gst':float(med['gst']),'discount':5.0}]
    app.refresh_cart=lambda:None;app.bill_search=lambda:None;app.sms_btn=types.SimpleNamespace(configure=lambda **k:None)
    m.PDF_DIR=os.path.join(td,'Bills');os.makedirs(m.PDF_DIR)
    printed=[];app.print_pdf=lambda path:(printed.append(path) or 'Printed')
    orig_info=m.messagebox.showinfo;orig_warn=m.messagebox.showwarning;orig_err=m.messagebox.showerror
    try:
        m.messagebox.showinfo=lambda *a,**k:None;m.messagebox.showwarning=lambda *a,**k:None
        def fail(*a,**k): raise AssertionError(a)
        m.messagebox.showerror=fail
        app.checkout()
    finally:
        m.messagebox.showinfo=orig_info;m.messagebox.showwarning=orig_warn;m.messagebox.showerror=orig_err
    bill=db.one('SELECT * FROM bills ORDER BY id DESC LIMIT 1'); assert bill and bill['customer_name']=='Checkout Customer' and bill['payment']=='UPI' and bill['discount']==5.0
    assert db.one('SELECT COUNT(*) c FROM customers WHERE phone=?',('9123456789',))['c']==1
    assert db.one('SELECT stock FROM medicines WHERE id=?',(med['id'],))['stock']==before-2
    item=db.one('SELECT discount FROM bill_items WHERE bill_id=?',(bill['id'],)); assert abs(float(item['discount'])-5.0)<1e-9
    assert printed and Path(printed[0]).exists()
    db.close()
print('Full checkout transaction/print smoke test: PASS')
