import os, re, json, sqlite3, hashlib, secrets, urllib.request, urllib.parse
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    Workbook = load_workbook = None

APP_DIR = os.path.join(os.path.expanduser('~'), 'MediBillPro')
os.makedirs(APP_DIR, exist_ok=True)
DB_PATH = os.path.join(APP_DIR, 'medibill.db')
SETTINGS_PATH = os.path.join(APP_DIR, 'settings.json')

ROLES = ['Admin', 'Store Manager', 'Cashier', 'Accounts']
PERMS = {
    'Admin': {'billing','inventory','customers','sales','reports','settings','users','purchases'},
    'Store Manager': {'billing','inventory','customers','sales','reports','settings','users','purchases'},
    'Cashier': {'billing','customers'},
    'Accounts': {'customers','sales','reports'},
}

def money(v):
    return float(Decimal(str(v or 0)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))

def hash_pw(pw, salt=None):
    salt = salt or secrets.token_hex(16)
    return salt, hashlib.pbkdf2_hmac('sha256', pw.encode(), salt.encode(), 120000).hex()

def verify_pw(pw, salt, digest):
    return hash_pw(pw, salt)[1] == digest

def valid_email(s):
    return bool(re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', s.strip()))

def strong_pw(s):
    return len(s)>=8 and bool(re.search(r'[A-Z]',s)) and bool(re.search(r'[a-z]',s)) and bool(re.search(r'\d',s)) and bool(re.search(r'[^A-Za-z0-9]',s))

def normalize_phone(p):
    return re.sub(r'\D','',p or '')

class DB:
    def __init__(self, path=DB_PATH):
        self.conn=sqlite3.connect(path)
        self.conn.row_factory=sqlite3.Row
        self.conn.execute('PRAGMA foreign_keys=ON')
        self.init()
    def init(self):
        c=self.conn.cursor()
        c.executescript('''
        CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, name TEXT NOT NULL, email TEXT UNIQUE NOT NULL, salt TEXT NOT NULL, password_hash TEXT NOT NULL, role TEXT NOT NULL, verified INTEGER DEFAULT 1, active INTEGER DEFAULT 1, created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS customers(id INTEGER PRIMARY KEY, name TEXT NOT NULL, phone TEXT UNIQUE NOT NULL, email TEXT DEFAULT '', address TEXT DEFAULT '', credit REAL DEFAULT 0, total_purchase REAL DEFAULT 0, last_purchase TEXT);
        CREATE TABLE IF NOT EXISTS medicines(id INTEGER PRIMARY KEY, name TEXT NOT NULL, batch TEXT DEFAULT '', expiry TEXT DEFAULT '', price REAL DEFAULT 0, cost REAL DEFAULT 0, stock INTEGER DEFAULT 0, gst REAL DEFAULT 0, supplier TEXT DEFAULT '', UNIQUE(name,batch));
        CREATE TABLE IF NOT EXISTS bills(id INTEGER PRIMARY KEY, invoice TEXT UNIQUE NOT NULL, bill_date TEXT NOT NULL, bill_time TEXT NOT NULL, customer_id INTEGER, customer_name TEXT, customer_phone TEXT, payment TEXT, subtotal REAL, gst REAL, total REAL, sms_status TEXT DEFAULT 'Not sent', created_by TEXT, FOREIGN KEY(customer_id) REFERENCES customers(id));
        CREATE TABLE IF NOT EXISTS bill_items(id INTEGER PRIMARY KEY, bill_id INTEGER NOT NULL, medicine_id INTEGER, name TEXT, batch TEXT, qty INTEGER, price REAL, gst REAL, amount REAL, FOREIGN KEY(bill_id) REFERENCES bills(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS sms_outbox(id INTEGER PRIMARY KEY, bill_id INTEGER, phone TEXT, message TEXT, status TEXT, provider_response TEXT DEFAULT '', created_at TEXT, FOREIGN KEY(bill_id) REFERENCES bills(id));
        ''')
        if not c.execute('SELECT 1 FROM users LIMIT 1').fetchone():
            for name,email,pw,role in [('Administrator','admin@medibillwb.in','Admin@1234','Admin'),('Store Manager','store@medibillwb.in','Demo@1234','Store Manager'),('Cashier','cashier@medibillwb.in','Demo@1234','Cashier'),('Accounts','accounts@medibillwb.in','Demo@1234','Accounts')]:
                salt,d=hash_pw(pw); c.execute('INSERT INTO users(name,email,salt,password_hash,role,created_at) VALUES(?,?,?,?,?,?)',(name,email,salt,d,role,datetime.now().isoformat(timespec='seconds')))
        if not c.execute('SELECT 1 FROM medicines LIMIT 1').fetchone():
            meds=[('Paracetamol 500mg','PCT001','2027-08-31',25,12,80,5,'Demo Supplier'),('Azithromycin 500mg','AZI001','2027-03-31',85,48,25,5,'Demo Supplier'),('ORS Lemon','ORS001','2028-01-31',22,13,50,5,'Demo Supplier'),('Pantoprazole 40mg','PAN001','2027-11-30',60,30,35,12,'Demo Supplier')]
            c.executemany('INSERT INTO medicines(name,batch,expiry,price,cost,stock,gst,supplier) VALUES(?,?,?,?,?,?,?,?)',meds)
        self.conn.commit()
    def q(self,sql,args=()): return self.conn.execute(sql,args).fetchall()
    def one(self,sql,args=()): return self.conn.execute(sql,args).fetchone()
    def close(self): self.conn.close()

class SMS:
    def __init__(self):
        self.s={}
        if os.path.exists(SETTINGS_PATH):
            try: self.s=json.load(open(SETTINGS_PATH,encoding='utf8'))
            except: pass
    def save(self): json.dump(self.s,open(SETTINGS_PATH,'w',encoding='utf8'),indent=2)
    def send(self, phone, message):
        mode=self.s.get('sms_mode','None')
        if mode=='None': return False,'SMS provider is not configured. Bill saved to SMS Outbox.'
        if mode=='HTTP JSON':
            url=self.s.get('sms_url','').strip(); token=self.s.get('sms_token','').strip()
            if not url: return False,'SMS URL is missing.'
            payload={'to':phone,'message':message}
            req=urllib.request.Request(url,data=json.dumps(payload).encode(),headers={'Content-Type':'application/json','Authorization':('Bearer '+token) if token else ''},method='POST')
            try:
                with urllib.request.urlopen(req,timeout=15) as r: return 200 <= r.status < 300, r.read().decode(errors='ignore')[:500]
            except Exception as e: return False,str(e)
        return False,'Unsupported SMS mode.'

class App:
    def __init__(self):
        self.db=DB(); self.sms=SMS(); self.user=None; self.cart=[]
        self.root=tk.Tk(); self.root.withdraw(); self.root.title('MediBill Pro'); self.root.geometry('1200x760'); self.root.minsize(1050,680)
        self.login()
        self.root.protocol('WM_DELETE_WINDOW', self.close)
    def close(self): self.db.close(); self.root.destroy()
    def clear(self):
        for w in self.root.winfo_children(): w.destroy()
    def login(self):
        self.clear(); self.root.deiconify(); self.root.geometry('560x500');
        f=ttk.Frame(self.root,padding=35); f.pack(fill='both',expand=True)
        ttk.Label(f,text='MediBill Pro',font=('Segoe UI',28,'bold')).pack(pady=(35,5)); ttk.Label(f,text='Pharmacy Billing & Inventory',font=('Segoe UI',11)).pack(pady=(0,25))
        form=ttk.Frame(f); form.pack()
        ttk.Label(form,text='Email').grid(row=0,column=0,sticky='w',pady=8); email=ttk.Entry(form,width=36); email.grid(row=0,column=1,pady=8)
        ttk.Label(form,text='Password').grid(row=1,column=0,sticky='w',pady=8); pw=ttk.Entry(form,width=36,show='•'); pw.grid(row=1,column=1,pady=8)
        ttk.Label(f,text='Demo: admin@medibillwb.in / Admin@1234',foreground='#666').pack(pady=12)
        def go():
            e=email.get().strip().lower(); row=self.db.one('SELECT * FROM users WHERE lower(email)=? AND active=1',(e,))
            if not row or not verify_pw(pw.get(),row['salt'],row['password_hash']): messagebox.showerror('Login failed','Invalid email or password.'); return
            self.user=dict(row); self.dashboard()
        ttk.Button(f,text='Login',command=go).pack(ipadx=40,ipady=7,pady=12)
        ttk.Button(f,text='Create worker account',command=self.register).pack()
    def register(self):
        win=tk.Toplevel(self.root); win.title('Create Worker Account'); win.geometry('500x470'); win.transient(self.root); win.grab_set()
        frm=ttk.Frame(win,padding=25); frm.pack(fill='both',expand=True)
        fields=[]
        for i,label in enumerate(['Full name','Email','Password','Confirm password']):
            ttk.Label(frm,text=label).grid(row=i,column=0,sticky='w',pady=8); e=ttk.Entry(frm,width=32,show='•' if 'password' in label.lower() else ''); e.grid(row=i,column=1,pady=8); fields.append(e)
        ttk.Label(frm,text='Role is assigned by Store Manager/Admin after creation.',wraplength=400).grid(row=4,column=0,columnspan=2,pady=15)
        def save():
            name,email,p1,p2=[x.get().strip() for x in fields]
            if not name or not valid_email(email): messagebox.showerror('Validation','Enter a valid name and email.'); return
            if not strong_pw(p1): messagebox.showerror('Validation','Password needs 8+ chars with upper, lower, number and special character.'); return
            if p1!=p2: messagebox.showerror('Validation','Passwords do not match.'); return
            if self.db.one('SELECT 1 FROM users WHERE lower(email)=?',(email.lower(),)): messagebox.showerror('Validation','This email already has an account.'); return
            salt,d=hash_pw(p1); self.db.conn.execute('INSERT INTO users(name,email,salt,password_hash,role,verified,created_at) VALUES(?,?,?,?,?,?,?)',(name,email,salt,d,'Cashier',1,datetime.now().isoformat(timespec='seconds'))); self.db.conn.commit(); win.destroy(); messagebox.showinfo('Created','Account created. A Store Manager/Admin can change its role.');
        ttk.Button(frm,text='Create Account',command=save).grid(row=5,column=0,columnspan=2,pady=10)
    def dashboard(self):
        self.clear(); self.root.geometry('1200x760');
        outer=ttk.Frame(self.root); outer.pack(fill='both',expand=True)
        side=ttk.Frame(outer,padding=12); side.pack(side='left',fill='y'); content=ttk.Frame(outer,padding=18); content.pack(side='left',fill='both',expand=True)
        self.content=content
        ttk.Label(side,text='MediBill',font=('Segoe UI',20,'bold')).pack(pady=(5,2)); ttk.Label(side,text=self.user['role']).pack(pady=(0,20))
        buttons=[('Dashboard','dashboard_home'),('Billing','billing'),('Inventory','inventory'),('Customers','customers'),('Sales / Excel','sales'),('Reports','reports'),('Users / Roles','users'),('SMS Outbox','sms_outbox'),('Settings','settings')]
        for label,fn in buttons:
            perm={'dashboard_home':'billing','billing':'billing','inventory':'inventory','customers':'customers','sales':'sales','reports':'reports','users':'users','sms_outbox':'billing','settings':'settings'}[fn]
            if perm in PERMS[self.user['role']]: ttk.Button(side,text=label,width=20,command=getattr(self,fn)).pack(fill='x',pady=3)
        ttk.Separator(side).pack(fill='x',pady=15); ttk.Button(side,text='Logout',command=self.login,width=20).pack(fill='x')
        self.dashboard_home()
    def title(self,t,sub=''):
        for w in self.content.winfo_children(): w.destroy()
        ttk.Label(self.content,text=t,font=('Segoe UI',24,'bold')).pack(anchor='w');
        if sub: ttk.Label(self.content,text=sub,foreground='#666').pack(anchor='w',pady=(0,15))
    def dashboard_home(self):
        self.title('Dashboard','Current local database status')
        stats=[('Medicines',self.db.one('SELECT COUNT(*) n FROM medicines')['n']),('Customers',self.db.one('SELECT COUNT(*) n FROM customers')['n']),('Bills',self.db.one('SELECT COUNT(*) n FROM bills')['n']),('Today Sales',money(self.db.one("SELECT COALESCE(SUM(total),0) n FROM bills WHERE bill_date=?",(datetime.now().strftime('%Y-%m-%d'),))['n']))]
        row=ttk.Frame(self.content); row.pack(fill='x',pady=10)
        for a,b in stats:
            x=ttk.LabelFrame(row,text=a,padding=20); x.pack(side='left',fill='x',expand=True,padx=5); ttk.Label(x,text=str(b),font=('Segoe UI',22,'bold')).pack()
        ttk.Label(self.content,text='Low stock / near expiry',font=('Segoe UI',14,'bold')).pack(anchor='w',pady=(25,8))
        tv=ttk.Treeview(self.content,columns=('name','stock','expiry'),show='headings',height=8); tv.pack(fill='both',expand=True)
        for c,h in [('name','Medicine'),('stock','Stock'),('expiry','Expiry')]: tv.heading(c,text=h)
        for r in self.db.q('SELECT name,stock,expiry FROM medicines WHERE stock<=10 OR expiry<=date(\'now\',\'+90 day\') ORDER BY expiry LIMIT 30'): tv.insert('', 'end',values=(r['name'],r['stock'],r['expiry']))
    def billing(self):
        self.title('Billing','Generate a bill; customer details are saved automatically.')
        top=ttk.Frame(self.content); top.pack(fill='x');
        self.b_name=tk.StringVar(); self.b_phone=tk.StringVar(); self.b_payment=tk.StringVar(value='Cash'); self.b_search=tk.StringVar();
        for i,(lab,var) in enumerate([('Customer name',self.b_name),('Mobile',self.b_phone)]): ttk.Label(top,text=lab).grid(row=0,column=i*2,sticky='w',padx=5); ttk.Entry(top,textvariable=var,width=25).grid(row=0,column=i*2+1,padx=5)
        ttk.Label(top,text='Payment').grid(row=0,column=4); ttk.Combobox(top,textvariable=self.b_payment,values=['Cash','UPI','Card','Credit'],state='readonly',width=12).grid(row=0,column=5,padx=5)
        ttk.Label(self.content,text='Add medicine',font=('Segoe UI',12,'bold')).pack(anchor='w',pady=(18,5)); sf=ttk.Frame(self.content); sf.pack(fill='x'); ttk.Entry(sf,textvariable=self.b_search,width=45).pack(side='left'); ttk.Button(sf,text='Search',command=self.bill_search).pack(side='left',padx=5); ttk.Button(sf,text='Clear cart',command=lambda:(self.cart.clear(),self.refresh_cart())).pack(side='left')
        self.bill_results=ttk.Treeview(self.content,columns=('id','name','batch','price','stock'),show='headings',height=5); self.bill_results.pack(fill='x',pady=8)
        for c,h in [('id','ID'),('name','Medicine'),('batch','Batch'),('price','Price'),('stock','Stock')]: self.bill_results.heading(c,text=h)
        self.bill_results.bind('<Double-1>',lambda e:self.add_selected())
        ttk.Button(self.content,text='Add selected (double-click also works)',command=self.add_selected).pack(anchor='w')
        self.cart_tv=ttk.Treeview(self.content,columns=('name','qty','price','gst','amount'),show='headings',height=8); self.cart_tv.pack(fill='both',expand=True,pady=10)
        for c,h in [('name','Medicine'),('qty','Qty'),('price','Price'),('gst','GST %'),('amount','Amount')]: self.cart_tv.heading(c,text=h)
        self.total_lbl=ttk.Label(self.content,text='Total: ₹0.00',font=('Segoe UI',18,'bold')); self.total_lbl.pack(anchor='e',pady=5); ttk.Button(self.content,text='GENERATE BILL + SAVE CUSTOMER + SEND SMS',command=self.checkout).pack(anchor='e',ipadx=12,ipady=8)
        self.bill_search()
    def bill_search(self):
        q='%'+self.b_search.get().strip()+'%';
        for x in self.bill_results.get_children(): self.bill_results.delete(x)
        for r in self.db.q('SELECT id,name,batch,price,stock FROM medicines WHERE name LIKE ? OR batch LIKE ? ORDER BY name LIMIT 50',(q,q)): self.bill_results.insert('', 'end',values=tuple(r))
    def add_selected(self):
        s=self.bill_results.selection()
        if not s: return
        rid=int(self.bill_results.item(s[0])['values'][0]); r=self.db.one('SELECT * FROM medicines WHERE id=?',(rid,));
        if r['stock']<=0: messagebox.showwarning('Stock','Out of stock.'); return
        for item in self.cart:
            if item['id']==rid:
                if item['qty']<r['stock']: item['qty']+=1
                self.refresh_cart(); return
        self.cart.append({'id':r['id'],'name':r['name'],'batch':r['batch'],'qty':1,'price':r['price'],'gst':r['gst']}); self.refresh_cart()
    def refresh_cart(self):
        if not hasattr(self,'cart_tv'): return
        for x in self.cart_tv.get_children(): self.cart_tv.delete(x)
        total=0
        for i,it in enumerate(self.cart):
            amt=money(it['qty']*it['price']*(1+it['gst']/100)); total+=amt; self.cart_tv.insert('', 'end',iid=str(i),values=(it['name'],it['qty'],f"₹{it['price']:.2f}",it['gst'],f"₹{amt:.2f}"))
        self.total_lbl.config(text=f'Total: ₹{total:.2f}')
    def checkout(self):
        if not self.cart: messagebox.showwarning('Bill','Add at least one medicine.'); return
        name=self.b_name.get().strip() or 'Walk-in Customer'; phone=normalize_phone(self.b_phone.get())
        if phone and len(phone)<10: messagebox.showerror('Mobile','Enter a valid mobile number.'); return
        now=datetime.now(); invoice='INV-'+now.strftime('%Y%m%d-%H%M%S')+'-'+secrets.token_hex(2).upper(); subtotal=gsttotal=total=0
        try:
            cur=self.db.conn.cursor(); cur.execute('BEGIN')
            cid=None
            if phone:
                old=cur.execute('SELECT id FROM customers WHERE phone=?',(phone,)).fetchone()
                if old: cid=old['id']; cur.execute('UPDATE customers SET name=?,last_purchase=? WHERE id=?',(name,now.isoformat(timespec='seconds'),cid))
                else: cur.execute('INSERT INTO customers(name,phone,last_purchase) VALUES(?,?,?)',(name,phone,now.isoformat(timespec='seconds'))); cid=cur.lastrowid
            for it in self.cart:
                stock=cur.execute('SELECT stock FROM medicines WHERE id=?',(it['id'],)).fetchone()['stock']
                if stock<it['qty']: raise ValueError(f"Insufficient stock for {it['name']}")
                base=Decimal(str(it['qty']*it['price'])); tax=(base*Decimal(str(it['gst']))/Decimal('100')).quantize(Decimal('0.01')); amt=base+tax; subtotal+=float(base); gsttotal+=float(tax); total+=float(amt)
            cur.execute('INSERT INTO bills(invoice,bill_date,bill_time,customer_id,customer_name,customer_phone,payment,subtotal,gst,total,sms_status,created_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',(invoice,now.strftime('%Y-%m-%d'),now.strftime('%H:%M:%S'),cid,name,phone,self.b_payment.get(),money(subtotal),money(gsttotal),money(total),'Pending' if phone else 'No mobile',self.user['email']))
            bid=cur.lastrowid
            for it in self.cart:
                base=money(it['qty']*it['price']); tax=money(base*it['gst']/100); cur.execute('INSERT INTO bill_items(bill_id,medicine_id,name,batch,qty,price,gst,amount) VALUES(?,?,?,?,?,?,?,?)',(bid,it['id'],it['name'],it['batch'],it['qty'],it['price'],it['gst'],money(base+tax))); cur.execute('UPDATE medicines SET stock=stock-? WHERE id=?',(it['qty'],it['id']))
            if cid: cur.execute('UPDATE customers SET total_purchase=total_purchase+? WHERE id=?',(money(total),cid))
            cur.execute('INSERT INTO sms_outbox(bill_id,phone,message,status,created_at) VALUES(?,?,?,?,?)',(bid,phone,self.sms_message(invoice,name,total), 'Queued' if phone else 'No mobile', now.isoformat(timespec='seconds')))
            self.db.conn.commit()
        except Exception as e:
            self.db.conn.rollback(); messagebox.showerror('Could not generate bill',str(e)); return
        sms_status='Not applicable'
        if phone:
            ok,resp=self.sms.send(phone,self.sms_message(invoice,name,total)); sms_status='Sent' if ok else 'Queued'
            self.db.conn.execute('UPDATE bills SET sms_status=? WHERE id=?',(sms_status,bid)); self.db.conn.execute('UPDATE sms_outbox SET status=?,provider_response=? WHERE bill_id=?',(sms_status,resp,bid)); self.db.conn.commit()
        self.cart.clear(); self.b_name.set(''); self.b_phone.set(''); self.refresh_cart(); self.bill_search()
        messagebox.showinfo('Bill generated',f'Invoice: {invoice}\nTotal: ₹{total:.2f}\nCustomer saved: {"Yes" if phone else "No mobile supplied"}\nSMS: {sms_status}')
    def sms_message(self,invoice,name,total): return f'MediBill Pro: Dear {name}, your bill {invoice} total is Rs.{total:.2f}. Thank you.'
    def inventory(self):
        self.title('Inventory','Add and manage medicines, stock, batches and expiry.')
        form=ttk.Frame(self.content); form.pack(fill='x',pady=5); vars=[tk.StringVar() for _ in range(8)]; labels=['Name','Batch','Expiry YYYY-MM-DD','Selling price','Cost','Stock','GST %','Supplier']
        for i,(lab,var) in enumerate(zip(labels,vars)): ttk.Label(form,text=lab).grid(row=i//4*2,column=i%4,sticky='w',padx=4); ttk.Entry(form,textvariable=var,width=20).grid(row=i//4*2+1,column=i%4,padx=4,pady=3)
        def add():
            try: self.db.conn.execute('INSERT INTO medicines(name,batch,expiry,price,cost,stock,gst,supplier) VALUES(?,?,?,?,?,?,?,?)',tuple(v.get().strip() if i<3 or i==7 else float(v.get() or 0) if i in [3,4,6] else int(v.get() or 0) for i,v in enumerate(vars))); self.db.conn.commit(); refresh(); [v.set('') for v in vars]
            except Exception as e: messagebox.showerror('Inventory',str(e))
        ttk.Button(form,text='Add medicine',command=add).grid(row=4,column=0,pady=8)
        tv=ttk.Treeview(self.content,columns=('id','name','batch','expiry','price','stock','gst','supplier'),show='headings'); tv.pack(fill='both',expand=True,pady=10)
        for c,h in [('id','ID'),('name','Name'),('batch','Batch'),('expiry','Expiry'),('price','Price'),('stock','Stock'),('gst','GST'),('supplier','Supplier')]: tv.heading(c,text=h)
        def refresh():
            for x in tv.get_children(): tv.delete(x)
            for r in self.db.q('SELECT id,name,batch,expiry,price,stock,gst,supplier FROM medicines ORDER BY name'): tv.insert('', 'end',values=tuple(r))
        refresh()
    def customers(self):
        self.title('Customers','Customer records saved automatically from completed bills.')
        tv=ttk.Treeview(self.content,columns=('name','phone','email','credit','total','last'),show='headings'); tv.pack(fill='both',expand=True)
        for c,h in [('name','Name'),('phone','Mobile'),('email','Email'),('credit','Credit'),('total','Total purchase'),('last','Last purchase')]: tv.heading(c,text=h)
        for r in self.db.q('SELECT name,phone,email,credit,total_purchase,last_purchase FROM customers ORDER BY last_purchase DESC'): tv.insert('', 'end',values=tuple(r))
    def sales(self):
        self.title('Sales / Excel','Filter bills by date, customer, mobile, invoice and payment; export exactly the filtered list.')
        f=ttk.Frame(self.content); f.pack(fill='x',pady=5); self.sf=tk.StringVar(); self.sfrom=tk.StringVar(); self.sto=tk.StringVar(); self.spay=tk.StringVar(value='All')
        for i,(lab,var,w) in enumerate([('Search',self.sf,20),('From YYYY-MM-DD',self.sfrom,15),('To YYYY-MM-DD',self.sto,15)]): ttk.Label(f,text=lab).grid(row=0,column=i*2); ttk.Entry(f,textvariable=var,width=w).grid(row=0,column=i*2+1,padx=5)
        ttk.Label(f,text='Payment').grid(row=0,column=6); ttk.Combobox(f,textvariable=self.spay,values=['All','Cash','UPI','Card','Credit'],state='readonly',width=10).grid(row=0,column=7,padx=5); ttk.Button(f,text='Filter',command=self.refresh_sales).grid(row=0,column=8,padx=5); ttk.Button(f,text='Download Excel',command=self.export_excel).grid(row=0,column=9,padx=5)
        self.sales_tv=ttk.Treeview(self.content,columns=('invoice','date','time','customer','phone','payment','subtotal','gst','total','sms'),show='headings'); self.sales_tv.pack(fill='both',expand=True,pady=10)
        for c,h in [('invoice','Invoice'),('date','Date'),('time','Time'),('customer','Customer'),('phone','Mobile'),('payment','Payment'),('subtotal','Taxable'),('gst','GST'),('total','Total'),('sms','SMS')]: self.sales_tv.heading(c,text=h)
        self.refresh_sales()
    def sales_rows(self):
        where=[]; args=[]; q=self.sf.get().strip();
        if q: where.append('(invoice LIKE ? OR customer_name LIKE ? OR customer_phone LIKE ?)'); args += [f'%{q}%']*3
        if self.sfrom.get().strip(): where.append('bill_date>=?'); args.append(self.sfrom.get().strip())
        if self.sto.get().strip(): where.append('bill_date<=?'); args.append(self.sto.get().strip())
        if self.spay.get()!='All': where.append('payment=?'); args.append(self.spay.get())
        sql='SELECT invoice,bill_date,bill_time,customer_name,customer_phone,payment,subtotal,gst,total,sms_status FROM bills'+((' WHERE '+' AND '.join(where)) if where else '')+' ORDER BY id DESC'; return self.db.q(sql,args)
    def refresh_sales(self):
        if not hasattr(self,'sales_tv'): return
        for x in self.sales_tv.get_children(): self.sales_tv.delete(x)
        for r in self.sales_rows(): self.sales_tv.insert('', 'end',values=tuple(r))
    def export_excel(self):
        if Workbook is None: messagebox.showerror('Excel','Install openpyxl: pip install openpyxl'); return
        rows=self.sales_rows(); path=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile='MediBill_Bills.xlsx',filetypes=[('Excel','*.xlsx')]);
        if not path:return
        wb=Workbook(); ws=wb.active; ws.title='Bills'; headers=['Invoice','Date','Time','Customer','Mobile','Payment','Taxable','GST','Total','SMS Status']; ws.append(headers)
        for r in rows: ws.append(list(r))
        wi=wb.create_sheet('Bill Items'); wi.append(['Invoice','Medicine','Batch','Qty','Price','GST %','Amount'])
        for r in rows:
            for it in self.db.q('SELECT b.invoice,i.name,i.batch,i.qty,i.price,i.gst,i.amount FROM bill_items i JOIN bills b ON b.id=i.bill_id WHERE b.invoice=?',(r['invoice'],)): wi.append(list(it))
        for wsx in wb.worksheets:
            wsx.freeze_panes='A2'; wsx.auto_filter.ref=wsx.dimensions
            for col in wsx.columns:
                letter=col[0].column_letter; wsx.column_dimensions[letter].width=min(28,max(12,max(len(str(c.value or '')) for c in col)+2))
        wb.save(path); messagebox.showinfo('Excel',f'Exported {len(rows)} bills to:\n{path}')
    def reports(self):
        self.title('Reports','Sales summary from the local SQLite database.')
        today=datetime.now().strftime('%Y-%m-%d'); month=datetime.now().strftime('%Y-%m')
        vals=[('Today',self.db.one('SELECT COALESCE(SUM(total),0) x FROM bills WHERE bill_date=?',(today,))['x']),('This month',self.db.one('SELECT COALESCE(SUM(total),0) x FROM bills WHERE bill_date LIKE ?',(month+'%',))['x']),('All time',self.db.one('SELECT COALESCE(SUM(total),0) x FROM bills')['x'])]
        for a,b in vals: ttk.LabelFrame(self.content,text=a,padding=25).pack(fill='x',pady=5); ttk.Label(self.content.winfo_children()[-1],text=f'₹{money(b):,.2f}',font=('Segoe UI',18,'bold')).pack()
    def users(self):
        if self.user['role'] not in ['Admin','Store Manager']: return
        self.title('Users & Roles','Only Store Manager/Admin can assign or change roles. One email ID has one role.')
        tv=ttk.Treeview(self.content,columns=('id','name','email','role','active'),show='headings'); tv.pack(fill='both',expand=True,pady=8)
        for c,h in [('id','ID'),('name','Name'),('email','Email'),('role','Role'),('active','Active')]: tv.heading(c,text=h)
        for r in self.db.q('SELECT id,name,email,role,active FROM users ORDER BY id'): tv.insert('', 'end',iid=str(r['id']),values=tuple(r))
        bar=ttk.Frame(self.content); bar.pack(fill='x'); role=tk.StringVar(value='Cashier'); ttk.Label(bar,text='New role').pack(side='left'); ttk.Combobox(bar,textvariable=role,values=ROLES[1:],state='readonly').pack(side='left',padx=5)
        def assign():
            s=tv.selection()
            if not s:return
            uid=int(s[0]); r=self.db.one('SELECT * FROM users WHERE id=?',(uid,));
            if r['role']=='Admin' and self.user['role']!='Admin': messagebox.showerror('Access','Only Admin can change an Admin role.'); return
            self.db.conn.execute('UPDATE users SET role=? WHERE id=?',(role.get(),uid)); self.db.conn.commit(); self.users()
        ttk.Button(bar,text='Assign selected',command=assign).pack(side='left')
    def sms_outbox(self):
        self.title('SMS Outbox','Bills are never lost if SMS is unavailable. Configure an HTTP SMS provider in Settings, then resend queued messages.')
        tv=ttk.Treeview(self.content,columns=('invoice','phone','status','created','response'),show='headings'); tv.pack(fill='both',expand=True)
        for c,h in [('invoice','Invoice'),('phone','Mobile'),('status','Status'),('created','Created'),('response','Provider response')]: tv.heading(c,text=h)
        for r in self.db.q('SELECT b.invoice,o.phone,o.status,o.created_at,o.provider_response FROM sms_outbox o LEFT JOIN bills b ON b.id=o.bill_id ORDER BY o.id DESC'): tv.insert('', 'end',values=tuple(r))
    def settings(self):
        self.title('Settings','Local app settings. No backend/server is required.')
        f=ttk.Frame(self.content); f.pack(anchor='w',pady=15); mode=tk.StringVar(value=self.sms.s.get('sms_mode','None')); url=tk.StringVar(value=self.sms.s.get('sms_url','')); token=tk.StringVar(value=self.sms.s.get('sms_token',''))
        ttk.Label(f,text='SMS mode').grid(row=0,column=0,sticky='w',pady=8); ttk.Combobox(f,textvariable=mode,values=['None','HTTP JSON'],state='readonly',width=30).grid(row=0,column=1,pady=8)
        ttk.Label(f,text='SMS HTTP URL').grid(row=1,column=0,sticky='w',pady=8); ttk.Entry(f,textvariable=url,width=60).grid(row=1,column=1,pady=8)
        ttk.Label(f,text='API token').grid(row=2,column=0,sticky='w',pady=8); ttk.Entry(f,textvariable=token,show='•',width=60).grid(row=2,column=1,pady=8)
        ttk.Label(f,text='The app sends POST JSON: {"to":"mobile","message":"..."}. Use your SMS provider endpoint. No separate backend is needed.',wraplength=650).grid(row=3,column=0,columnspan=2,sticky='w',pady=15)
        def save(): self.sms.s.update({'sms_mode':mode.get(),'sms_url':url.get().strip(),'sms_token':token.get().strip()}); self.sms.save(); messagebox.showinfo('Settings','Saved locally.')
        ttk.Button(f,text='Save SMS settings',command=save).grid(row=4,column=1,sticky='w')
        ttk.Separator(self.content).pack(fill='x',pady=20)
        ttk.Label(self.content,text=f'Database: {DB_PATH}',foreground='#666').pack(anchor='w')
    def run(self): self.root.mainloop()

if __name__=='__main__': App().run()
