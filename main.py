import os, re, json, sqlite3, hashlib, secrets, urllib.request, urllib.parse, calendar as calmod, csv, webbrowser, subprocess
from datetime import datetime, date, timedelta
from decimal import Decimal, ROUND_HALF_UP
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    Workbook = load_workbook = None
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.pdfbase.pdfmetrics import stringWidth
except ImportError:
    A4 = canvas = colors = stringWidth = None

APP_NAME = 'MediBill Pro'
DEFAULT_COMPANY = 'MediBill Pharmacy'
APP_DIR = os.path.join(os.path.expanduser('~'), 'MediBillPro')
os.makedirs(APP_DIR, exist_ok=True)
DB_PATH = os.path.join(APP_DIR, 'medibill.db')
SETTINGS_PATH = os.path.join(APP_DIR, 'settings.json')
PDF_DIR = os.path.join(APP_DIR, 'Bills')
os.makedirs(PDF_DIR, exist_ok=True)

ROLES = ['Admin', 'Store Manager', 'Cashier', 'Accounts']
PERMS = {
    'Admin': {'dashboard','billing','inventory','customers','sales','reports','settings','users','sms','account'},
    'Store Manager': {'dashboard','billing','inventory','customers','sales','reports','settings','users','sms','account'},
    'Cashier': {'dashboard','billing','customers','sms','account'},
    'Accounts': {'dashboard','customers','sales','reports','account'},
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
    return len(s) >= 8 and bool(re.search(r'[A-Z]', s)) and bool(re.search(r'[a-z]', s)) and bool(re.search(r'\d', s)) and bool(re.search(r'[^A-Za-z0-9]', s))

def normalize_phone(p):
    return re.sub(r'\D', '', p or '')

def parse_date(s):
    return datetime.strptime(s.strip(), '%Y-%m-%d').date()

def fmt_date(s):
    try: return parse_date(s).strftime('%d-%m-%Y')
    except Exception: return s or '-'

def expiry_status(expiry, stock, today=None):
    today = today or date.today()
    try: d = parse_date(expiry)
    except Exception: return ('In Stock', 3)
    if d < today: return ('Expired', 0)
    if d <= today + timedelta(days=90): return ('Near Expiry', 1)
    if int(stock or 0) <= 10: return ('Low Stock', 2)
    return ('In Stock', 3)

class DB:
    def __init__(self, path=DB_PATH):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute('PRAGMA foreign_keys=ON')
        self.init()
    def init(self):
        c = self.conn.cursor()
        c.executescript('''
        CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, name TEXT NOT NULL, email TEXT UNIQUE NOT NULL, salt TEXT NOT NULL, password_hash TEXT NOT NULL, role TEXT NOT NULL, verified INTEGER DEFAULT 1, active INTEGER DEFAULT 1, must_change INTEGER DEFAULT 0, created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS customers(id INTEGER PRIMARY KEY, name TEXT NOT NULL, phone TEXT UNIQUE NOT NULL, email TEXT DEFAULT '', address TEXT DEFAULT '', credit REAL DEFAULT 0, total_purchase REAL DEFAULT 0, last_purchase TEXT);
        CREATE TABLE IF NOT EXISTS medicines(id INTEGER PRIMARY KEY, name TEXT NOT NULL, batch TEXT DEFAULT '', expiry TEXT DEFAULT '', price REAL DEFAULT 0, cost REAL DEFAULT 0, stock INTEGER DEFAULT 0, gst REAL DEFAULT 0, supplier TEXT DEFAULT '', UNIQUE(name,batch));
        CREATE TABLE IF NOT EXISTS bills(id INTEGER PRIMARY KEY, invoice TEXT UNIQUE NOT NULL, bill_date TEXT NOT NULL, bill_time TEXT NOT NULL, customer_id INTEGER, customer_name TEXT, customer_phone TEXT, payment TEXT, subtotal REAL, gst REAL, total REAL, sms_status TEXT DEFAULT 'Not sent', created_by TEXT, FOREIGN KEY(customer_id) REFERENCES customers(id));
        CREATE TABLE IF NOT EXISTS bill_items(id INTEGER PRIMARY KEY, bill_id INTEGER NOT NULL, medicine_id INTEGER, name TEXT, batch TEXT, qty INTEGER, price REAL, gst REAL, amount REAL, FOREIGN KEY(bill_id) REFERENCES bills(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS sms_outbox(id INTEGER PRIMARY KEY, bill_id INTEGER, phone TEXT, message TEXT, status TEXT, provider_response TEXT DEFAULT '', created_at TEXT, FOREIGN KEY(bill_id) REFERENCES bills(id));
        ''')
        # Migrate older databases created before the credential-change feature.
        cols={r['name'] for r in c.execute('PRAGMA table_info(users)').fetchall()}
        if 'must_change' not in cols:
            c.execute('ALTER TABLE users ADD COLUMN must_change INTEGER DEFAULT 0')
        if not c.execute('SELECT 1 FROM users LIMIT 1').fetchone():
            for name,email,pw,role in [('Administrator','admin@medibillwb.in','Admin@1234','Admin'),('Store Manager','store@medibillwb.in','Demo@1234','Store Manager'),('Cashier','cashier@medibillwb.in','Demo@1234','Cashier'),('Accounts','accounts@medibillwb.in','Demo@1234','Accounts')]:
                salt,d=hash_pw(pw); c.execute('INSERT INTO users(name,email,salt,password_hash,role,verified,active,must_change,created_at) VALUES(?,?,?,?,?,?,?,?,?)',(name,email,salt,d,role,1,1,1 if email=='admin@medibillwb.in' else 0,datetime.now().isoformat(timespec='seconds')))
        if not c.execute('SELECT 1 FROM medicines LIMIT 1').fetchone():
            meds=[('Paracetamol 500mg','PCT001','2027-08-31',25,12,80,5,'Demo Supplier'),('Azithromycin 500mg','AZI001','2027-03-31',85,48,25,5,'Demo Supplier'),('ORS Lemon','ORS001','2028-01-31',22,13,50,5,'Demo Supplier'),('Pantoprazole 40mg','PAN001','2027-11-30',60,30,35,12,'Demo Supplier'),('Vitamin C 500mg','VIT001','2026-11-30',35,20,7,12,'Demo Supplier'),('Cetirizine 10mg','CET001','2026-10-15',18,9,15,5,'Demo Supplier')]
            c.executemany('INSERT INTO medicines(name,batch,expiry,price,cost,stock,gst,supplier) VALUES(?,?,?,?,?,?,?,?)',meds)
        self.conn.commit()
    def q(self,sql,args=()): return self.conn.execute(sql,args).fetchall()
    def one(self,sql,args=()): return self.conn.execute(sql,args).fetchone()
    def close(self): self.conn.close()

class SMS:
    """MSG91 Flow API integration. No backend is required; the desktop app
    calls MSG91 directly over HTTPS.
    """
    def __init__(self):
        self.s={}
        if os.path.exists(SETTINGS_PATH):
            try:
                with open(SETTINGS_PATH,encoding='utf8') as f: self.s=json.load(f)
            except Exception: pass

    def save(self):
        with open(SETTINGS_PATH,'w',encoding='utf8') as f: json.dump(self.s,f,indent=2)

    def configured(self):
        return all(self.s.get(k,'').strip() for k in ('msg91_authkey','msg91_flow_id','msg91_sender_id'))

    def _mobile(self, phone):
        p=normalize_phone(phone)
        if len(p)==10: return '91'+p
        if len(p)==12 and p.startswith('91'): return p
        return p

    def send(self, phone, message, invoice='', customer_name='', total=0):
        if not self.configured():
            return False,'MSG91 is not configured. Configure Auth Key, Flow ID and Sender ID in Settings. The message remains in SMS Outbox.'
        mobile=self._mobile(phone)
        if len(mobile)!=12 or not mobile.startswith('91'):
            return False,'Invalid Indian mobile number. Use a 10-digit mobile number.'
        url='https://control.msg91.com/api/v5/flow'
        payload={
            'flow_id':self.s['msg91_flow_id'].strip(),
            'sender':self.s['msg91_sender_id'].strip(),
            'recipients':[{
                'mobiles':mobile,
                'company':self.s.get('company_name',DEFAULT_COMPANY),
                'name':customer_name or 'Customer',
                'invoice':invoice,
                'amount':f'{float(total):.2f}'
            }]
        }
        req=urllib.request.Request(
            url,
            data=json.dumps(payload).encode('utf-8'),
            headers={'Content-Type':'application/json','accept':'application/json','authkey':self.s['msg91_authkey'].strip()},
            method='POST'
        )
        try:
            with urllib.request.urlopen(req,timeout=20) as r:
                body=r.read().decode(errors='ignore')[:1000]
                try: data=json.loads(body)
                except Exception: data={}
                ok=(200 <= r.status < 300) and data.get('type','success') != 'error'
                return ok,body
        except Exception as e:
            return False,str(e)

class DatePicker(tk.Toplevel):
    def __init__(self, parent, initial='', callback=None):
        super().__init__(parent); self.callback=callback; self.title('Select Date'); self.resizable(False,False); self.transient(parent); self.grab_set()
        try: d=parse_date(initial)
        except Exception: d=date.today()
        self.year=d.year; self.month=d.month
        self.body=tk.Frame(self,padx=12,pady=12); self.body.pack()
        self.draw()
    def draw(self):
        for w in self.body.winfo_children(): w.destroy()
        nav=tk.Frame(self.body); nav.pack(fill='x',pady=(0,8))
        tk.Button(nav,text='‹',command=self.prev,bd=0,font=('Segoe UI',16)).pack(side='left')
        tk.Label(nav,text=f'{calmod.month_name[self.month]} {self.year}',font=('Segoe UI',12,'bold')).pack(side='left',expand=True)
        tk.Button(nav,text='›',command=self.next,bd=0,font=('Segoe UI',16)).pack(side='right')
        grid=tk.Frame(self.body); grid.pack()
        for j,h in enumerate(['Mo','Tu','We','Th','Fr','Sa','Su']): tk.Label(grid,text=h,width=4,font=('Segoe UI',9,'bold')).grid(row=0,column=j)
        for i,week in enumerate(calmod.monthcalendar(self.year,self.month),1):
            for j,day in enumerate(week):
                if day:
                    tk.Button(grid,text=str(day),width=4,bd=0,command=lambda d=day:self.pick(d)).grid(row=i,column=j,padx=1,pady=1)
        tk.Button(self.body,text='Today',command=lambda:self.pick(date.today().day if self.month==date.today().month and self.year==date.today().year else None)).pack(pady=(8,0))
    def prev(self):
        self.month-=1
        if self.month==0:self.month=12;self.year-=1
        self.draw()
    def next(self):
        self.month+=1
        if self.month==13:self.month=1;self.year+=1
        self.draw()
    def pick(self,d):
        if d is None: d=date.today().day; self.year=date.today().year; self.month=date.today().month
        value=f'{self.year:04d}-{self.month:02d}-{d:02d}'
        if self.callback:self.callback(value)
        self.destroy()

class App:
    def __init__(self):
        self.db=DB(); self.sms=SMS(); self.user=None; self.cart=[]; self.last_invoice=None; self.last_bill_id=None; self.company=self.sms.s.get('company_name',DEFAULT_COMPANY)
        self.last_bill=None; self.root=tk.Tk(); self.root.withdraw(); self.root.title(APP_NAME); self.root.geometry('1280x800'); self.root.minsize(1100,700); self.root.protocol('WM_DELETE_WINDOW',self.close)
        self.setup_style(); self.root.deiconify(); self.login(); self.root.deiconify(); self.root.lift(); self.root.focus_force()
    def setup_style(self):
        style=ttk.Style();
        try: style.theme_use('clam')
        except Exception: pass
        style.configure('.',font=('Segoe UI',10)); style.configure('TButton',padding=(12,7),font=('Segoe UI',10,'bold')); style.configure('TEntry',padding=6); style.configure('TCombobox',padding=5)
        style.configure('Treeview',rowheight=30,font=('Segoe UI',10)); style.configure('Treeview.Heading',font=('Segoe UI',10,'bold'))
        style.configure('Title.TLabel',font=('Segoe UI',20,'bold')); style.configure('Sub.TLabel',foreground='#64748b'); style.configure('Card.TFrame',relief='solid',borderwidth=1)
    def close(self): self.db.close(); self.root.destroy()
    def clear(self):
        for w in self.root.winfo_children(): w.destroy()
    def hover(self,btn,normal='#2563eb',over='#1d4ed8'):
        btn.configure(background=normal,foreground='white',activebackground=over,activeforeground='white')
        btn.bind('<Enter>',lambda e:btn.configure(background=over)); btn.bind('<Leave>',lambda e:btn.configure(background=normal))
    def button(self,parent,text,command,kind='primary',**kw):
        palette={'primary':('#2563eb','#1d4ed8'),'success':('#16a34a','#15803d'),'danger':('#dc2626','#b91c1c'),'warning':('#d97706','#b45309'),'dark':('#334155','#1e293b')}
        b=tk.Button(parent,text=text,command=command,bd=0,relief='flat',cursor='hand2',font=('Segoe UI',10,'bold'),**kw); n,o=palette.get(kind,palette['primary']); self.hover(b,n,o); return b
    def login(self):
        self.clear(); self.root.geometry('560x600'); bg=tk.Frame(self.root,bg='#f1f5f9'); bg.pack(fill='both',expand=True)
        card=tk.Frame(bg,bg='white',padx=42,pady=35); card.place(relx=.5,rely=.5,anchor='center',relwidth=.82,relheight=.78)
        tk.Label(card,text='MediBill Pro',bg='white',fg='#0f172a',font=('Segoe UI',28,'bold')).pack(pady=(10,3)); tk.Label(card,text=self.company,bg='white',fg='#64748b',font=('Segoe UI',11)).pack(pady=(0,25))
        tk.Label(card,text='Sign in to continue',bg='white',fg='#334155',font=('Segoe UI',11)).pack(anchor='w')
        email=tk.StringVar(); pw=tk.StringVar();
        tk.Label(card,text='Email',bg='white',fg='#334155').pack(anchor='w',pady=(18,4)); e=tk.Entry(card,textvariable=email,font=('Segoe UI',11),bd=1,relief='solid'); e.pack(fill='x',ipady=7)
        tk.Label(card,text='Password',bg='white',fg='#334155').pack(anchor='w',pady=(12,4)); p=tk.Entry(card,textvariable=pw,show='•',font=('Segoe UI',11),bd=1,relief='solid'); p.pack(fill='x',ipady=7)
        def go(event=None):
            row=self.db.one('SELECT * FROM users WHERE lower(email)=? AND active=1',(email.get().strip().lower(),))
            if not row or not verify_pw(pw.get(),row['salt'],row['password_hash']): messagebox.showerror('Login failed','Invalid email or password.'); return
            self.user=dict(row)
            self.dashboard()
            if int(self.user.get('must_change',0)):
                self.after_login_change_credentials()
        self.button(card,'Login',go,kind='primary').pack(fill='x',pady=20,ipady=3); p.bind('<Return>',go)
        tk.Label(card,text='First time? Worker accounts can be created here and assigned a role by Store Manager/Admin.',bg='white',fg='#64748b',wraplength=400).pack(pady=3)
        self.button(card,'Create Worker Account',self.register,kind='dark').pack(pady=8)
    def register(self):
        win=tk.Toplevel(self.root); win.title('Create Worker Account'); win.geometry('520x500'); win.transient(self.root); win.grab_set(); frm=tk.Frame(win,bg='white',padx=30,pady=25); frm.pack(fill='both',expand=True)
        fields=[]
        for i,label in enumerate(['Full name','Email','Password','Confirm password']):
            tk.Label(frm,text=label,bg='white',anchor='w').grid(row=i*2,column=0,columnspan=2,sticky='w',pady=(7,3)); e=tk.Entry(frm,show='•' if 'password' in label.lower() else '',font=('Segoe UI',10)); e.grid(row=i*2+1,column=0,columnspan=2,sticky='ew',ipady=6); fields.append(e)
        frm.columnconfigure(0,weight=1); tk.Label(frm,text='The initial role is Cashier. Only Store Manager/Admin can change it.',bg='white',fg='#64748b',wraplength=430).grid(row=8,column=0,columnspan=2,pady=15)
        def save():
            name,email,p1,p2=[x.get().strip() for x in fields]
            if not name or not valid_email(email): messagebox.showerror('Validation','Enter a valid name and email.',parent=win); return
            if not strong_pw(p1): messagebox.showerror('Validation','Password needs 8+ chars with upper, lower, number and special character.',parent=win); return
            if p1!=p2: messagebox.showerror('Validation','Passwords do not match.',parent=win); return
            if self.db.one('SELECT 1 FROM users WHERE lower(email)=?',(email.lower(),)): messagebox.showerror('Validation','This email already has an account.',parent=win); return
            salt,d=hash_pw(p1); self.db.conn.execute('INSERT INTO users(name,email,salt,password_hash,role,verified,active,must_change,created_at) VALUES(?,?,?,?,?,?,?,?,?)',(name,email,salt,d,'Cashier',1,1,0,datetime.now().isoformat(timespec='seconds'))); self.db.conn.commit(); win.destroy(); messagebox.showinfo('Created','Account created. A Store Manager/Admin can assign its role.')
        self.button(frm,'Create Account',save,kind='success').grid(row=9,column=0,columnspan=2,pady=8)
    def title(self,title,subtitle=''):
        for w in self.content.winfo_children(): w.destroy()
        head=tk.Frame(self.content,bg='#f8fafc'); head.pack(fill='x',pady=(0,18)); tk.Label(head,text=title,bg='#f8fafc',fg='#0f172a',font=('Segoe UI',22,'bold')).pack(anchor='w'); tk.Label(head,text=subtitle,bg='#f8fafc',fg='#64748b',font=('Segoe UI',10)).pack(anchor='w',pady=(2,0))
    def dashboard(self):
        self.clear(); self.root.geometry('1280x720'); self.root.minsize(1100,650); outer=tk.Frame(self.root,bg='#f8fafc'); outer.pack(fill='both',expand=True)
        side=tk.Frame(outer,bg='#0f172a',width=220); side.pack(side='left',fill='y'); side.pack_propagate(False)
        mainwrap=tk.Frame(outer,bg='#f8fafc'); mainwrap.pack(side='left',fill='both',expand=True)
        self.content_canvas=tk.Canvas(mainwrap,bg='#f8fafc',highlightthickness=0,bd=0)
        content_scroll=ttk.Scrollbar(mainwrap,orient='vertical',command=self.content_canvas.yview)
        self.content_canvas.configure(yscrollcommand=content_scroll.set)
        content_scroll.pack(side='right',fill='y'); self.content_canvas.pack(side='left',fill='both',expand=True)
        self.content=tk.Frame(self.content_canvas,bg='#f8fafc',padx=22,pady=20)
        self._content_window=self.content_canvas.create_window((0,0),window=self.content,anchor='nw')
        self.content.bind('<Configure>',lambda e:self.content_canvas.configure(scrollregion=self.content_canvas.bbox('all')))
        self.content_canvas.bind('<Configure>',lambda e:self.content_canvas.itemconfigure(self._content_window,width=e.width))
        self.content_canvas.bind('<MouseWheel>',lambda e:self.content_canvas.yview_scroll(int(-e.delta/120),'units'))
        self.content.bind('<MouseWheel>',lambda e:self.content_canvas.yview_scroll(int(-e.delta/120),'units'))
        tk.Label(side,text='MediBill',bg='#0f172a',fg='white',font=('Segoe UI',21,'bold')).pack(anchor='w',padx=20,pady=(25,2)); tk.Label(side,text=self.company,bg='#0f172a',fg='#94a3b8',font=('Segoe UI',9),wraplength=180).pack(anchor='w',padx=20,pady=(0,15)); tk.Label(side,text=f"{self.user['name']}\n{self.user['role']}",bg='#1e293b',fg='#e2e8f0',justify='left',anchor='w',padx=12,pady=10).pack(fill='x',padx=12,pady=(0,15))
        items=[('Dashboard','dashboard_home'),('Billing','billing'),('Inventory','inventory'),('Customers','customers'),('Sales','sales'),('Reports','reports'),('SMS Outbox','sms_outbox'),('Settings','settings'),('Users & Roles','users'),('My Account','account')]
        for label,method in items:
            perm='dashboard' if method=='dashboard_home' else method
            if perm not in PERMS.get(self.user['role'],set()): continue
            b=tk.Button(side,text=label,command=getattr(self,method),bg='#0f172a',fg='#cbd5e1',activebackground='#1e40af',activeforeground='white',bd=0,anchor='w',padx=20,pady=10,cursor='hand2',font=('Segoe UI',10,'bold')); b.pack(fill='x',padx=8,pady=2); b.bind('<Enter>',lambda e,b=b:b.configure(bg='#1e293b',fg='white')); b.bind('<Leave>',lambda e,b=b:b.configure(bg='#0f172a',fg='#cbd5e1'))
        self.button(side,'Logout',self.logout,kind='danger').pack(side='bottom',fill='x',padx=18,pady=18)
        self.dashboard_home()
    def logout(self): self.user=None; self.cart=[]; self.login()
    def dashboard_home(self):
        self.title('Dashboard',f'Good {"morning" if datetime.now().hour<12 else "afternoon" if datetime.now().hour<18 else "evening"}, {self.user["name"]}. Here is today\'s pharmacy overview.')
        stats=tk.Frame(self.content,bg='#f8fafc'); stats.pack(fill='x')
        today=datetime.now().strftime('%Y-%m-%d'); vals=[('Today\'s Sales',self.db.one('SELECT COALESCE(SUM(total),0) x FROM bills WHERE bill_date=?',(today,))['x'],'#2563eb'),('Bills Today',self.db.one('SELECT COUNT(*) x FROM bills WHERE bill_date=?',(today,))['x'],'#7c3aed'),('Customers',self.db.one('SELECT COUNT(*) x FROM customers')['x'],'#0891b2'),('Medicines',self.db.one('SELECT COUNT(*) x FROM medicines')['x'],'#16a34a')]
        for i,(label,val,accent) in enumerate(vals):
            c=tk.Frame(stats,bg='white',highlightthickness=1,highlightbackground='#e2e8f0'); c.grid(row=0,column=i,sticky='ew',padx=5); stats.columnconfigure(i,weight=1); tk.Label(c,text=label,bg='white',fg='#64748b',font=('Segoe UI',10)).pack(anchor='w',padx=15,pady=(13,3)); tk.Label(c,text=f'₹{money(val):,.2f}' if 'Sales' in label else str(val),bg='white',fg=accent,font=('Segoe UI',20,'bold')).pack(anchor='w',padx=15,pady=(0,13))
        tk.Label(self.content,text='Medicine Status',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',16,'bold')).pack(anchor='w',pady=(25,8))
        table=tk.Frame(self.content,bg='white'); table.pack(fill='both',expand=True); tv=ttk.Treeview(table,columns=('name','batch','stock','expiry','status'),show='headings'); tv.pack(side='left',fill='both',expand=True); sb=ttk.Scrollbar(table,orient='vertical',command=tv.yview); sb.pack(side='right',fill='y'); tv.configure(yscrollcommand=sb.set)
        for c,h in [('name','Medicine'),('batch','Batch'),('stock','Stock'),('expiry','Expiry'),('status','Status')]: tv.heading(c,text=h); tv.column(c,width=160 if c!='name' else 300)
        tv.tag_configure('near',foreground='#dc2626'); tv.tag_configure('expired',foreground='#991b1b'); tv.tag_configure('low',foreground='#d97706'); tv.tag_configure('instock',foreground='#15803d')
        rows=[]
        for r in self.db.q('SELECT name,batch,stock,expiry FROM medicines ORDER BY name'):
            status,prio=expiry_status(r['expiry'],r['stock']); rows.append((prio,r))
        for _,r in sorted(rows,key=lambda x:(x[0],parse_date(x[1]['expiry']) if re.fullmatch(r'\d{4}-\d{2}-\d{2}',x[1]['expiry'] or '') else date.max,x[1]['name'])):
            status,_=expiry_status(r['expiry'],r['stock']); tag={'Near Expiry':'near','Expired':'expired','Low Stock':'low','In Stock':'instock'}[status]; tv.insert('', 'end',values=(r['name'],r['batch'],r['stock'],fmt_date(r['expiry']),status),tags=(tag,))
    def date_entry(self,parent,var,width=14):
        e=tk.Entry(parent,textvariable=var,font=('Segoe UI',10),bd=1,relief='solid',width=width); e.pack(side='left',ipady=5)
        self.button(parent,'📅',lambda:DatePicker(self.root,var.get(),lambda v:var.set(v)),kind='dark',width=3).pack(side='left',padx=(3,10)); return e
    def billing(self):
        self.title('Billing','Add medicines, adjust quantities, generate the bill, and send SMS only when you choose.')
        self.b_name=tk.StringVar(); self.b_phone=tk.StringVar(); self.b_payment=tk.StringVar(value='Cash'); self.b_search=tk.StringVar(); self.search_after=None
        info=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#e2e8f0',padx=15,pady=12); info.pack(fill='x')
        for i,(lab,var) in enumerate([('Customer name',self.b_name),('Mobile number',self.b_phone)]): tk.Label(info,text=lab,bg='white',fg='#475569').grid(row=0,column=i*2,sticky='w',padx=5); tk.Entry(info,textvariable=var,width=27,bd=1,relief='solid').grid(row=1,column=i*2,columnspan=2,padx=5,ipady=6,sticky='ew')
        tk.Label(info,text='Payment method',bg='white',fg='#475569').grid(row=0,column=4,sticky='w',padx=5); ttk.Combobox(info,textvariable=self.b_payment,values=['Cash','UPI','Card','Credit'],state='readonly',width=14).grid(row=1,column=4,padx=5,sticky='w')
        tk.Label(self.content,text='Find medicine',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(18,6)); sf=tk.Frame(self.content,bg='#f8fafc'); sf.pack(fill='x'); search=tk.Entry(sf,textvariable=self.b_search,font=('Segoe UI',11),bd=1,relief='solid'); search.pack(side='left',fill='x',expand=True,ipady=7); tk.Label(sf,text='  Type to search instantly',bg='#f8fafc',fg='#64748b').pack(side='left')
        results_box=tk.Frame(self.content,bg='#f8fafc'); results_box.pack(fill='x',pady=8)
        self.bill_results=ttk.Treeview(results_box,columns=('id','name','batch','price','stock'),show='headings',height=5); self.bill_results.pack(side='left',fill='both',expand=True)
        rsb=ttk.Scrollbar(results_box,orient='vertical',command=self.bill_results.yview); rsb.pack(side='right',fill='y'); self.bill_results.configure(yscrollcommand=rsb.set)
        for c,h in [('id','ID'),('name','Medicine'),('batch','Batch'),('price','Selling Price'),('stock','Available')]: self.bill_results.heading(c,text=h)
        self.bill_results.bind('<Double-1>',lambda e:self.add_selected()); self.bill_results.bind('<Return>',lambda e:self.add_selected()); self.bill_search()
        tk.Label(self.content,text='Bill items',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(10,6))
        cartbox=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#e2e8f0'); cartbox.pack(fill='x',expand=False)
        cart_table=tk.Frame(cartbox,bg='white'); cart_table.pack(fill='x',expand=False)
        self.cart_tv=ttk.Treeview(cart_table,columns=('name','qty','price','gst','amount'),show='headings',height=5); self.cart_tv.pack(side='left',fill='both',expand=True)
        cart_sb=ttk.Scrollbar(cart_table,orient='vertical',command=self.cart_tv.yview); cart_sb.pack(side='right',fill='y'); self.cart_tv.configure(yscrollcommand=cart_sb.set)
        
        for c,h in [('name','Medicine'),('qty','Quantity'),('price','Price'),('gst','GST %'),('amount','Amount')]: self.cart_tv.heading(c,text=h)
        self.cart_tv.bind('<Double-1>',lambda e:self.edit_cart_qty()); self.cart_tv.bind('<Delete>',lambda e:self.remove_cart())
        actions=tk.Frame(self.content,bg='#e2e8f0',highlightthickness=1,highlightbackground='#cbd5e1'); actions.pack(fill='x',pady=(8,5),ipady=4)
        tk.Label(actions,text='Selected item:',bg='#e2e8f0',fg='#475569',font=('Segoe UI',9,'bold')).pack(side='left',padx=(8,5))
        self.button(actions,'− Decrease',lambda:self.change_qty(-1),kind='dark').pack(side='left',padx=3)
        self.button(actions,'+ Increase',lambda:self.change_qty(1),kind='success').pack(side='left',padx=3)
        self.button(actions,'Remove Item',self.remove_cart,kind='danger').pack(side='left',padx=3)
        self.button(actions,'Clear Bill',self.clear_bill,kind='dark').pack(side='left',padx=3)
        tk.Label(actions,text='Tip: select a row, then use + / − or Delete. Double-click quantity to edit.',bg='#e2e8f0',fg='#64748b').pack(side='right',padx=8)
        foot=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#cbd5e1'); foot.pack(fill='x',pady=(2,0),ipady=5); self.total_lbl=tk.Label(foot,text='Total: ₹0.00',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',20,'bold')); self.total_lbl.pack(side='left',pady=8)
        self.generate_btn=self.button(foot,'Generate Bill',self.checkout,kind='success'); self.generate_btn.pack(side='right',padx=4,ipadx=14,ipady=5); self.sms_btn=self.button(foot,'Send SMS',self.send_last_sms,kind='primary'); self.sms_btn.pack(side='right',padx=4,ipadx=14,ipady=5); self.pdf_btn=self.button(foot,'Print / Save PDF',self.generate_last_pdf,kind='dark'); self.pdf_btn.pack(side='right',padx=4,ipadx=8,ipady=5); self.sms_btn.configure(state='disabled'); self.pdf_btn.configure(state='disabled')
        self.b_search.trace_add('write',lambda *_:self.bill_search()); search.focus_set()
    def bill_search(self):
        if not hasattr(self,'bill_results'): return
        q='%'+self.b_search.get().strip()+'%'
        for x in self.bill_results.get_children(): self.bill_results.delete(x)
        for r in self.db.q('SELECT id,name,batch,price,stock FROM medicines WHERE name LIKE ? OR batch LIKE ? ORDER BY name LIMIT 50',(q,q)): self.bill_results.insert('', 'end',values=tuple(r))
    def add_selected(self):
        s=self.bill_results.selection()
        if not s:return
        rid=int(self.bill_results.item(s[0])['values'][0]); r=self.db.one('SELECT * FROM medicines WHERE id=?',(rid,))
        if r['stock']<=0:messagebox.showwarning('Stock','Out of stock.');return
        for it in self.cart:
            if it['id']==rid:
                if it['qty']<r['stock']:it['qty']+=1
                else:messagebox.showwarning('Stock','Maximum available quantity reached.')
                self.refresh_cart();return
        self.cart.append({'id':r['id'],'name':r['name'],'batch':r['batch'],'qty':1,'price':float(r['price']),'gst':float(r['gst'])});self.refresh_cart()
    def selected_cart_index(self):
        s=self.cart_tv.selection(); return int(s[0]) if s and s[0].isdigit() else None
    def change_qty(self,delta):
        i=self.selected_cart_index()
        if i is None:return
        it=self.cart[i]; stock=self.db.one('SELECT stock FROM medicines WHERE id=?',(it['id'],))['stock']; it['qty']=max(1,min(stock,it['qty']+delta)); self.refresh_cart(); self.cart_tv.selection_set(str(i))
    def edit_cart_qty(self):
        i=self.selected_cart_index()
        if i is None:return
        it=self.cart[i]; win=tk.Toplevel(self.root);win.title('Change Quantity');win.geometry('300x170');win.transient(self.root);win.grab_set();v=tk.IntVar(value=it['qty']);tk.Label(win,text='Quantity',font=('Segoe UI',11,'bold')).pack(pady=(20,6));sp=tk.Spinbox(win,from_=1,to=max(1,int(self.db.one('SELECT stock FROM medicines WHERE id=?',(it['id'],))['stock'])),textvariable=v,width=10,font=('Segoe UI',12));sp.pack();
        def save():it['qty']=int(v.get());self.refresh_cart();win.destroy()
        self.button(win,'Save',save,kind='success').pack(pady=15)
    def remove_cart(self):
        i=self.selected_cart_index()
        if i is not None:self.cart.pop(i);self.refresh_cart()
    def clear_bill(self):
        self.cart.clear();self.b_name.set('');self.b_phone.set('');self.refresh_cart();self.last_invoice=None;self.last_bill_id=None;self.sms_btn.configure(state='disabled');self.pdf_btn.configure(state='disabled')
    def refresh_cart(self):
        if not hasattr(self,'cart_tv'):return
        for x in self.cart_tv.get_children():self.cart_tv.delete(x)
        total=0
        for i,it in enumerate(self.cart):
            amt=money(it['qty']*it['price']*(1+it['gst']/100));total+=amt;self.cart_tv.insert('', 'end',iid=str(i),values=(it['name'],it['qty'],f'₹{it["price"]:.2f}',it['gst'],f'₹{amt:.2f}'))
        self.total_lbl.config(text=f'Total: ₹{total:,.2f}')
    def sms_message(self,invoice,name,total):return f'{self.company}: Dear {name}, your bill {invoice} total is Rs.{total:.2f}. Thank you for shopping with us.'
    def checkout(self):
        if not self.cart:messagebox.showwarning('Bill','Add at least one medicine.');return
        name=self.b_name.get().strip() or 'Walk-in Customer';phone=normalize_phone(self.b_phone.get())
        if phone and len(phone)<10:messagebox.showerror('Mobile','Enter a valid mobile number.');return
        now=datetime.now();invoice='INV-'+now.strftime('%Y%m%d-%H%M%S')+'-'+secrets.token_hex(2).upper();subtotal=gsttotal=total=0
        try:
            cur=self.db.conn.cursor();cur.execute('BEGIN');cid=None
            if phone:
                old=cur.execute('SELECT id FROM customers WHERE phone=?',(phone,)).fetchone()
                if old:cid=old['id'];cur.execute('UPDATE customers SET name=?,last_purchase=? WHERE id=?',(name,now.isoformat(timespec='seconds'),cid))
                else:cur.execute('INSERT INTO customers(name,phone,last_purchase) VALUES(?,?,?)',(name,phone,now.isoformat(timespec='seconds')));cid=cur.lastrowid
            for it in self.cart:
                stock=cur.execute('SELECT stock FROM medicines WHERE id=?',(it['id'],)).fetchone()['stock']
                if stock<it['qty']:raise ValueError(f'Insufficient stock for {it["name"]}')
                base=Decimal(str(it['qty']*it['price']));tax=(base*Decimal(str(it['gst']))/Decimal('100')).quantize(Decimal('0.01'));subtotal+=float(base);gsttotal+=float(tax);total+=float(base+tax)
            cur.execute('INSERT INTO bills(invoice,bill_date,bill_time,customer_id,customer_name,customer_phone,payment,subtotal,gst,total,sms_status,created_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',(invoice,now.strftime('%Y-%m-%d'),now.strftime('%H:%M:%S'),cid,name,phone,self.b_payment.get(),money(subtotal),money(gsttotal),money(total),'Not sent' if phone else 'No mobile',self.user['email']));bid=cur.lastrowid
            for it in self.cart:
                base=money(it['qty']*it['price']);tax=money(base*it['gst']/100);cur.execute('INSERT INTO bill_items(bill_id,medicine_id,name,batch,qty,price,gst,amount) VALUES(?,?,?,?,?,?,?,?)',(bid,it['id'],it['name'],it['batch'],it['qty'],it['price'],it['gst'],money(base+tax)));cur.execute('UPDATE medicines SET stock=stock-? WHERE id=?',(it['qty'],it['id']))
            if cid:cur.execute('UPDATE customers SET total_purchase=total_purchase+? WHERE id=?',(money(total),cid))
            if phone:cur.execute('INSERT INTO sms_outbox(bill_id,phone,message,status,created_at) VALUES(?,?,?,?,?)',(bid,phone,self.sms_message(invoice,name,total),'Queued',now.isoformat(timespec='seconds')))
            self.db.conn.commit()
        except Exception as e:self.db.conn.rollback();messagebox.showerror('Could not generate bill',str(e));return
        self.last_invoice=invoice;self.last_bill_id=bid;self.last_bill={'invoice':invoice,'date':now.strftime('%Y-%m-%d'),'time':now.strftime('%H:%M:%S'),'name':name,'phone':phone,'payment':self.b_payment.get(),'subtotal':money(subtotal),'gst':money(gsttotal),'total':money(total),'items':[dict(x) for x in self.cart]}
        self.cart.clear();self.refresh_cart();self.bill_search();self.pdf_btn.configure(state='normal');self.sms_btn.configure(state='normal' if phone else 'disabled');
        messagebox.showinfo('Bill generated',f'Invoice: {invoice}\nTotal: ₹{total:,.2f}\nCustomer saved: {"Yes" if phone else "No mobile supplied"}\n\nChoose “Print / Save PDF” or “Send SMS” separately.')
    def send_last_sms(self):
        if not self.last_bill or not self.last_bill.get('phone'):messagebox.showwarning('SMS','No customer mobile number is available for this bill.');return
        ok,resp=self.sms.send(self.last_bill['phone'],self.sms_message(self.last_bill['invoice'],self.last_bill['name'],self.last_bill['total']),invoice=self.last_bill['invoice'],customer_name=self.last_bill['name'],total=self.last_bill['total']);status='Sent' if ok else 'Queued';self.db.conn.execute('UPDATE bills SET sms_status=? WHERE id=?',(status,self.last_bill_id));self.db.conn.execute('UPDATE sms_outbox SET status=?,provider_response=? WHERE bill_id=?',(status,resp,self.last_bill_id));self.db.conn.commit();messagebox.showinfo('SMS',('SMS sent successfully.' if ok else 'SMS could not be sent; it remains in the local SMS Outbox.')+f'\n\n{resp}')
    def generate_last_pdf(self):
        if not self.last_bill:messagebox.showwarning('PDF','Generate a bill first.');return
        path=self.create_pdf(self.last_bill)
        if path:
            if messagebox.askyesno('PDF ready',f'PDF saved as:\n{path}\n\nOpen it now?'):self.open_file(path)
    def create_pdf(self,b):
        if canvas is None:messagebox.showerror('PDF','PDF support is unavailable. Please use the packaged app or install reportlab for source mode.');return None
        stamp=f"{b['date']}_{b['time'].replace(':','-')}";path=os.path.join(PDF_DIR,f"{b['invoice']}-{stamp}.pdf");c=canvas.Canvas(path,pagesize=A4);W,H=A4
        c.setFillColor(colors.HexColor('#e2e8f0'));c.saveState();c.translate(W/2,H/2);c.rotate(35);c.setFont('Helvetica-Bold',42);c.drawCentredString(0,0,self.company);c.restoreState();
        c.setFillColor(colors.HexColor('#0f172a'));c.setFont('Helvetica-Bold',20);c.drawString(45,H-55,self.company);c.setFont('Helvetica',10);c.drawString(45,H-72,'Pharmacy Billing & Inventory');c.setFont('Helvetica-Bold',14);c.drawRightString(W-45,H-60,'TAX INVOICE');c.setFont('Helvetica',10);c.drawRightString(W-45,H-76,f"Invoice: {b['invoice']}");c.drawRightString(W-45,H-90,f"Date: {fmt_date(b['date'])}  Time: {b['time']}")
        y=H-125;c.setStrokeColor(colors.HexColor('#cbd5e1'));c.line(45,y,W-45,y);y-=25;c.setFont('Helvetica-Bold',10);c.drawString(45,y,'Customer');c.setFont('Helvetica',10);c.drawString(110,y,b['name']);c.drawString(320,y,'Mobile: '+(b['phone'] or '-'));y-=25
        headers=['Medicine','Batch','Qty','Price','GST','Amount'];xs=[45,275,360,405,465,520];c.setFont('Helvetica-Bold',9)
        for x,h in zip(xs,headers):c.drawString(x,y,h)
        y-=7;c.line(45,y,W-45,y);y-=18;c.setFont('Helvetica',9)
        for it in b['items']:
            vals=[it['name'],it['batch'],str(it['qty']),f"Rs.{it['price']:.2f}",f"{it['gst']:.1f}%",f"Rs.{money(it['qty']*it['price']*(1+it['gst']/100)):.2f}"]
            for x,v in zip(xs,vals):c.drawString(x,y,str(v)[:34])
            y-=18
            if y<100:c.showPage();y=H-60
        y-=8;c.line(350,y,W-45,y);y-=20;c.setFont('Helvetica',10);c.drawRightString(W-130,y,'Taxable Amount:');c.drawRightString(W-45,y,f"Rs.{b['subtotal']:.2f}");y-=18;c.drawRightString(W-130,y,'GST:');c.drawRightString(W-45,y,f"Rs.{b['gst']:.2f}");y-=22;c.setFont('Helvetica-Bold',13);c.drawRightString(W-130,y,'Grand Total:');c.drawRightString(W-45,y,f"Rs.{b['total']:.2f}");y-=30;c.setFont('Helvetica',9);c.drawString(45,y,'Payment: '+b['payment']);c.drawString(45,y-18,'Thank you for your purchase.');c.save();return path
    def open_file(self,path):
        try:
            if os.name=='nt':os.startfile(path)
            else:webbrowser.open('file://'+os.path.abspath(path))
        except Exception:pass
    def inventory(self):
        self.title('Inventory','Add medicines manually or import many medicines at once from Excel / CSV.')
        top=tk.Frame(self.content,bg='#f8fafc');top.pack(fill='x',pady=(0,10));self.button(top,'Bulk Import Excel / CSV',self.bulk_import,kind='primary').pack(side='left');tk.Label(top,text='Expected columns: Name, Batch, Expiry, Selling Price, Cost, Stock, GST %, Supplier',bg='#f8fafc',fg='#64748b').pack(side='left',padx=12)
        form=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#e2e8f0',padx=12,pady=10);form.pack(fill='x');vars=[tk.StringVar() for _ in range(8)];labels=['Name','Batch','Expiry','Selling price','Cost','Stock','GST %','Supplier']
        for i,(lab,var) in enumerate(zip(labels,vars)):
            col=i%4;row=(i//4)*2;tk.Label(form,text=lab,bg='white',fg='#475569').grid(row=row,column=col,sticky='w',padx=5,pady=(0,3));cell=tk.Frame(form,bg='white');cell.grid(row=row+1,column=col,sticky='w',padx=5,pady=(0,5));
            if lab=='Expiry':self.date_entry(cell,var,15)
            else:tk.Entry(cell,textvariable=var,width=20,bd=1,relief='solid').pack(ipady=5)
        def add():
            try:
                name,batch,expiry,sprice,cost,stock,gst,supplier=[v.get().strip() for v in vars]
                parse_date(expiry);vals=(name,batch,expiry,float(sprice or 0),float(cost or 0),int(stock or 0),float(gst or 0),supplier)
                self.db.conn.execute('INSERT INTO medicines(name,batch,expiry,price,cost,stock,gst,supplier) VALUES(?,?,?,?,?,?,?,?)',vals);self.db.conn.commit();refresh();[v.set('') for v in vars]
            except Exception as e:messagebox.showerror('Inventory','Please check the medicine fields.\n\n'+str(e))
        self.button(form,'Add Medicine',add,kind='success').grid(row=4,column=0,padx=5,pady=5,sticky='w')
        tvbox=tk.Frame(self.content,bg='white');tvbox.pack(fill='both',expand=True,pady=10)
        tv=ttk.Treeview(tvbox,columns=('id','name','batch','expiry','price','stock','gst','supplier'),show='headings');tv.pack(side='left',fill='both',expand=True)
        tvsb=ttk.Scrollbar(tvbox,orient='vertical',command=tv.yview);tvsb.pack(side='right',fill='y');tv.configure(yscrollcommand=tvsb.set)
        for c,h in [('id','ID'),('name','Medicine'),('batch','Batch'),('expiry','Expiry'),('price','Selling Price'),('stock','Stock'),('gst','GST %'),('supplier','Supplier')]:tv.heading(c,text=h)
        def refresh():
            for x in tv.get_children():tv.delete(x)
            for r in self.db.q('SELECT id,name,batch,expiry,price,stock,gst,supplier FROM medicines ORDER BY name'):tv.insert('', 'end',values=(r['id'],r['name'],r['batch'],fmt_date(r['expiry']),f"₹{r['price']:.2f}",r['stock'],r['gst'],r['supplier']))
        refresh()
    def bulk_import(self):
        path=filedialog.askopenfilename(filetypes=[('Excel files','*.xlsx'),('CSV files','*.csv')]);
        if not path:return
        records=[]
        try:
            if path.lower().endswith('.csv'):
                with open(path,'r',encoding='utf-8-sig',newline='') as f:records=list(csv.DictReader(f))
            else:
                if load_workbook is None:raise RuntimeError('Excel support is unavailable in this source build.')
                ws=load_workbook(path,data_only=True).active;rows=list(ws.iter_rows(values_only=True));headers=[str(x or '').strip() for x in rows[0]];records=[dict(zip(headers,r)) for r in rows[1:] if any(x is not None for x in r)]
            aliases={'name':['name','medicine','medicine name'],'batch':['batch','batch no'],'expiry':['expiry','expiry date'],'price':['selling price','price','mrp'],'cost':['cost','purchase price'],'stock':['stock','quantity','qty'],'gst':['gst','gst %','gst percent'],'supplier':['supplier','supplier name']}
            inserted=updated=errors=0
            def get(rec,key):
                for k in aliases[key]:
                    for actual in rec:
                        if str(actual).strip().lower()==k:return rec[actual]
                return ''
            for rec in records:
                try:
                    name=str(get(rec,'name') or '').strip();batch=str(get(rec,'batch') or '').strip();expiry=str(get(rec,'expiry') or '').strip()
                    if isinstance(get(rec,'expiry'),datetime):expiry=get(rec,'expiry').strftime('%Y-%m-%d')
                    elif isinstance(get(rec,'expiry'),date):expiry=get(rec,'expiry').strftime('%Y-%m-%d')
                    parse_date(expiry)
                    vals=(name,batch,expiry,float(get(rec,'price') or 0),float(get(rec,'cost') or 0),int(float(get(rec,'stock') or 0)),float(get(rec,'gst') or 0),str(get(rec,'supplier') or ''))
                    if not name:raise ValueError('missing Name')
                    old=self.db.one('SELECT id FROM medicines WHERE name=? AND batch=?',(name,batch))
                    if old:self.db.conn.execute('UPDATE medicines SET expiry=?,price=?,cost=?,stock=?,gst=?,supplier=? WHERE id=?',(expiry,vals[3],vals[4],vals[5],vals[6],vals[7],old['id']));updated+=1
                    else:self.db.conn.execute('INSERT INTO medicines(name,batch,expiry,price,cost,stock,gst,supplier) VALUES(?,?,?,?,?,?,?,?)',vals);inserted+=1
                except Exception:errors+=1
            self.db.conn.commit();self.inventory();messagebox.showinfo('Bulk import',f'Imported: {inserted}\nUpdated: {updated}\nSkipped/invalid: {errors}')
        except Exception as e:messagebox.showerror('Bulk import failed',str(e))
    def customers(self):
        self.title('Customers','Every completed bill with a mobile number automatically saves or updates the customer.')
        f=tk.Frame(self.content,bg='#f8fafc');f.pack(fill='x',pady=(0,8));q=tk.StringVar();ent=tk.Entry(f,textvariable=q,width=45,bd=1,relief='solid');ent.pack(side='left',ipady=6);tk.Label(f,text='  Instant search',bg='#f8fafc',fg='#64748b').pack(side='left')
        tvbox=tk.Frame(self.content,bg='white');tvbox.pack(fill='both',expand=True);tv=ttk.Treeview(tvbox,columns=('name','phone','email','credit','total','last'),show='headings');tv.pack(side='left',fill='both',expand=True);tv_sb=ttk.Scrollbar(tvbox,orient='vertical',command=tv.yview);tv_sb.pack(side='right',fill='y');tv.configure(yscrollcommand=tv_sb.set)
        for c,h in [('name','Name'),('phone','Mobile'),('email','Email'),('credit','Credit'),('total','Total Purchase'),('last','Last Purchase')]:tv.heading(c,text=h)
        def refresh(*_):
            for x in tv.get_children():tv.delete(x)
            x='%'+q.get().strip()+'%'
            for r in self.db.q('SELECT name,phone,email,credit,total_purchase,last_purchase FROM customers WHERE name LIKE ? OR phone LIKE ? ORDER BY last_purchase DESC',(x,x)):tv.insert('', 'end',values=(r['name'],r['phone'],r['email'],f"₹{r['credit']:.2f}",f"₹{r['total_purchase']:.2f}",r['last_purchase'] or '-'))
        q.trace_add('write',refresh);refresh()
    def sales(self):
        self.title('Sales','Search and filter all bills. Export the filtered bills and their line items to Excel.')
        f=tk.Frame(self.content,bg='#f8fafc');f.pack(fill='x',pady=(0,8));self.sf=tk.StringVar();self.sfrom=tk.StringVar();self.sto=tk.StringVar();self.spay=tk.StringVar(value='All')
        tk.Label(f,text='Search',bg='#f8fafc').pack(side='left');tk.Entry(f,textvariable=self.sf,width=22,bd=1,relief='solid').pack(side='left',padx=5,ipady=5);tk.Label(f,text='From',bg='#f8fafc').pack(side='left');self.date_entry(f,self.sfrom,11);tk.Label(f,text='To',bg='#f8fafc').pack(side='left');self.date_entry(f,self.sto,11);tk.Label(f,text='Payment',bg='#f8fafc').pack(side='left');ttk.Combobox(f,textvariable=self.spay,values=['All','Cash','UPI','Card','Credit'],state='readonly',width=10).pack(side='left',padx=5);self.button(f,'Export Excel',self.export_excel,kind='success').pack(side='left',padx=5)
        for v in [self.sf,self.sfrom,self.sto,self.spay]:v.trace_add('write',lambda *_:self.refresh_sales())
        tvbox=tk.Frame(self.content,bg='white');tvbox.pack(fill='both',expand=True);self.sales_tv=ttk.Treeview(tvbox,columns=('invoice','date','time','customer','phone','payment','subtotal','gst','total','sms'),show='headings');self.sales_tv.pack(side='left',fill='both',expand=True);sales_sb=ttk.Scrollbar(tvbox,orient='vertical',command=self.sales_tv.yview);sales_sb.pack(side='right',fill='y');self.sales_tv.configure(yscrollcommand=sales_sb.set)
        for c,h in [('invoice','Invoice'),('date','Date'),('time','Time'),('customer','Customer'),('phone','Mobile'),('payment','Payment'),('subtotal','Taxable'),('gst','GST'),('total','Total'),('sms','SMS')]:self.sales_tv.heading(c,text=h)
        self.sales_tv.bind('<Double-1>',lambda e:self.open_selected_bill_pdf());self.refresh_sales()
    def sales_rows(self):
        where=[];args=[];q=self.sf.get().strip()
        if q:where.append('(invoice LIKE ? OR customer_name LIKE ? OR customer_phone LIKE ?)');args += [f'%{q}%']*3
        if self.sfrom.get().strip():where.append('bill_date>=?');args.append(self.sfrom.get().strip())
        if self.sto.get().strip():where.append('bill_date<=?');args.append(self.sto.get().strip())
        if self.spay.get()!='All':where.append('payment=?');args.append(self.spay.get())
        return self.db.q('SELECT invoice,bill_date,bill_time,customer_name,customer_phone,payment,subtotal,gst,total,sms_status FROM bills'+((' WHERE '+' AND '.join(where)) if where else '')+' ORDER BY id DESC',args)
    def refresh_sales(self):
        if not hasattr(self,'sales_tv'):return
        for x in self.sales_tv.get_children():self.sales_tv.delete(x)
        for r in self.sales_rows():self.sales_tv.insert('', 'end',values=(r['invoice'],fmt_date(r['bill_date']),r['bill_time'],r['customer_name'],r['customer_phone'],r['payment'],f"₹{r['subtotal']:.2f}",f"₹{r['gst']:.2f}",f"₹{r['total']:.2f}",r['sms_status']))
    def export_excel(self):
        if Workbook is None:messagebox.showerror('Excel','Excel support is unavailable.');return
        rows=self.sales_rows();path=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile=f'MediBill_Bills_{datetime.now():%Y%m%d_%H%M%S}.xlsx',filetypes=[('Excel','*.xlsx')]);
        if not path:return
        wb=Workbook();ws=wb.active;ws.title='Bills';ws.append(['Invoice','Date','Time','Customer','Mobile','Payment','Taxable','GST','Total','SMS Status'])
        for r in rows:ws.append([r['invoice'],r['bill_date'],r['bill_time'],r['customer_name'],r['customer_phone'],r['payment'],r['subtotal'],r['gst'],r['total'],r['sms_status']])
        wi=wb.create_sheet('Bill Items');wi.append(['Invoice','Medicine','Batch','Qty','Price','GST %','Amount'])
        for r in rows:
            for it in self.db.q('SELECT b.invoice,i.name,i.batch,i.qty,i.price,i.gst,i.amount FROM bill_items i JOIN bills b ON b.id=i.bill_id WHERE b.invoice=?',(r['invoice'],)):wi.append(list(it))
        for wsx in wb.worksheets:
            wsx.freeze_panes='A2';wsx.auto_filter.ref=wsx.dimensions
            for col in wsx.columns:
                letter=col[0].column_letter;wsx.column_dimensions[letter].width=min(30,max(12,max(len(str(c.value or '')) for c in col)+2))
        wb.save(path);messagebox.showinfo('Excel',f'Exported {len(rows)} bills.')
    def open_selected_bill_pdf(self):
        s=self.sales_tv.selection();
        if not s:return
        invoice=self.sales_tv.item(s[0])['values'][0];r=self.db.one('SELECT * FROM bills WHERE invoice=?',(invoice,));items=[dict(x) for x in self.db.q('SELECT * FROM bill_items WHERE bill_id=?',(r['id'],))];b={'invoice':r['invoice'],'date':r['bill_date'],'time':r['bill_time'],'name':r['customer_name'],'phone':r['customer_phone'],'payment':r['payment'],'subtotal':r['subtotal'],'gst':r['gst'],'total':r['total'],'items':items};p=self.create_pdf(b)
        if p and messagebox.askyesno('PDF ready','Open the PDF now?'):self.open_file(p)
    def reports(self):
        self.title('Reports','Sales summary from your local database.')
        cards=tk.Frame(self.content,bg='#f8fafc');cards.pack(fill='x');today=datetime.now().strftime('%Y-%m-%d');month=datetime.now().strftime('%Y-%m')
        vals=[('Today',self.db.one('SELECT COALESCE(SUM(total),0) x FROM bills WHERE bill_date=?',(today,))['x']),('This Month',self.db.one('SELECT COALESCE(SUM(total),0) x FROM bills WHERE bill_date LIKE ?',(month+'%',))['x']),('All Time',self.db.one('SELECT COALESCE(SUM(total),0) x FROM bills')['x'])]
        for i,(a,b) in enumerate(vals):lf=tk.LabelFrame(cards,text=a,bg='white',fg='#64748b',padx=30,pady=15);lf.grid(row=0,column=i,sticky='ew',padx=5);cards.columnconfigure(i,weight=1);tk.Label(lf,text=f'₹{money(b):,.2f}',bg='white',fg='#2563eb',font=('Segoe UI',19,'bold')).pack()
        tk.Label(self.content,text='Top Medicines by Quantity Sold',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',14,'bold')).pack(anchor='w',pady=(25,8));rbox=tk.Frame(self.content,bg='white');rbox.pack(fill='both',expand=True);tv=ttk.Treeview(rbox,columns=('medicine','qty','sales'),show='headings');tv.pack(side='left',fill='both',expand=True);r_sb=ttk.Scrollbar(rbox,orient='vertical',command=tv.yview);r_sb.pack(side='right',fill='y');tv.configure(yscrollcommand=r_sb.set);tv.heading('medicine',text='Medicine');tv.heading('qty',text='Units Sold');tv.heading('sales',text='Sales Value')
        for r in self.db.q('SELECT name,SUM(qty) qty,SUM(amount) sales FROM bill_items GROUP BY name ORDER BY qty DESC LIMIT 20'):tv.insert('', 'end',values=(r['name'],r['qty'],f"₹{r['sales']:.2f}"))
    def users(self):
        if self.user['role'] not in ['Admin','Store Manager']:return
        self.title('Users & Roles','One email ID has one role. The Store Manager assigns worker roles; Admin has full access but does not assign roles.')
        tvbox=tk.Frame(self.content,bg='white');tvbox.pack(fill='both',expand=True,pady=8);tv=ttk.Treeview(tvbox,columns=('id','name','email','role','active'),show='headings');tv.pack(side='left',fill='both',expand=True);u_sb=ttk.Scrollbar(tvbox,orient='vertical',command=tv.yview);u_sb.pack(side='right',fill='y');tv.configure(yscrollcommand=u_sb.set)
        for c,h in [('id','ID'),('name','Name'),('email','Email'),('role','Role'),('active','Active')]:tv.heading(c,text=h)
        for r in self.db.q('SELECT id,name,email,role,active FROM users ORDER BY id'):tv.insert('', 'end',iid=str(r['id']),values=tuple(r))
        bar=tk.Frame(self.content,bg='#f8fafc');bar.pack(fill='x');role=tk.StringVar(value='Cashier');tk.Label(bar,text='Role for selected user',bg='#f8fafc').pack(side='left');ttk.Combobox(bar,textvariable=role,values=ROLES[1:],state='readonly',width=18).pack(side='left',padx=5);assign=self.button(bar,'Assign Role',lambda:self.assign_role(tv,role),kind='primary');assign.pack(side='left');
        if self.user['role']!='Store Manager': assign.configure(state='disabled',bg='#94a3b8',cursor='arrow')
        self.button(bar,'Reset Password',lambda:self.reset_user_password(tv),kind='warning').pack(side='left',padx=5); self.button(bar,'Activate / Deactivate',lambda:self.toggle_user(tv),kind='dark').pack(side='left',padx=5)
    def assign_role(self,tv,role):
        s=tv.selection()
        if not s:return
        uid=int(s[0]); r=self.db.one('SELECT * FROM users WHERE id=?',(uid,))
        if not r:return
        if uid==self.user['id']:
            messagebox.showwarning('Users','Change your own credentials from My Account. You cannot change your own role here.'); return
        if self.user['role']!='Store Manager':
            messagebox.showerror('Access','Only the Store Manager can assign or change worker roles. Admin has full page access but does not assign roles.'); return
        if r['role']=='Admin':
            messagebox.showerror('Access','Admin role cannot be assigned or changed from this screen.'); return
        self.db.conn.execute('UPDATE users SET role=? WHERE id=?',(role.get(),uid)); self.db.conn.commit(); self.users()

    def reset_user_password(self,tv):
        s=tv.selection()
        if not s:return
        uid=int(s[0]); r=self.db.one('SELECT * FROM users WHERE id=?',(uid,))
        if not r:return
        if r['role']=='Admin' and self.user['role']!='Admin': messagebox.showerror('Access','Only Admin can reset an Admin password.'); return
        win=tk.Toplevel(self.root); win.title('Reset Worker Password'); win.geometry('430x300'); win.transient(self.root); win.grab_set(); frm=tk.Frame(win,bg='white',padx=25,pady=20); frm.pack(fill='both',expand=True)
        tk.Label(frm,text=f'Reset password for {r["name"]}',bg='white',font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(0,15))
        e1=tk.Entry(frm,show='•',width=35);e2=tk.Entry(frm,show='•',width=35)
        tk.Label(frm,text='New password',bg='white').pack(anchor='w',pady=3);e1.pack(fill='x',ipady=6)
        tk.Label(frm,text='Confirm password',bg='white').pack(anchor='w',pady=8);e2.pack(fill='x',ipady=6)
        def save():
            if not strong_pw(e1.get()):messagebox.showerror('Validation','Password needs 8+ characters with uppercase, lowercase, number and special character.',parent=win);return
            if e1.get()!=e2.get():messagebox.showerror('Validation','Passwords do not match.',parent=win);return
            salt,d=hash_pw(e1.get());self.db.conn.execute('UPDATE users SET salt=?,password_hash=?,must_change=1 WHERE id=?',(salt,d,uid));self.db.conn.commit();win.destroy();messagebox.showinfo('Password reset','Password reset. The worker will be asked to change it at next login.')
        self.button(frm,'Reset Password',save,kind='success').pack(anchor='w',pady=18)

    def toggle_user(self,tv):
        s=tv.selection()
        if not s:return
        uid=int(s[0]);r=self.db.one('SELECT * FROM users WHERE id=?',(uid,))
        if not r:return
        if uid==self.user['id']:messagebox.showwarning('Users','You cannot deactivate your own account.');return
        self.db.conn.execute('UPDATE users SET active=? WHERE id=?',(0 if r['active'] else 1,uid));self.db.conn.commit();self.users()

    def sms_outbox(self):
        self.title('SMS Outbox','Bills are saved even when SMS is unavailable. Send a queued SMS again after configuring your provider.')
        tvbox=tk.Frame(self.content,bg='white');tvbox.pack(fill='both',expand=True);tv=ttk.Treeview(tvbox,columns=('invoice','phone','status','created','response'),show='headings');tv.pack(side='left',fill='both',expand=True);sms_sb=ttk.Scrollbar(tvbox,orient='vertical',command=tv.yview);sms_sb.pack(side='right',fill='y');tv.configure(yscrollcommand=sms_sb.set)
        for c,h in [('invoice','Invoice'),('phone','Mobile'),('status','Status'),('created','Created'),('response','Provider Response')]:tv.heading(c,text=h)
        for r in self.db.q('SELECT b.invoice,o.phone,o.status,o.created_at,o.provider_response FROM sms_outbox o LEFT JOIN bills b ON b.id=o.bill_id ORDER BY o.id DESC'):tv.insert('', 'end',values=tuple(r))
    def settings(self):
        self.title('Settings','Company information and direct MSG91 SMS configuration. No backend server is required.')
        box=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#e2e8f0',padx=18,pady=18);box.pack(fill='x')
        company=tk.StringVar(value=self.company)
        tk.Label(box,text='Company / Shop Name',bg='white',font=('Segoe UI',10,'bold')).grid(row=0,column=0,sticky='w',pady=8)
        tk.Entry(box,textvariable=company,width=55,bd=1,relief='solid').grid(row=0,column=1,padx=12,ipady=6,sticky='w')
        self.button(box,'Save Company Name',lambda:self.save_company(company),kind='success').grid(row=0,column=2,padx=5)

        tk.Label(box,text='MSG91 SMS',bg='white',fg='#0f172a',font=('Segoe UI',14,'bold')).grid(row=2,column=0,columnspan=3,sticky='w',pady=(25,5))
        tk.Label(box,text='Auth Key',bg='white',font=('Segoe UI',10,'bold')).grid(row=3,column=0,sticky='w',pady=7)
        auth=tk.StringVar(value=self.sms.s.get('msg91_authkey',''))
        tk.Entry(box,textvariable=auth,show='•',width=55,bd=1,relief='solid').grid(row=3,column=1,columnspan=2,padx=12,ipady=6,sticky='w')
        tk.Label(box,text='Flow ID / Template ID',bg='white',font=('Segoe UI',10,'bold')).grid(row=4,column=0,sticky='w',pady=7)
        flow=tk.StringVar(value=self.sms.s.get('msg91_flow_id',''))
        tk.Entry(box,textvariable=flow,width=55,bd=1,relief='solid').grid(row=4,column=1,columnspan=2,padx=12,ipady=6,sticky='w')
        tk.Label(box,text='Sender ID / Header',bg='white',font=('Segoe UI',10,'bold')).grid(row=5,column=0,sticky='w',pady=7)
        sender=tk.StringVar(value=self.sms.s.get('msg91_sender_id',''))
        tk.Entry(box,textvariable=sender,width=55,bd=1,relief='solid').grid(row=5,column=1,columnspan=2,padx=12,ipady=6,sticky='w')
        tk.Label(box,text='MSG91 template variables expected by this app:  company, name, invoice, amount',bg='white',fg='#475569',wraplength=780).grid(row=6,column=0,columnspan=3,sticky='w',pady=(12,4))
        tk.Label(box,text='Your DLT-approved MSG91 template must contain these variables (for example ##name##, ##invoice##, ##amount##). The actual DLT text must match your approved template.',bg='white',fg='#64748b',wraplength=780,justify='left').grid(row=7,column=0,columnspan=3,sticky='w',pady=(0,12))
        self.button(box,'Save MSG91 Settings',lambda:self.save_sms(auth,flow,sender),kind='primary').grid(row=8,column=1,sticky='w',padx=12)
        test_frame=tk.Frame(box,bg='white');test_frame.grid(row=8,column=2,sticky='e',padx=5)
        test_phone=tk.StringVar()
        tk.Entry(test_frame,textvariable=test_phone,width=18,bd=1,relief='solid').pack(side='left',ipady=5)
        self.button(test_frame,'Send Test SMS',lambda:self.test_msg91(test_phone),kind='dark').pack(side='left',padx=5)

        tk.Label(self.content,text='PDF / Database',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(22,6))
        tk.Label(self.content,text=f'Bills PDF folder: {PDF_DIR}\nSQLite database: {DB_PATH}\nCompany name is printed on bills and used as the PDF watermark.',bg='#f8fafc',fg='#64748b',justify='left').pack(anchor='w')

    def save_company(self,var):
        name=var.get().strip() or DEFAULT_COMPANY;self.company=name;self.sms.s['company_name']=name;self.sms.save();messagebox.showinfo('Company','Company name saved. It will appear on the app and future PDFs.');self.dashboard()
    def save_sms(self,auth,flow,sender):
        self.sms.s.update({'msg91_authkey':auth.get().strip(),'msg91_flow_id':flow.get().strip(),'msg91_sender_id':sender.get().strip(),'company_name':self.company})
        self.sms.save()
        if self.sms.configured(): messagebox.showinfo('MSG91','MSG91 settings saved locally. You can use Send Test SMS to verify the configuration.')
        else: messagebox.showwarning('MSG91','Settings saved, but Auth Key, Flow ID and Sender ID are all required before SMS can be sent.')

    def test_msg91(self,var):
        phone=normalize_phone(var.get())
        if len(phone) not in (10,12):
            messagebox.showerror('Test SMS','Enter a valid 10-digit Indian mobile number.')
            return
        ok,resp=self.sms.send(phone,'MSG91 test message',invoice='TEST',customer_name='Customer',total=0)
        if ok: messagebox.showinfo('Test SMS','MSG91 accepted the test SMS.\n\n'+resp)
        else: messagebox.showerror('Test SMS','MSG91 did not accept the request.\n\n'+resp)

    def account(self):
        self.title('My Account','Change your name, email address and password. Each email ID belongs to one account.')
        box=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#e2e8f0',padx=22,pady=22); box.pack(fill='x')
        name=tk.StringVar(value=self.user['name']); email=tk.StringVar(value=self.user['email'])
        tk.Label(box,text='Full name',bg='white',font=('Segoe UI',10,'bold')).grid(row=0,column=0,sticky='w',pady=8)
        tk.Entry(box,textvariable=name,width=48,bd=1,relief='solid').grid(row=0,column=1,padx=15,ipady=6,sticky='w')
        tk.Label(box,text='Email address',bg='white',font=('Segoe UI',10,'bold')).grid(row=1,column=0,sticky='w',pady=8)
        tk.Entry(box,textvariable=email,width=48,bd=1,relief='solid').grid(row=1,column=1,padx=15,ipady=6,sticky='w')
        tk.Label(box,text='Current password',bg='white',font=('Segoe UI',10,'bold')).grid(row=2,column=0,sticky='w',pady=8)
        current=tk.Entry(box,show='•',width=48,bd=1,relief='solid'); current.grid(row=2,column=1,padx=15,ipady=6,sticky='w')
        tk.Label(box,text='New password',bg='white',font=('Segoe UI',10,'bold')).grid(row=3,column=0,sticky='w',pady=8)
        newpw=tk.Entry(box,show='•',width=48,bd=1,relief='solid'); newpw.grid(row=3,column=1,padx=15,ipady=6,sticky='w')
        tk.Label(box,text='Confirm new password',bg='white',font=('Segoe UI',10,'bold')).grid(row=4,column=0,sticky='w',pady=8)
        confirm=tk.Entry(box,show='•',width=48,bd=1,relief='solid'); confirm.grid(row=4,column=1,padx=15,ipady=6,sticky='w')
        tk.Label(box,text=f'Current role: {self.user["role"]}',bg='white',fg='#64748b').grid(row=5,column=1,sticky='w',padx=15,pady=(5,12))
        def save():
            nm=name.get().strip(); em=email.get().strip().lower(); cp=current.get(); np=newpw.get(); cf=confirm.get()
            if not nm or not valid_email(em): messagebox.showerror('Validation','Enter a valid name and email.'); return
            row=self.db.one('SELECT * FROM users WHERE id=?',(self.user['id'],))
            if not row or not verify_pw(cp,row['salt'],row['password_hash']): messagebox.showerror('Validation','Current password is incorrect.'); return
            other=self.db.one('SELECT id FROM users WHERE lower(email)=? AND id<>?',(em,self.user['id']))
            if other: messagebox.showerror('Validation','That email is already used by another account.'); return
            salt,d=row['salt'],row['password_hash']
            if np or cf:
                if not strong_pw(np): messagebox.showerror('Validation','New password needs 8+ characters with uppercase, lowercase, number and special character.'); return
                if np!=cf: messagebox.showerror('Validation','New passwords do not match.'); return
                salt,d=hash_pw(np)
            self.db.conn.execute('UPDATE users SET name=?,email=?,salt=?,password_hash=?,must_change=0 WHERE id=?',(nm,em,salt,d,self.user['id'])); self.db.conn.commit()
            self.user=dict(self.db.one('SELECT * FROM users WHERE id=?',(self.user['id'],))); self.company=self.sms.s.get('company_name',self.company)
            messagebox.showinfo('Account updated','Your account details have been updated.'); self.dashboard()
        self.button(box,'Save Account Changes',save,kind='success').grid(row=6,column=1,sticky='w',padx=15,pady=12)

    def after_login_change_credentials(self):
        win=tk.Toplevel(self.root); win.title('Change Admin Credentials'); win.geometry('560x430'); win.transient(self.root); win.grab_set()
        frm=tk.Frame(win,bg='white',padx=30,pady=25); frm.pack(fill='both',expand=True)
        tk.Label(frm,text='First-time Admin setup',bg='white',fg='#0f172a',font=('Segoe UI',18,'bold')).pack(anchor='w')
        tk.Label(frm,text='For security, change the default Admin password before using the shop.',bg='white',fg='#64748b',wraplength=470,justify='left').pack(anchor='w',pady=(5,18))
        old=tk.Entry(frm,show='•',width=42); new=tk.Entry(frm,show='•',width=42); cf=tk.Entry(frm,show='•',width=42)
        for label,e in [('Current password',old),('New password',new),('Confirm new password',cf)]:
            tk.Label(frm,text=label,bg='white').pack(anchor='w',pady=(8,3)); e.pack(fill='x',ipady=6)
        def save():
            row=self.db.one('SELECT * FROM users WHERE id=?',(self.user['id'],))
            if not verify_pw(old.get(),row['salt'],row['password_hash']): messagebox.showerror('Validation','Current password is incorrect.',parent=win); return
            if not strong_pw(new.get()): messagebox.showerror('Validation','Password needs 8+ characters with uppercase, lowercase, number and special character.',parent=win); return
            if new.get()!=cf.get(): messagebox.showerror('Validation','Passwords do not match.',parent=win); return
            salt,d=hash_pw(new.get()); self.db.conn.execute('UPDATE users SET salt=?,password_hash=?,must_change=0 WHERE id=?',(salt,d,self.user['id'])); self.db.conn.commit(); self.user=dict(self.db.one('SELECT * FROM users WHERE id=?',(self.user['id'],))); win.destroy(); messagebox.showinfo('Security','Admin password changed successfully.')
        self.button(frm,'Save New Password',save,kind='success').pack(anchor='w',pady=20)
        win.protocol('WM_DELETE_WINDOW',lambda:(win.destroy()))

    def run(self):self.root.mainloop()

if __name__=='__main__':App().run()
