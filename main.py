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

def amount_words(n):
    ones=['Zero','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen']
    tens=['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety']
    def under1000(x):
        parts=[]
        if x>=100: parts.append(ones[x//100]+' Hundred'); x%=100
        if x>=20: parts.append(tens[x//10]); x%=10
        if x: parts.append(ones[x])
        return ' '.join(parts)
    x=int(round(float(n)));
    if x==0:return 'Zero Only'
    parts=[]
    if x>=10000000: parts.append(under1000(x//10000000)+' Crore'); x%=10000000
    if x>=100000: parts.append(under1000(x//100000)+' Lakh'); x%=100000
    if x>=1000: parts.append(under1000(x//1000)+' Thousand'); x%=1000
    if x: parts.append(under1000(x))
    return ' '.join(parts)+' Only'

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
        CREATE TABLE IF NOT EXISTS medicines(id INTEGER PRIMARY KEY, name TEXT NOT NULL, batch TEXT DEFAULT '', expiry TEXT DEFAULT '', price REAL DEFAULT 0, mrp REAL DEFAULT 0, cost REAL DEFAULT 0, stock INTEGER DEFAULT 0, gst REAL DEFAULT 0, supplier TEXT DEFAULT '', manufacturer TEXT DEFAULT '', pack TEXT DEFAULT '', UNIQUE(name,batch));
        CREATE TABLE IF NOT EXISTS bills(id INTEGER PRIMARY KEY, invoice TEXT UNIQUE NOT NULL, bill_date TEXT NOT NULL, bill_time TEXT NOT NULL, customer_id INTEGER, customer_name TEXT, customer_phone TEXT, customer_address TEXT DEFAULT '', doctor_name TEXT DEFAULT '', doctor_reg TEXT DEFAULT '', payment TEXT, subtotal REAL, gst REAL, total REAL, discount REAL DEFAULT 0, round_off REAL DEFAULT 0, sms_status TEXT DEFAULT 'Not sent', created_by TEXT, FOREIGN KEY(customer_id) REFERENCES customers(id));
        CREATE TABLE IF NOT EXISTS bill_items(id INTEGER PRIMARY KEY, bill_id INTEGER NOT NULL, medicine_id INTEGER, name TEXT, batch TEXT, expiry TEXT DEFAULT '', mrp REAL DEFAULT 0, manufacturer TEXT DEFAULT '', pack TEXT DEFAULT '', qty INTEGER, price REAL, gst REAL, discount REAL DEFAULT 0, amount REAL, FOREIGN KEY(bill_id) REFERENCES bills(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS sms_outbox(id INTEGER PRIMARY KEY, bill_id INTEGER, phone TEXT, message TEXT, status TEXT, provider_response TEXT DEFAULT '', created_at TEXT, FOREIGN KEY(bill_id) REFERENCES bills(id));
        ''')
        # Migrate older databases created before the credential-change feature.
        cols={r['name'] for r in c.execute('PRAGMA table_info(users)').fetchall()}
        if 'must_change' not in cols:
            c.execute('ALTER TABLE users ADD COLUMN must_change INTEGER DEFAULT 0')
        mcols={r['name'] for r in c.execute('PRAGMA table_info(medicines)').fetchall()}
        for col,typ in [('mrp','REAL DEFAULT 0'),('manufacturer',"TEXT DEFAULT ''"),('pack',"TEXT DEFAULT ''")]:
            if col not in mcols: c.execute(f'ALTER TABLE medicines ADD COLUMN {col} {typ}')
        bcols={r['name'] for r in c.execute('PRAGMA table_info(bills)').fetchall()}
        for col,typ in [('customer_address',"TEXT DEFAULT ''"),('doctor_name',"TEXT DEFAULT ''"),('doctor_reg',"TEXT DEFAULT ''"),('discount','REAL DEFAULT 0'),('round_off','REAL DEFAULT 0')]:
            if col not in bcols: c.execute(f'ALTER TABLE bills ADD COLUMN {col} {typ}')
        icols={r['name'] for r in c.execute('PRAGMA table_info(bill_items)').fetchall()}
        for col,typ in [('expiry',"TEXT DEFAULT ''"),('mrp','REAL DEFAULT 0'),('manufacturer',"TEXT DEFAULT ''"),('pack',"TEXT DEFAULT ''"),('discount','REAL DEFAULT 0')]:
            if col not in icols: c.execute(f'ALTER TABLE bill_items ADD COLUMN {col} {typ}')
        if not c.execute('SELECT 1 FROM users LIMIT 1').fetchone():
            for name,email,pw,role in [('Administrator','admin@medibillwb.in','Admin@1234','Admin'),('Store Manager','store@medibillwb.in','Demo@1234','Store Manager'),('Cashier','cashier@medibillwb.in','Demo@1234','Cashier'),('Accounts','accounts@medibillwb.in','Demo@1234','Accounts')]:
                salt,d=hash_pw(pw); c.execute('INSERT INTO users(name,email,salt,password_hash,role,verified,active,must_change,created_at) VALUES(?,?,?,?,?,?,?,?,?)',(name,email,salt,d,role,1,1,1 if email=='admin@medibillwb.in' else 0,datetime.now().isoformat(timespec='seconds')))
        if not c.execute('SELECT 1 FROM medicines LIMIT 1').fetchone():
            meds=[('Paracetamol 500mg','PCT001','2027-08-31',25,25,12,80,5,'Demo Supplier','Demo Pharma','10 tablets'),('Azithromycin 500mg','AZI001','2027-03-31',85,85,48,25,5,'Demo Supplier','Demo Pharma','6 tablets'),('ORS Lemon','ORS001','2028-01-31',22,22,13,50,5,'Demo Supplier','Demo Pharma','21 g'),('Pantoprazole 40mg','PAN001','2027-11-30',60,60,30,35,12,'Demo Supplier','Demo Pharma','10 tablets'),('Vitamin C 500mg','VIT001','2026-11-30',35,35,20,7,12,'Demo Supplier','Demo Pharma','10 tablets'),('Cetirizine 10mg','CET001','2026-10-15',18,18,9,15,5,'Demo Supplier','Demo Pharma','10 tablets')]
            c.executemany('INSERT INTO medicines(name,batch,expiry,price,mrp,cost,stock,gst,supplier,manufacturer,pack) VALUES(?,?,?,?,?,?,?,?,?,?,?)',meds)
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
        self.last_bill=None; self.root=tk.Tk(); self.root.withdraw(); self.root.title(APP_NAME); self.root.geometry('1180x680'); self.root.minsize(860,580); self.root.protocol('WM_DELETE_WINDOW',self.close)
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
        self.clear(); self.root.geometry('1180x680'); self.root.minsize(860,580); outer=tk.Frame(self.root,bg='#f8fafc'); outer.pack(fill='both',expand=True)
        side=tk.Frame(outer,bg='#0f172a',width=205); side.pack(side='left',fill='y'); side.pack_propagate(False)
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
        self.title('Dashboard','Live pharmacy overview. Filter the status list by date, status or medicine name.')
        today=datetime.now().strftime('%Y-%m-%d')
        self.dash_date=getattr(self,'dash_date',tk.StringVar(value=today))
        self.dash_status=getattr(self,'dash_status',tk.StringVar(value='All'))
        self.dash_search=getattr(self,'dash_search',tk.StringVar())

        filters=tk.Frame(self.content,bg='#eef2ff',highlightthickness=1,highlightbackground='#c7d2fe',padx=12,pady=10);filters.pack(fill='x',pady=(0,12))
        tk.Label(filters,text='View date',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).pack(side='left',padx=(0,5))
        self.date_entry(filters,self.dash_date,11)
        tk.Label(filters,text='Status',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).pack(side='left',padx=(4,5))
        status_box=ttk.Combobox(filters,textvariable=self.dash_status,values=['All','Expired','Near Expiry','Low Stock','In Stock'],state='readonly',width=14);status_box.pack(side='left',padx=(0,10))
        tk.Label(filters,text='Medicine',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).pack(side='left',padx=(0,5))
        search=tk.Entry(filters,textvariable=self.dash_search,width=28,bd=1,relief='solid');search.pack(side='left',ipady=5)
        self.button(filters,'Reset Filters',lambda:self.reset_dashboard_filters(),kind='dark').pack(side='left',padx=8)

        stats=tk.Frame(self.content,bg='#f8fafc'); stats.pack(fill='x')
        selected=self.dash_date.get().strip() or today
        vals=[('Sales',self.db.one('SELECT COALESCE(SUM(total),0) x FROM bills WHERE bill_date=?',(selected,))['x'],'#2563eb'),('Bills',self.db.one('SELECT COUNT(*) x FROM bills WHERE bill_date=?',(selected,))['x'],'#7c3aed'),('Customers',self.db.one('SELECT COUNT(*) x FROM customers')['x'],'#0891b2'),('Medicines',self.db.one('SELECT COUNT(*) x FROM medicines')['x'],'#16a34a')]
        for i,(label,val,accent) in enumerate(vals):
            c=tk.Frame(stats,bg='white',highlightthickness=1,highlightbackground='#e2e8f0'); c.grid(row=0,column=i,sticky='ew',padx=5); stats.columnconfigure(i,weight=1)
            tk.Label(c,text=f"{label} — {fmt_date(selected)}",bg='white',fg='#64748b',font=('Segoe UI',9)).pack(anchor='w',padx=15,pady=(12,2))
            tk.Label(c,text=f'₹{money(val):,.2f}' if label=='Sales' else str(val),bg='white',fg=accent,font=('Segoe UI',19,'bold')).pack(anchor='w',padx=15,pady=(0,11))

        tk.Label(self.content,text='Medicine Status',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',16,'bold')).pack(anchor='w',pady=(18,8))
        table=tk.Frame(self.content,bg='white'); table.pack(fill='both',expand=True)
        tv=ttk.Treeview(table,columns=('name','batch','stock','expiry','status'),show='headings');tv.pack(side='left',fill='both',expand=True)
        sb=ttk.Scrollbar(table,orient='vertical',command=tv.yview);sb.pack(side='right',fill='y');tv.configure(yscrollcommand=sb.set)
        for c,h in [('name','Medicine'),('batch','Batch'),('stock','Stock'),('expiry','Expiry'),('status','Status')]: tv.heading(c,text=h)
        tv.column('name',width=280,minwidth=150,stretch=True);tv.column('batch',width=120,minwidth=90,stretch=True);tv.column('stock',width=90,minwidth=70,stretch=False);tv.column('expiry',width=120,minwidth=100,stretch=False);tv.column('status',width=130,minwidth=110,stretch=False)
        tv.tag_configure('near',foreground='#dc2626');tv.tag_configure('expired',foreground='#991b1b');tv.tag_configure('low',foreground='#d97706');tv.tag_configure('instock',foreground='#15803d')
        q=self.dash_search.get().strip().lower(); selected_status=self.dash_status.get()
        rows=[]
        for r in self.db.q('SELECT name,batch,stock,expiry FROM medicines'):
            status,prio=expiry_status(r['expiry'],r['stock'])
            if selected_status!='All' and status!=selected_status: continue
            if q and q not in (r['name'] or '').lower() and q not in (r['batch'] or '').lower(): continue
            rows.append((prio,r))
        for _,r in sorted(rows,key=lambda x:(x[0],parse_date(x[1]['expiry']) if re.fullmatch(r'\d{4}-\d{2}-\d{2}',x[1]['expiry'] or '') else date.max,x[1]['name'].lower())):
            status,_=expiry_status(r['expiry'],r['stock']);tag={'Near Expiry':'near','Expired':'expired','Low Stock':'low','In Stock':'instock'}[status]
            tv.insert('', 'end',values=(r['name'],r['batch'],r['stock'],fmt_date(r['expiry']),status),tags=(tag,))
        for v in (self.dash_date,self.dash_status,self.dash_search):
            v.trace_add('write',lambda *_: self.dashboard_home())

    def reset_dashboard_filters(self):
        self.dash_date.set(datetime.now().strftime('%Y-%m-%d'));self.dash_status.set('All');self.dash_search.set('')

    def date_entry(self,parent,var,width=14):
        e=tk.Entry(parent,textvariable=var,font=('Segoe UI',10),bd=1,relief='solid',width=width); e.pack(side='left',ipady=5)
        self.button(parent,'📅',lambda:DatePicker(self.root,var.get(),lambda v:var.set(v)),kind='dark',width=3).pack(side='left',padx=(3,10)); return e
    def billing(self):
        self.title('Billing','Add medicines, adjust quantities, generate the bill, and send SMS only when you choose.')
        self.b_name=tk.StringVar(); self.b_phone=tk.StringVar(); self.b_address=tk.StringVar(); self.b_doctor=tk.StringVar(); self.b_doctor_reg=tk.StringVar(); self.b_payment=tk.StringVar(value='Cash'); self.b_search=tk.StringVar(); self.search_after=None
        info=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#e2e8f0',padx=15,pady=12); info.pack(fill='x')
        fields=[('Customer name',self.b_name,27),('Mobile number',self.b_phone,20),('Customer address',self.b_address,30),('Doctor name',self.b_doctor,22),('Doctor Reg. No.',self.b_doctor_reg,18)]
        for i,(lab,var,w) in enumerate(fields):
            tk.Label(info,text=lab,bg='white',fg='#475569').grid(row=0,column=i,sticky='w',padx=5); tk.Entry(info,textvariable=var,width=w,bd=1,relief='solid').grid(row=1,column=i,padx=5,ipady=6,sticky='ew')
        tk.Label(info,text='Payment method',bg='white',fg='#475569').grid(row=0,column=4,sticky='w',padx=5); ttk.Combobox(info,textvariable=self.b_payment,values=['Cash','UPI','Card','Credit'],state='readonly',width=14).grid(row=1,column=4,padx=5,sticky='w')
        tk.Label(self.content,text='Find medicine',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(18,6)); sf=tk.Frame(self.content,bg='#f8fafc'); sf.pack(fill='x'); search=tk.Entry(sf,textvariable=self.b_search,font=('Segoe UI',11),bd=1,relief='solid'); search.pack(side='left',fill='x',expand=True,ipady=7); tk.Label(sf,text='  Type to search instantly',bg='#f8fafc',fg='#64748b').pack(side='left')
        results_box=tk.Frame(self.content,bg='#f8fafc'); results_box.pack(fill='x',pady=8)
        self.bill_results=ttk.Treeview(results_box,columns=('id','name','batch','expiry','mrp','price','stock'),show='headings',height=5); self.bill_results.pack(side='left',fill='both',expand=True)
        rsb=ttk.Scrollbar(results_box,orient='vertical',command=self.bill_results.yview); rsb.pack(side='right',fill='y'); self.bill_results.configure(yscrollcommand=rsb.set)
        for c,h in [('id','ID'),('name','Medicine'),('batch','Batch'),('expiry','Expiry'),('mrp','MRP'),('price','Selling Price'),('stock','Available')]: self.bill_results.heading(c,text=h)
        self.bill_results.bind('<Double-1>',lambda e:self.add_selected()); self.bill_results.bind('<Return>',lambda e:self.add_selected()); self.bill_search()
        tk.Label(self.content,text='Bill items',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(10,6))
        cartbox=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#e2e8f0'); cartbox.pack(fill='x',expand=False)
        cart_table=tk.Frame(cartbox,bg='white'); cart_table.pack(fill='x',expand=False)
        self.cart_tv=ttk.Treeview(cart_table,columns=('name','batch','qty','mrp','price','gst','discount','amount'),show='headings',height=5); self.cart_tv.pack(side='left',fill='both',expand=True)
        cart_sb=ttk.Scrollbar(cart_table,orient='vertical',command=self.cart_tv.yview); cart_sb.pack(side='right',fill='y'); self.cart_tv.configure(yscrollcommand=cart_sb.set)
        
        for c,h in [('name','Medicine'),('batch','Batch'),('qty','Quantity'),('mrp','MRP'),('price','Selling Price'),('gst','GST %'),('discount','Disc.'),('amount','Amount')]: self.cart_tv.heading(c,text=h)
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
        for r in self.db.q('SELECT id,name,batch,expiry,mrp,price,stock,manufacturer,pack FROM medicines WHERE name LIKE ? OR batch LIKE ? ORDER BY name LIMIT 50',(q,q)): self.bill_results.insert('', 'end',values=tuple(r))
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
        self.cart.append({'id':r['id'],'name':r['name'],'batch':r['batch'],'expiry':r['expiry'],'mrp':float(r['mrp'] or r['price']),'manufacturer':r['manufacturer'],'pack':r['pack'],'qty':1,'price':float(r['price']),'gst':float(r['gst']),'discount':0.0});self.refresh_cart()
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
        self.cart.clear();self.b_name.set('');self.b_phone.set('');self.b_address.set('');self.b_doctor.set('');self.b_doctor_reg.set('');self.refresh_cart();self.last_invoice=None;self.last_bill_id=None;self.sms_btn.configure(state='disabled');self.pdf_btn.configure(state='disabled')
    def refresh_cart(self):
        if not hasattr(self,'cart_tv'):return
        for x in self.cart_tv.get_children():self.cart_tv.delete(x)
        total=0
        for i,it in enumerate(self.cart):
            amt=money(it['qty']*it['price']*(1+it['gst']/100));total+=amt;self.cart_tv.insert('', 'end',iid=str(i),values=(it['name'],it['batch'],it['qty'],f'₹{it["mrp"]:.2f}',f'₹{it["price"]:.2f}',it['gst'],f'₹{it.get("discount",0):.2f}',f'₹{amt:.2f}'))
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
            cur.execute('INSERT INTO bills(invoice,bill_date,bill_time,customer_id,customer_name,customer_phone,customer_address,doctor_name,doctor_reg,payment,subtotal,gst,total,discount,round_off,sms_status,created_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(invoice,now.strftime('%Y-%m-%d'),now.strftime('%H:%M:%S'),cid,name,phone,self.b_address.get().strip(),self.b_doctor.get().strip(),self.b_doctor_reg.get().strip(),self.b_payment.get(),money(subtotal),money(gsttotal),money(total),0,0,'Not sent' if phone else 'No mobile',self.user['email']));bid=cur.lastrowid
            for it in self.cart:
                base=money(it['qty']*it['price']);tax=money(base*it['gst']/100);cur.execute('INSERT INTO bill_items(bill_id,medicine_id,name,batch,expiry,mrp,manufacturer,pack,qty,price,gst,discount,amount) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(bid,it['id'],it['name'],it['batch'],it.get('expiry',''),it.get('mrp',it['price']),it.get('manufacturer',''),it.get('pack',''),it['qty'],it['price'],it['gst'],it.get('discount',0),money(base+tax)));cur.execute('UPDATE medicines SET stock=stock-? WHERE id=?',(it['qty'],it['id']))
            if cid:cur.execute('UPDATE customers SET total_purchase=total_purchase+? WHERE id=?',(money(total),cid))
            if phone:cur.execute('INSERT INTO sms_outbox(bill_id,phone,message,status,created_at) VALUES(?,?,?,?,?)',(bid,phone,self.sms_message(invoice,name,total),'Queued',now.isoformat(timespec='seconds')))
            self.db.conn.commit()
        except Exception as e:self.db.conn.rollback();messagebox.showerror('Could not generate bill',str(e));return
        self.last_invoice=invoice;self.last_bill_id=bid;self.last_bill={'invoice':invoice,'date':now.strftime('%Y-%m-%d'),'time':now.strftime('%H:%M:%S'),'name':name,'phone':phone,'address':self.b_address.get().strip(),'doctor':self.b_doctor.get().strip(),'doctor_reg':self.b_doctor_reg.get().strip(),'payment':self.b_payment.get(),'subtotal':money(subtotal),'gst':money(gsttotal),'total':money(total),'items':[dict(x) for x in self.cart]}
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
        if canvas is None: messagebox.showerror('PDF','PDF support is unavailable.'); return None
        stamp=f"{b['date']}_{b['time'].replace(':','-')}"; path=os.path.join(PDF_DIR,f"{b['invoice']}-{stamp}.pdf")
        c=canvas.Canvas(path,pagesize=A4); W,H=A4
        # Dynamic company watermark — always uses the current company name.
        c.saveState(); c.setFillColor(colors.HexColor('#e2e8f0')); c.translate(W/2,H/2); c.rotate(35); c.setFont('Helvetica-Bold',42); c.drawCentredString(0,0,self.company); c.restoreState()
        # Header inspired by the supplied pharmacy bill.
        y=H-42; c.setFillColor(colors.HexColor('#111827')); c.setFont('Helvetica-Bold',17); c.drawCentredString(W/2,y,self.company.upper())
        y-=17; c.setFont('Helvetica',8.5); info=self.sms.s; addr=info.get('company_address',''); phones=info.get('company_phone',''); reg=info.get('company_reg',''); gstin=info.get('company_gstin','')
        if addr: c.drawCentredString(W/2,y,addr[:110]); y-=12
        details='  '.join(x for x in [phones and 'Phone: '+phones, reg and 'Reg/Lic: '+reg, gstin and 'GSTIN: '+gstin] if x)
        if details: c.drawCentredString(W/2,y,details[:120]); y-=12
        c.setFont('Helvetica-Bold',12); c.drawString(45,y,'GST INVOICE'); c.setFont('Helvetica',8.5); c.drawRightString(W-45,y,f"Bill No: {b['invoice']}"); y-=14; c.drawRightString(W-45,y,f"Date: {fmt_date(b['date'])}   Time: {b['time']}")
        c.setStrokeColor(colors.HexColor('#334155')); c.line(45,y-8,W-45,y-8); y-=25
        c.setFont('Helvetica-Bold',8.5); c.drawString(45,y,'Customer:'); c.setFont('Helvetica',8.5); c.drawString(95,y,(b.get('name') or 'Walk-in Customer')[:38]); c.drawString(330,y,'Mobile: '+(b.get('phone') or '-')); y-=14
        if b.get('address'): c.drawString(45,y,'Address: '+b['address'][:90]); y-=14
        if b.get('doctor'): c.drawString(45,y,'Doctor: '+b['doctor'][:60] + (('   Reg: '+b.get('doctor_reg','')[:25]) if b.get('doctor_reg') else '')); y-=14
        # Table matching the supplied bill's useful fields.
        cols=[('Qty',45,25),('Pack',70,30),('Description',100,130),('Mfg',230,42),('Batch',272,55),('Exp',327,43),('MRP',370,45),('SGST%',415,35),('CGST%',450,35),('Disc',485,35),('Amount',520,50)]
        c.setFillColor(colors.HexColor('#111827')); c.setFont('Helvetica-Bold',7.3)
        for h,x,w in cols: c.drawString(x,y,h)
        y-=8; c.setStrokeColor(colors.HexColor('#64748b')); c.line(45,y,W-45,y); y-=14; c.setFont('Helvetica',7.2)
        for it in b['items']:
            if y<115:
                c.showPage(); y=H-50; c.setFont('Helvetica',7.2)
            gst=float(it.get('gst',0)); half=gst/2; disc=float(it.get('discount',0)); amount=money(it['qty']*it['price']*(1+gst/100)-disc)
            vals=[str(it['qty']), '1', str(it['name'])[:27], str(it.get('manufacturer',''))[:8], str(it.get('batch',''))[:10], fmt_date(it.get('expiry',''))[:8], f"{float(it.get('mrp',it['price'])):.2f}", f"{half:.1f}", f"{half:.1f}", f"{disc:.2f}", f"{amount:.2f}"]
            for (h,x,w),v in zip(cols,vals): c.drawString(x,y,v)
            y-=15
        y-=5; c.line(45,y,W-45,y); y-=18
        subtotal=float(b.get('subtotal',0)); gsttotal=float(b.get('gst',0)); total=float(b.get('total',0)); discount=float(b.get('discount',0)); roundoff=float(b.get('round_off',0))
        c.setFont('Helvetica',8.5); c.drawString(45,y,'Amount in words: '+amount_words(total)[:105]); c.drawRightString(W-125,y,'Sub Total:'); c.drawRightString(W-45,y,f"Rs. {subtotal:.2f}"); y-=16
        c.drawString(45,y,'Payment: '+str(b.get('payment','Cash'))); c.drawRightString(W-125,y,'Discount:'); c.drawRightString(W-45,y,f"Rs. {discount:.2f}"); y-=16
        c.drawRightString(W-125,y,'SGST:'); c.drawRightString(W-45,y,f"Rs. {gsttotal/2:.2f}"); y-=16
        c.drawRightString(W-125,y,'CGST:'); c.drawRightString(W-45,y,f"Rs. {gsttotal/2:.2f}"); y-=16
        c.drawRightString(W-125,y,'Round Off:'); c.drawRightString(W-45,y,f"Rs. {roundoff:.2f}"); y-=20
        c.setFont('Helvetica-Bold',12); c.drawRightString(W-125,y,'PARTY TOTAL:'); c.drawRightString(W-45,y,f"Rs. {total:.2f}"); y-=25
        c.setFont('Helvetica',8); c.drawString(45,y,'For '+self.company); c.drawRightString(W-45,y,'E. & O.E.'); y-=18; c.drawString(45,y,'Thank you for your purchase. Goods once sold are subject to applicable return policy.')
        c.save(); return path
    def open_file(self,path):
        try:
            if os.name=='nt':os.startfile(path)
            else:webbrowser.open('file://'+os.path.abspath(path))
        except Exception:pass
    def inventory(self):
        self.title('Inventory','Add medicines manually or import many medicines at once from Excel / CSV.')
        top=tk.Frame(self.content,bg='#f8fafc');top.pack(fill='x',pady=(0,10));self.button(top,'Bulk Import Excel / CSV',self.bulk_import,kind='primary').pack(side='left');tk.Label(top,text='Expected columns: Name, Batch, Expiry, MRP, Selling Price, Cost, Stock, GST %, Supplier, Manufacturer, Pack',bg='#f8fafc',fg='#64748b').pack(side='left',padx=12)
        form=tk.Frame(self.content,bg='white',highlightthickness=1,highlightbackground='#e2e8f0',padx=12,pady=10);form.pack(fill='x');vars=[tk.StringVar() for _ in range(11)];labels=['Name','Batch','Expiry','MRP','Selling price','Cost','Stock','GST %','Supplier','Manufacturer','Pack']
        for i,(lab,var) in enumerate(zip(labels,vars)):
            col=i%4;row=(i//4)*2;tk.Label(form,text=lab,bg='white',fg='#475569').grid(row=row,column=col,sticky='w',padx=5,pady=(0,3));cell=tk.Frame(form,bg='white');cell.grid(row=row+1,column=col,sticky='w',padx=5,pady=(0,5));
            if lab=='Expiry':self.date_entry(cell,var,15)
            else:tk.Entry(cell,textvariable=var,width=20,bd=1,relief='solid').pack(ipady=5)
        def add():
            try:
                name,batch,expiry,mrp,sprice,cost,stock,gst,supplier,manufacturer,pack=[v.get().strip() for v in vars]
                parse_date(expiry);vals=(name,batch,expiry,float(sprice or 0),float(mrp or sprice or 0),float(cost or 0),int(stock or 0),float(gst or 0),supplier,manufacturer,pack)
                self.db.conn.execute('INSERT INTO medicines(name,batch,expiry,price,mrp,cost,stock,gst,supplier,manufacturer,pack) VALUES(?,?,?,?,?,?,?,?,?,?,?)',vals);self.db.conn.commit();refresh();[v.set('') for v in vars]
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
            aliases={'name':['name','medicine','medicine name'],'batch':['batch','batch no'],'expiry':['expiry','expiry date'],'mrp':['mrp','maximum retail price'],'price':['selling price','price'],'cost':['cost','purchase price'],'stock':['stock','quantity','qty'],'gst':['gst','gst %','gst percent'],'supplier':['supplier','supplier name'],'manufacturer':['manufacturer','mfg','company'],'pack':['pack','packing','pack size']}
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
                    vals=(name,batch,expiry,float(get(rec,'price') or 0),float(get(rec,'mrp') or get(rec,'price') or 0),float(get(rec,'cost') or 0),int(float(get(rec,'stock') or 0)),float(get(rec,'gst') or 0),str(get(rec,'supplier') or ''),str(get(rec,'manufacturer') or ''),str(get(rec,'pack') or ''))
                    if not name:raise ValueError('missing Name')
                    old=self.db.one('SELECT id FROM medicines WHERE name=? AND batch=?',(name,batch))
                    if old:self.db.conn.execute('UPDATE medicines SET expiry=?,price=?,mrp=?,cost=?,stock=?,gst=?,supplier=?,manufacturer=?,pack=? WHERE id=?',(expiry,vals[3],vals[4],vals[5],vals[6],vals[7],vals[8],vals[9],vals[10],old['id']));updated+=1
                    else:self.db.conn.execute('INSERT INTO medicines(name,batch,expiry,price,mrp,cost,stock,gst,supplier,manufacturer,pack) VALUES(?,?,?,?,?,?,?,?,?,?,?)',vals);inserted+=1
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
        self.title('Sales','Search and filter bills instantly. Export the filtered bills to Excel or double-click a bill to open its PDF.')
        f=tk.Frame(self.content,bg='#eef2ff',highlightthickness=1,highlightbackground='#c7d2fe',padx=10,pady=9);f.pack(fill='x',pady=(0,8))
        self.sf=tk.StringVar();self.sfrom=tk.StringVar();self.sto=tk.StringVar();self.spay=tk.StringVar(value='All')
        f.columnconfigure(1,weight=1)
        tk.Label(f,text='Search',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).grid(row=0,column=0,sticky='w',padx=(0,4),pady=3)
        tk.Entry(f,textvariable=self.sf,width=24,bd=1,relief='solid').grid(row=0,column=1,sticky='ew',padx=(0,8),ipady=5,pady=3)
        tk.Label(f,text='From',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).grid(row=0,column=2,sticky='w',padx=4)
        cell=tk.Frame(f,bg='#eef2ff');cell.grid(row=0,column=3,sticky='w');self.date_entry(cell,self.sfrom,11)
        tk.Label(f,text='To',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).grid(row=0,column=4,sticky='w',padx=4)
        cell=tk.Frame(f,bg='#eef2ff');cell.grid(row=0,column=5,sticky='w');self.date_entry(cell,self.sto,11)
        tk.Label(f,text='Payment',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).grid(row=1,column=0,sticky='w',padx=(0,4),pady=3)
        ttk.Combobox(f,textvariable=self.spay,values=['All','Cash','UPI','Card','Credit'],state='readonly',width=12).grid(row=1,column=1,sticky='w',padx=(0,8),pady=3)
        self.button(f,'Export Excel',self.export_excel,kind='success').grid(row=1,column=2,columnspan=2,sticky='w',padx=4,pady=3)
        for v in [self.sf,self.sfrom,self.sto,self.spay]:v.trace_add('write',lambda *_:self.refresh_sales())
        tvbox=tk.Frame(self.content,bg='white');tvbox.pack(fill='both',expand=True)
        self.sales_tv=ttk.Treeview(tvbox,columns=('invoice','date','time','customer','phone','payment','subtotal','gst','total','sms'),show='headings');self.sales_tv.pack(side='left',fill='both',expand=True)
        vsb=ttk.Scrollbar(tvbox,orient='vertical',command=self.sales_tv.yview);vsb.pack(side='right',fill='y')
        hsb=ttk.Scrollbar(self.content,orient='horizontal',command=self.sales_tv.xview);hsb.pack(fill='x',pady=(0,6));self.sales_tv.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        headers=[('invoice','Invoice',190),('date','Date',100),('time','Time',85),('customer','Customer',170),('phone','Mobile',120),('payment','Payment',90),('subtotal','Taxable',95),('gst','GST',80),('total','Total',95),('sms','SMS',110)]
        for c,h,w in headers:self.sales_tv.heading(c,text=h);self.sales_tv.column(c,width=w,minwidth=70,stretch=False)
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
        invoice=self.sales_tv.item(s[0])['values'][0];r=self.db.one('SELECT * FROM bills WHERE invoice=?',(invoice,));items=[dict(x) for x in self.db.q('SELECT * FROM bill_items WHERE bill_id=?',(r['id'],))];b={'invoice':r['invoice'],'date':r['bill_date'],'time':r['bill_time'],'name':r['customer_name'],'phone':r['customer_phone'],'address':r['customer_address'] if 'customer_address' in r.keys() else '','doctor':r['doctor_name'] if 'doctor_name' in r.keys() else '','doctor_reg':r['doctor_reg'] if 'doctor_reg' in r.keys() else '','payment':r['payment'],'subtotal':r['subtotal'],'gst':r['gst'],'total':r['total'],'items':items};p=self.create_pdf(b)
        if p and messagebox.askyesno('PDF ready','Open the PDF now?'):self.open_file(p)
    def reports(self):
        self.title('Reports','Filter sales by date, payment method and customer/invoice, then review the filtered summary and top medicines.')
        f=tk.Frame(self.content,bg='#eef2ff',highlightthickness=1,highlightbackground='#c7d2fe',padx=10,pady=9);f.pack(fill='x',pady=(0,10))
        self.rf=tk.StringVar();self.rfrom=tk.StringVar();self.rto=tk.StringVar();self.rpay=tk.StringVar(value='All')
        f.columnconfigure(1,weight=1)
        tk.Label(f,text='Search',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).grid(row=0,column=0,sticky='w',padx=(0,4),pady=3)
        tk.Entry(f,textvariable=self.rf,width=22,bd=1,relief='solid').grid(row=0,column=1,sticky='ew',padx=(0,8),ipady=5,pady=3)
        tk.Label(f,text='From',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).grid(row=0,column=2,sticky='w',padx=4)
        cell=tk.Frame(f,bg='#eef2ff');cell.grid(row=0,column=3,sticky='w');self.date_entry(cell,self.rfrom,11)
        tk.Label(f,text='To',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).grid(row=0,column=4,sticky='w',padx=4)
        cell=tk.Frame(f,bg='#eef2ff');cell.grid(row=0,column=5,sticky='w');self.date_entry(cell,self.rto,11)
        tk.Label(f,text='Payment',bg='#eef2ff',fg='#334155',font=('Segoe UI',9,'bold')).grid(row=1,column=0,sticky='w',padx=(0,4),pady=3)
        ttk.Combobox(f,textvariable=self.rpay,values=['All','Cash','UPI','Card','Credit'],state='readonly',width=12).grid(row=1,column=1,sticky='w',padx=(0,8),pady=3)
        self.button(f,'Reset Filters',self.reset_reports,kind='dark').grid(row=1,column=2,columnspan=2,sticky='w',padx=4,pady=3)
        for v in [self.rf,self.rfrom,self.rto,self.rpay]:v.trace_add('write',lambda *_:self.refresh_reports())
        self.report_cards=tk.Frame(self.content,bg='#f8fafc');self.report_cards.pack(fill='x')
        tk.Label(self.content,text='Top Medicines by Quantity Sold',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',14,'bold')).pack(anchor='w',pady=(18,8))
        rbox=tk.Frame(self.content,bg='white');rbox.pack(fill='both',expand=True)
        self.report_tv=ttk.Treeview(rbox,columns=('medicine','qty','sales'),show='headings');self.report_tv.pack(side='left',fill='both',expand=True)
        r_sb=ttk.Scrollbar(rbox,orient='vertical',command=self.report_tv.yview);r_sb.pack(side='right',fill='y');self.report_tv.configure(yscrollcommand=r_sb.set)
        for c,h in [('medicine','Medicine'),('qty','Units Sold'),('sales','Sales Value')]:self.report_tv.heading(c,text=h)
        self.report_tv.column('medicine',width=320,minwidth=150,stretch=True);self.report_tv.column('qty',width=140,minwidth=90,stretch=False);self.report_tv.column('sales',width=160,minwidth=100,stretch=False)
        self.refresh_reports()

    def report_rows(self):
        where=[];args=[];q=self.rf.get().strip()
        if q:where.append('(b.invoice LIKE ? OR b.customer_name LIKE ? OR b.customer_phone LIKE ? OR i.name LIKE ?)');args += [f'%{q}%']*4
        if self.rfrom.get().strip():where.append('b.bill_date>=?');args.append(self.rfrom.get().strip())
        if self.rto.get().strip():where.append('b.bill_date<=?');args.append(self.rto.get().strip())
        if self.rpay.get()!='All':where.append('b.payment=?');args.append(self.rpay.get())
        clause=(' WHERE '+' AND '.join(where)) if where else ''
        return where,args,clause

    def refresh_reports(self):
        if not hasattr(self,'report_tv'):return
        where,args,clause=self.report_rows()
        for w in self.report_cards.winfo_children():w.destroy()
        total=self.db.one('SELECT COALESCE(SUM(b.total),0) x FROM bills b'+clause,args)['x']
        bills=self.db.one('SELECT COUNT(*) x FROM bills b'+clause,args)['x']
        avg=(float(total)/bills) if bills else 0
        for i,(label,val,accent) in enumerate([('Filtered Sales',total,'#2563eb'),('Bills',bills,'#7c3aed'),('Average Bill',avg,'#0891b2')]):
            lf=tk.LabelFrame(self.report_cards,text=label,bg='white',fg='#64748b',padx=25,pady=10);lf.grid(row=0,column=i,sticky='ew',padx=5);self.report_cards.columnconfigure(i,weight=1)
            tk.Label(lf,text=f'₹{money(val):,.2f}' if label!='Bills' else str(val),bg='white',fg=accent,font=('Segoe UI',18,'bold')).pack()
        for x in self.report_tv.get_children():self.report_tv.delete(x)
        top_sql='SELECT i.name,SUM(i.qty) qty,SUM(i.amount) sales FROM bill_items i JOIN bills b ON b.id=i.bill_id'
        # bill-level filters involving invoice/customer/date/payment can be applied directly to the joined query.
        if where:
            top_sql+=' WHERE '+' AND '.join(where)
        top_sql+=' GROUP BY i.name ORDER BY qty DESC LIMIT 50'
        # If search contains bill/customer fields, this query is still valid because b is joined.
        for r in self.db.q(top_sql,args):self.report_tv.insert('', 'end',values=(r['name'],r['qty'],f"₹{r['sales']:.2f}"))

    def reset_reports(self):
        self.rf.set('');self.rfrom.set('');self.rto.set('');self.rpay.set('All')

    def users(self):
        if self.user['role'] not in ['Admin','Store Manager']:return
        self.title('Users & Roles','Deactivate a worker to remove login access without losing their account. Rehire restores the same account; no separate activation button is needed.')
        bar_top=tk.Frame(self.content,bg='#eef2ff',highlightthickness=1,highlightbackground='#c7d2fe',padx=10,pady=8);bar_top.pack(fill='x')
        self.show_inactive=tk.BooleanVar(value=False)
        tk.Checkbutton(bar_top,text='Show deactivated accounts',variable=self.show_inactive,bg='#eef2ff',command=lambda:self.refresh_users(tv,role)).pack(side='left')
        tvbox=tk.Frame(self.content,bg='white');tvbox.pack(fill='both',expand=True,pady=8)
        tv=ttk.Treeview(tvbox,columns=('id','name','email','role','status'),show='headings');tv.pack(side='left',fill='both',expand=True)
        u_sb=ttk.Scrollbar(tvbox,orient='vertical',command=tv.yview);u_sb.pack(side='right',fill='y');tv.configure(yscrollcommand=u_sb.set)
        hsb=ttk.Scrollbar(self.content,orient='horizontal',command=tv.xview);hsb.pack(fill='x',pady=(0,6));tv.configure(xscrollcommand=hsb.set)
        for c,h in [('id','ID'),('name','Name'),('email','Email'),('role','Role'),('status','Status')]:tv.heading(c,text=h)
        for c,w in [('id',60),('name',180),('email',260),('role',150),('status',120)]:tv.column(c,width=w,minwidth=70,stretch=False)
        bar=tk.Frame(self.content,bg='#f8fafc');bar.pack(fill='x')
        role=tk.StringVar(value='Cashier');tk.Label(bar,text='Role',bg='#f8fafc').pack(side='left');ttk.Combobox(bar,textvariable=role,values=ROLES[1:],state='readonly',width=18).pack(side='left',padx=5)
        assign=self.button(bar,'Assign Role',lambda:self.assign_role(tv,role),kind='primary');assign.pack(side='left')
        if self.user['role']!='Store Manager':assign.configure(state='disabled',bg='#94a3b8',cursor='arrow')
        self.button(bar,'Reset Password',lambda:self.reset_user_password(tv),kind='warning').pack(side='left',padx=5)
        self.button(bar,'Deactivate Account',lambda:self.deactivate_user(tv),kind='danger').pack(side='left',padx=5)
        self.button(bar,'Rehire Account',lambda:self.rehire_user(tv),kind='success').pack(side='left',padx=5)
        self.refresh_users(tv,role)

    def refresh_users(self,tv,role=None):
        if not hasattr(self,'show_inactive'):return
        for x in tv.get_children():tv.delete(x)
        clause='' if self.show_inactive.get() else ' WHERE active=1'
        for r in self.db.q('SELECT id,name,email,role,active FROM users'+clause+' ORDER BY id'):
            tv.insert('', 'end',iid=str(r['id']),values=(r['id'],r['name'],r['email'],r['role'],'Active' if r['active'] else 'Deactivated'))

    def deactivate_user(self,tv):
        s=tv.selection()
        if not s:return
        uid=int(s[0]);r=self.db.one('SELECT * FROM users WHERE id=?',(uid,))
        if not r:return
        if uid==self.user['id']:messagebox.showwarning('Users','You cannot deactivate your own account.');return
        if r['role']=='Admin' and self.user['role']!='Admin':messagebox.showerror('Access','Only Admin can deactivate an Admin account.');return
        if not r['active']:return
        if not messagebox.askyesno('Deactivate account',f"Deactivate {r['name']}?\n\nThey will immediately lose login access. The account is retained so it can be rehired later."):return
        self.db.conn.execute('UPDATE users SET active=0 WHERE id=?',(uid,));self.db.conn.commit();self.refresh_users(tv)

    def rehire_user(self,tv):
        s=tv.selection()
        if not s:return
        uid=int(s[0]);r=self.db.one('SELECT * FROM users WHERE id=?',(uid,))
        if not r or r['active'] :messagebox.showinfo('Rehire','Select a deactivated account.');return
        if r['role']=='Admin' and self.user['role']!='Admin':messagebox.showerror('Access','Only Admin can rehire an Admin account.');return
        if not messagebox.askyesno('Rehire account',f"Rehire {r['name']}?\n\nThe existing email, role and password will be restored."):return
        self.db.conn.execute('UPDATE users SET active=1 WHERE id=?',(uid,));self.db.conn.commit();self.refresh_users(tv)

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
        fields=[('Shop address','company_address',1),('Phone number','company_phone',2),('Drug/Shop Reg. No.','company_reg',3),('GSTIN','company_gstin',4)]
        for lab,key,row in fields:
            tk.Label(box,text=lab,bg='white',font=('Segoe UI',10,'bold')).grid(row=row,column=0,sticky='w',pady=6)
            v=tk.StringVar(value=self.sms.s.get(key,'')); setattr(self,'_set_'+key,v)
            tk.Entry(box,textvariable=v,width=55,bd=1,relief='solid').grid(row=row,column=1,columnspan=2,padx=12,ipady=6,sticky='w')
        self.button(box,'Save Shop Details',self.save_shop_details,kind='success').grid(row=5,column=1,sticky='w',padx=12,pady=7)

        tk.Label(box,text='MSG91 SMS',bg='white',fg='#0f172a',font=('Segoe UI',14,'bold')).grid(row=6,column=0,columnspan=3,sticky='w',pady=(25,5))
        tk.Label(box,text='Auth Key',bg='white',font=('Segoe UI',10,'bold')).grid(row=7,column=0,sticky='w',pady=7)
        auth=tk.StringVar(value=self.sms.s.get('msg91_authkey',''))
        tk.Entry(box,textvariable=auth,show='•',width=55,bd=1,relief='solid').grid(row=7,column=1,columnspan=2,padx=12,ipady=6,sticky='w')
        tk.Label(box,text='Flow ID / Template ID',bg='white',font=('Segoe UI',10,'bold')).grid(row=8,column=0,sticky='w',pady=7)
        flow=tk.StringVar(value=self.sms.s.get('msg91_flow_id',''))
        tk.Entry(box,textvariable=flow,width=55,bd=1,relief='solid').grid(row=8,column=1,columnspan=2,padx=12,ipady=6,sticky='w')
        tk.Label(box,text='Sender ID / Header',bg='white',font=('Segoe UI',10,'bold')).grid(row=9,column=0,sticky='w',pady=7)
        sender=tk.StringVar(value=self.sms.s.get('msg91_sender_id',''))
        tk.Entry(box,textvariable=sender,width=55,bd=1,relief='solid').grid(row=9,column=1,columnspan=2,padx=12,ipady=6,sticky='w')
        tk.Label(box,text='MSG91 template variables expected by this app:  company, name, invoice, amount',bg='white',fg='#475569',wraplength=780).grid(row=10,column=0,columnspan=3,sticky='w',pady=(12,4))
        tk.Label(box,text='Your DLT-approved MSG91 template must contain these variables (for example ##name##, ##invoice##, ##amount##). The actual DLT text must match your approved template.',bg='white',fg='#64748b',wraplength=780,justify='left').grid(row=11,column=0,columnspan=3,sticky='w',pady=(0,12))
        self.button(box,'Save MSG91 Settings',lambda:self.save_sms(auth,flow,sender),kind='primary').grid(row=12,column=1,sticky='w',padx=12)
        test_frame=tk.Frame(box,bg='white');test_frame.grid(row=12,column=2,sticky='e',padx=5)
        test_phone=tk.StringVar()
        tk.Entry(test_frame,textvariable=test_phone,width=18,bd=1,relief='solid').pack(side='left',ipady=5)
        self.button(test_frame,'Send Test SMS',lambda:self.test_msg91(test_phone),kind='dark').pack(side='left',padx=5)

        tk.Label(self.content,text='PDF / Database',bg='#f8fafc',fg='#0f172a',font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(22,6))
        tk.Label(self.content,text=f'Bills PDF folder: {PDF_DIR}\nSQLite database: {DB_PATH}\nCompany name is printed on bills and used as the PDF watermark.',bg='#f8fafc',fg='#64748b',justify='left').pack(anchor='w')

    def save_shop_details(self):
        self.sms.s.update({
            'company_name':self.company,
            'company_address':self._set_company_address.get().strip(),
            'company_phone':self._set_company_phone.get().strip(),
            'company_reg':self._set_company_reg.get().strip(),
            'company_gstin':self._set_company_gstin.get().strip(),
        }); self.sms.save(); messagebox.showinfo('Shop details','Shop details saved. They will appear on future invoices.'); self.dashboard()

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
