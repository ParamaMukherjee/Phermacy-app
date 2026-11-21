"""MediBill Pro licensing API. Deploy privately; never distribute private_key.pem."""
import base64, json, os, sqlite3, time, uuid
from pathlib import Path
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel, Field
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives.serialization import load_pem_private_key

ROOT=Path(__file__).resolve().parent
DB=Path(os.getenv('MEDIBILL_LICENSE_DB', str(ROOT/'licenses.db')))
KEY=Path(os.getenv('MEDIBILL_SIGNING_KEY', str(ROOT/'private_key.pem')))
OWNER_KEY=os.getenv('MEDIBILL_OWNER_KEY','')
app=FastAPI(title='MediBill Pro Licensing API', version='1.0.0')

def conn():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row
    c.execute('''CREATE TABLE IF NOT EXISTS licenses(license_id TEXT PRIMARY KEY,device_id TEXT UNIQUE NOT NULL,state TEXT NOT NULL,trial_until INTEGER NOT NULL,amc_until INTEGER NOT NULL DEFAULT 0,created_at INTEGER NOT NULL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS payments(id INTEGER PRIMARY KEY,license_id TEXT NOT NULL,utr TEXT NOT NULL,contact TEXT DEFAULT '',status TEXT DEFAULT 'pending',created_at INTEGER NOT NULL,UNIQUE(license_id,utr))''')
    c.commit(); return c

def signer():
    if not KEY.exists(): raise HTTPException(503,'Licensing signing key is not configured.')
    return load_pem_private_key(KEY.read_bytes(),password=None)

def signed(row):
    now=int(time.time()); state=row['state']
    p={'license_id':row['license_id'],'device_id':row['device_id'],'state':state,'trial_until':row['trial_until'],'amc_until':row['amc_until'],'issued_at':now,'valid_until':now+86400}
    raw=json.dumps(p,sort_keys=True,separators=(',',':')).encode()
    sig=signer().sign(raw)
    envelope={'payload':p,'signature':base64.b64encode(sig).decode()}
    token=base64.urlsafe_b64encode(json.dumps(envelope,separators=(',',':')).encode()).decode().rstrip('=')
    return {'license_id':row['license_id'],'token':token,'state':state}

class Device(BaseModel): device_id:str=Field(min_length=16,max_length=128)
class Status(Device): license_id:str
class Payment(Device): license_id:str; utr:str=Field(min_length=6,max_length=80); contact:str=''
class AdminAction(BaseModel): license_id:str; years:int=1

@app.get('/health')
def health(): return {'ok':True,'service':'MediBill Pro Licensing API'}

@app.post('/v1/activate')
def activate(x:Device):
    c=conn(); row=c.execute('SELECT * FROM licenses WHERE device_id=?',(x.device_id,)).fetchone()
    if row is None:
        now=int(time.time()); lid='MBP-'+uuid.uuid4().hex[:12].upper()
        c.execute('INSERT INTO licenses(license_id,device_id,state,trial_until,created_at) VALUES(?,?,?,?,?)',(lid,x.device_id,'trial',now+30*86400,now)); c.commit()
        row=c.execute('SELECT * FROM licenses WHERE license_id=?',(lid,)).fetchone()
    c.close(); return signed(row)

@app.post('/v1/status')
def status(x:Status):
    c=conn(); row=c.execute('SELECT * FROM licenses WHERE license_id=? AND device_id=?',(x.license_id,x.device_id)).fetchone(); c.close()
    if not row: raise HTTPException(404,'Licence/device not found.')
    if row['state']=='suspended': raise HTTPException(403,'Licence suspended. Contact support.')
    return signed(row)

@app.post('/v1/payment-request')
def payment(x:Payment):
    c=conn(); row=c.execute('SELECT * FROM licenses WHERE license_id=? AND device_id=?',(x.license_id,x.device_id)).fetchone()
    if not row: c.close(); raise HTTPException(404,'Licence not found.')
    try: c.execute('INSERT INTO payments(license_id,utr,contact,created_at) VALUES(?,?,?,?)',(x.license_id,x.utr.strip(),x.contact.strip(),int(time.time()))); c.commit()
    except sqlite3.IntegrityError: c.close(); raise HTTPException(409,'This UTR has already been submitted for this licence.')
    c.close(); return {'received':True,'message':'Payment reference submitted. Activation occurs only after owner verifies the payment.'}

def require_owner(key):
    if not OWNER_KEY or key != OWNER_KEY: raise HTTPException(403,'Owner authentication failed.')

@app.get('/owner/payments')
def pending_payments(x_owner_key:str=Header(default='')):
    require_owner(x_owner_key); c=conn(); rows=[dict(r) for r in c.execute("SELECT * FROM payments WHERE status='pending' ORDER BY created_at")]; c.close(); return rows

@app.post('/owner/approve-lifetime')
def approve(x:AdminAction,x_owner_key:str=Header(default='')):
    require_owner(x_owner_key); c=conn(); row=c.execute('SELECT * FROM licenses WHERE license_id=?',(x.license_id,)).fetchone()
    if not row: c.close(); raise HTTPException(404,'Licence not found.')
    now=int(time.time()); until=now+365*86400
    c.execute("UPDATE licenses SET state='lifetime',amc_until=? WHERE license_id=?",(until,x.license_id))
    c.execute("UPDATE payments SET status='approved' WHERE license_id=? AND status='pending'",(x.license_id,)); c.commit()
    row=c.execute('SELECT * FROM licenses WHERE license_id=?',(x.license_id,)).fetchone(); c.close(); return signed(row)

@app.post('/owner/renew-amc')
def renew(x:AdminAction,x_owner_key:str=Header(default='')):
    require_owner(x_owner_key); c=conn(); row=c.execute("SELECT * FROM licenses WHERE license_id=? AND state='lifetime'",(x.license_id,)).fetchone()
    if not row: c.close(); raise HTTPException(404,'Lifetime licence not found.')
    base=max(int(row['amc_until']),int(time.time())); until=base+max(1,min(x.years,5))*365*86400
    c.execute('UPDATE licenses SET amc_until=? WHERE license_id=?',(until,x.license_id)); c.commit()
    row=c.execute('SELECT * FROM licenses WHERE license_id=?',(x.license_id,)).fetchone(); c.close(); return signed(row)
