import os, tempfile, time, json, base64, importlib.util
from pathlib import Path
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives.serialization import Encoding, PublicFormat

# Verify the same canonical JSON signature format used by the API and desktop client.
key=Ed25519PrivateKey.generate(); pub=key.public_key().public_bytes(Encoding.Raw,PublicFormat.Raw)
payload={'license_id':'MBP-TEST','device_id':'device-test','state':'trial','trial_until':int(time.time())+30*86400,'amc_until':0,'issued_at':int(time.time()),'valid_until':int(time.time())+86400}
raw=json.dumps(payload,sort_keys=True,separators=(',',':')).encode(); sig=key.sign(raw)
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
Ed25519PublicKey.from_public_bytes(pub).verify(sig,raw)
try: Ed25519PublicKey.from_public_bytes(pub).verify(sig,b'bad')
except Exception: pass
else: raise AssertionError('tampered payload accepted')

# Verify the actual desktop verifier accepts a valid envelope and rejects tampering.
import license_client
license_client.PUBLIC_KEY_B64=base64.b64encode(pub).decode()
envelope={'payload':payload,'signature':base64.b64encode(sig).decode()}
token=base64.urlsafe_b64encode(json.dumps(envelope,separators=(',',':')).encode()).decode().rstrip('=')
assert license_client.verify_token(token)==payload
bad=dict(envelope); bad['payload']=dict(payload, state='suspended')
bad_token=base64.urlsafe_b64encode(json.dumps(bad,separators=(',',':')).encode()).decode().rstrip('=')
try: license_client.verify_token(bad_token)
except license_client.LicenseError: pass
else: raise AssertionError('desktop verifier accepted tampered token')

# Exercise API endpoints against a temporary database and generated signing key.
os.environ['MEDIBILL_LICENSE_DB']=str(Path(tempfile.mkdtemp())/'test.db')
keyfile=Path(tempfile.mkdtemp())/'private.pem'
from cryptography.hazmat.primitives.serialization import PrivateFormat, NoEncryption
keyfile.write_bytes(key.private_bytes(Encoding.PEM,PrivateFormat.PKCS8,NoEncryption()))
os.environ['MEDIBILL_SIGNING_KEY']=str(keyfile); os.environ['MEDIBILL_OWNER_KEY']='test-owner-key'
from fastapi.testclient import TestClient
from licensing_server.app import app
client=TestClient(app)
r=client.post('/v1/activate',json={'device_id':'a'*64}); assert r.status_code==200, r.text
body=r.json(); assert body['state']=='trial' and body['license_id'].startswith('MBP-')
r=client.post('/v1/status',json={'device_id':'a'*64,'license_id':body['license_id']}); assert r.status_code==200
r=client.post('/v1/payment-request',json={'device_id':'a'*64,'license_id':body['license_id'],'utr':'123456789012','contact':'9000000000'}); assert r.status_code==200
r=client.get('/owner/payments',headers={'X-Owner-Key':'test-owner-key'}); assert r.status_code==200 and len(r.json())==1
r=client.post('/owner/approve-lifetime',headers={'X-Owner-Key':'test-owner-key'},json={'license_id':body['license_id'],'years':1}); assert r.status_code==200 and r.json()['state']=='lifetime'
r=client.post('/owner/renew-amc',headers={'X-Owner-Key':'test-owner-key'},json={'license_id':body['license_id'],'years':1}); assert r.status_code==200
assert client.get('/owner/payments',headers={'X-Owner-Key':'wrong'}).status_code==403
print('Licensing API/signature/trial/payment/owner approval/AMC renewal tests: PASS')
