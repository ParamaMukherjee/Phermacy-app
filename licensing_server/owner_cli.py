"""Owner helper: list pending UTRs and approve/renew a licence."""
import argparse, json, os, urllib.request
p=argparse.ArgumentParser(); p.add_argument('--url',required=True); p.add_argument('--key',default=os.getenv('MEDIBILL_OWNER_KEY',''))
sub=p.add_subparsers(dest='cmd',required=True); sub.add_parser('pending')
a=sub.add_parser('approve'); a.add_argument('license_id')
r=sub.add_parser('renew'); r.add_argument('license_id'); r.add_argument('--years',type=int,default=1)
a=p.parse_args()
if not a.key: raise SystemExit('Set MEDIBILL_OWNER_KEY or pass --key.')
path='/owner/payments' if a.cmd=='pending' else ('/owner/approve-lifetime' if a.cmd=='approve' else '/owner/renew-amc')
data=None if a.cmd=='pending' else json.dumps({'license_id':a.license_id,'years':getattr(a,'years',1)}).encode()
req=urllib.request.Request(a.url.rstrip('/')+path,data=data,headers={'X-Owner-Key':a.key,'Content-Type':'application/json'},method='GET' if data is None else 'POST')
with urllib.request.urlopen(req,timeout=15) as r: print(json.dumps(json.loads(r.read()),indent=2))
