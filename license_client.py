"""MediBill Pro signed online licence client.

Set MEDIBILL_LICENSE_API in the environment or license_config.json before building.
The server signing public key is injected into PUBLIC_KEY_B64 at release time.
"""
import base64, json, os, time, urllib.request, urllib.error
from pathlib import Path

PUBLIC_KEY_B64 = "REPLACE_WITH_SERVER_PUBLIC_KEY_B64"
APP_DIR = Path(os.path.join(os.path.expanduser('~'), 'MediBillPro'))
LICENSE_FILE = APP_DIR / 'license.json'
CONFIG_FILE = (Path(os.path.abspath(__import__('sys').executable)).with_name('license_config.json') if getattr(__import__('sys'), 'frozen', False) else Path(__file__).with_name('license_config.json'))

class LicenseError(Exception):
    pass

def api_url():
    value = os.environ.get('MEDIBILL_LICENSE_API', '').strip()
    if not value and CONFIG_FILE.exists():
        try: value = json.loads(CONFIG_FILE.read_text(encoding='utf-8')).get('api_url', '').strip()
        except Exception: pass
    return value.rstrip('/')

def device_id():
    import hashlib, platform, uuid
    raw = '|'.join((platform.node(), platform.system(), platform.machine(), str(uuid.getnode())))
    return hashlib.sha256(raw.encode()).hexdigest()

def _request(path, payload=None, timeout=8):
    url = api_url()
    if not url: raise LicenseError('Licensing service is not configured. Set MEDIBILL_LICENSE_API or license_config.json.')
    data = None if payload is None else json.dumps(payload).encode()
    req = urllib.request.Request(url + path, data=data, headers={'Content-Type':'application/json'}, method='GET' if payload is None else 'POST')
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response: return json.loads(response.read().decode())
    except urllib.error.HTTPError as e:
        try: detail = e.read().decode()
        except Exception: detail = str(e)
        raise LicenseError(f'Licensing server rejected the request ({e.code}): {detail[:300]}')
    except Exception as e: raise LicenseError(f'Could not contact licensing server: {e}')

def verify_token(token):
    if PUBLIC_KEY_B64.startswith('REPLACE_'): raise LicenseError('The release signing public key has not been configured.')
    try:
        from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
        envelope = json.loads(base64.urlsafe_b64decode(token + '=' * (-len(token) % 4)))
        payload = envelope['payload']; signature = base64.b64decode(envelope['signature'])
        key = Ed25519PublicKey.from_public_bytes(base64.b64decode(PUBLIC_KEY_B64))
        key.verify(signature, json.dumps(payload, sort_keys=True, separators=(',', ':')).encode())
        return payload
    except LicenseError: raise
    except Exception as e: raise LicenseError(f'Licence signature is invalid: {e}')

def check_or_activate():
    """Activate a 30-day trial on first run, then refresh a signed entitlement."""
    APP_DIR.mkdir(parents=True, exist_ok=True)
    did = device_id()
    cached = {}
    if LICENSE_FILE.exists():
        try: cached = json.loads(LICENSE_FILE.read_text(encoding='utf-8'))
        except Exception: cached = {}
    try:
        if not cached.get('activated'):
            result = _request('/v1/activate', {'device_id': did, 'app_version':'20.0'})
            cached['activated'] = True
            cached['license_id'] = result['license_id']
            cached['token'] = result['token']
        else:
            result = _request('/v1/status', {'device_id': did, 'license_id': cached['license_id']})
            cached['token'] = result['token']
        payload = verify_token(cached['token'])
        now = int(time.time())
        if payload.get('device_id') != did: raise LicenseError('This licence is registered to another device.')
        if now > int(payload.get('valid_until', 0)): raise LicenseError('Licence validation expired. Connect to the internet to renew the validation.')
        if payload.get('state') not in ('trial','lifetime'): raise LicenseError('This licence is not active.')
        if payload.get('state') == 'trial' and now > int(payload.get('trial_until', 0)): raise LicenseError('Your 30-day trial has ended. Contact MediBill Pro support to purchase the lifetime licence.')
        if payload.get('state') == 'lifetime' and now > int(payload.get('amc_until', 0)): raise LicenseError('Annual Maintenance Contract has expired. Renew AMC to continue using MediBill Pro. Your local data remains safe.')
        LICENSE_FILE.write_text(json.dumps(cached, indent=2), encoding='utf-8')
        return payload
    except Exception:
        # Allow only a short, signed-token offline grace; never extend the signed entitlement.
        token = cached.get('token')
        if token:
            payload = verify_token(token)
            now = int(time.time())
            issued = int(payload.get('issued_at', 0))
            if payload.get('device_id') == did and now <= int(payload.get('valid_until', 0)) + 3*86400:
                if payload.get('state') == 'trial' and now <= int(payload.get('trial_until', 0)): return payload
                if payload.get('state') == 'lifetime' and now <= int(payload.get('amc_until', 0)): return payload
        raise

def submit_payment(utr, contact=''):
    data = json.loads(LICENSE_FILE.read_text(encoding='utf-8'))
    return _request('/v1/payment-request', {'device_id':device_id(), 'license_id':data['license_id'], 'utr':utr.strip(), 'contact':contact.strip()})
