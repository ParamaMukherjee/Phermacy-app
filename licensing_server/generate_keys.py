from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives.serialization import Encoding, PrivateFormat, PublicFormat, NoEncryption
import base64
key=Ed25519PrivateKey.generate()
open('private_key.pem','wb').write(key.private_bytes(Encoding.PEM,PrivateFormat.PKCS8,NoEncryption()))
public=key.public_key().public_bytes(Encoding.Raw,PublicFormat.Raw)
print('PUBLIC_KEY_B64='+base64.b64encode(public).decode())
print('Keep private_key.pem secret. Put it only in the licensing server secret storage.')
