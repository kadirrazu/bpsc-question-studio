"""File-based operator access for this installation."""

import hmac
import os

from .file_io import write_bytes

DEFAULT_EMAIL = "kadir@bpsc.gov.bd"


def initialize_access(app_folder):
    private = app_folder / ".studio"
    private.mkdir(mode=0o700, exist_ok=True)
    if os.name == "nt":
        import ctypes
        kernel = ctypes.windll.kernel32
        attributes = kernel.GetFileAttributesW(str(private))
        if attributes != -1:
            kernel.SetFileAttributesW(str(private), attributes | 2)
    path = private / "credentials.txt"
    if not path.exists():
        write_bytes(path, f"email={DEFAULT_EMAIL}\npassword=96321@Aa\n".encode("utf-8"))
    return path


def matches(path, password, email=None):
    fields = {}
    for line in path.read_text("utf-8-sig").splitlines():
        if line.strip():
            name, separator, value = line.partition("=")
            name = name.strip().lower()
            if not separator or name not in ("email", "password") or name in fields:
                raise ValueError("Use one email= line and one password= line in credentials.txt.")
            fields[name] = value if name == "password" else value.strip()
    if not fields.get("email") or not fields.get("password"):
        raise ValueError("Add the email and password to credentials.txt.")
    same_password = hmac.compare_digest(password.encode("utf-8"), fields["password"].encode("utf-8"))
    if email is None:
        return same_password
    same_email = hmac.compare_digest(email.strip().casefold().encode("utf-8"), fields["email"].casefold().encode("utf-8"))
    return same_email and same_password
