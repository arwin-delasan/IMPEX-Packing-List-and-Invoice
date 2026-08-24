"""Minimal Odoo XML-RPC client for product lookups.

ODOO_URL/ODOO_DB always come from .env - the server and database are the
same for everyone, so there's no reason to ask for them per-user. Username
and password are different: OdooClient accepts them as explicit arguments
(the GUI prompts for these via a login dialog rather than baking one
shared login into .env) and only falls back to ODOO_USERNAME/ODOO_PASSWORD
in .env if they aren't given - kept for CLI/scripting convenience, not
used by the GUI. Uses only the standard library (xmlrpc.client) - no new
dependency needed.
"""

import xmlrpc.client

from app_paths import app_path


class OdooError(Exception):
    pass


def _load_env(env_path=None):
    env_path = env_path or app_path(".env")
    env = {}
    with open(env_path, encoding="utf-8") as f:
        for line in f:
            line = line.split("#", 1)[0].strip()
            if not line or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    return env


class OdooClient:
    def __init__(self, username=None, password=None, env_path=None):
        env_path = env_path or app_path(".env")
        env = _load_env(env_path)
        try:
            url = env["ODOO_URL"]
            db = env["ODOO_DB"]
        except KeyError as e:
            raise OdooError(f"Missing {e.args[0]} in {env_path}")

        username = username or env.get("ODOO_USERNAME")
        password = password or env.get("ODOO_PASSWORD")
        if not username or not password:
            raise OdooError(
                "Odoo username/password not provided, and ODOO_USERNAME/"
                f"ODOO_PASSWORD aren't set in {env_path} either."
            )

        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
        uid = common.authenticate(db, username, password, {})
        if not uid:
            raise OdooError(f"Odoo authentication failed for {username} at {url}")

        self._models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")
        self._db = db
        self._uid = uid
        self._password = password

    def execute(self, model, method, *args, **kwargs):
        return self._models.execute_kw(
            self._db, self._uid, self._password, model, method, list(args), kwargs
        )

    def lookup_products_by_code(self, codes):
        """Return {code: {"name": str, "categ_id": [id, complete_name] or False}}
        for every given default_code found in Odoo. Codes not found are
        simply absent from the result."""
        codes = list(dict.fromkeys(codes))  # de-dupe, preserve order
        ids = self.execute("product.product", "search", [["default_code", "in", codes]])
        recs = self.execute("product.product", "read", ids, fields=["default_code", "name", "categ_id"])
        return {r["default_code"]: r for r in recs}


def prompt_login():
    """CLI convenience: try ODOO_USERNAME/ODOO_PASSWORD from .env first
    (unchanged behavior when they're set); if they're missing - the
    normal case now that the GUI prompts for login instead of storing
    them in .env - fall back to an interactive terminal prompt rather
    than just failing."""
    try:
        return OdooClient()
    except OdooError:
        import getpass
        username = input("Odoo username: ")
        password = getpass.getpass("Odoo password: ")
        return OdooClient(username=username, password=password)
