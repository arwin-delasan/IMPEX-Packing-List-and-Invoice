"""Minimal Odoo XML-RPC client for product lookups.

Reads connection settings from .env (ODOO_URL, ODOO_DB, ODOO_USERNAME,
ODOO_PASSWORD). Uses only the standard library (xmlrpc.client) - no new
dependency needed.
"""

import xmlrpc.client


class OdooError(Exception):
    pass


def _load_env(env_path=".env"):
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
    def __init__(self, env_path=".env"):
        env = _load_env(env_path)
        try:
            url = env["ODOO_URL"]
            db = env["ODOO_DB"]
            username = env["ODOO_USERNAME"]
            password = env["ODOO_PASSWORD"]
        except KeyError as e:
            raise OdooError(f"Missing {e.args[0]} in {env_path}")

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

    def lookup_contact_by_company_name(self, company_name):
        """Find the company in res.partner matching `company_name` (a
        contains/ilike search - PDF consignee names don't always match
        Odoo's spelling exactly), and return its first *named* contact as
        {"name", "phone"}, or None if no company match or no named
        contact exists.

        A trailing "." is stripped first, since PDF consignee names are
        often written like "SPARTA SPA LTD." while Odoo has "Sparta Spa
        LTD" with no trailing period.

        A company's child_ids can include auto-generated address records
        (type "invoice"/"delivery"/etc.) with no name of their own - e.g.
        one real Odoo company had child_ids [invoice-address (name=False),
        Scott Campbell (type=contact)]. Picking child_ids[0] blindly would
        have returned the nameless address record, so this prefers
        type="contact" records and falls back to any child with a name."""
        search_name = company_name.strip()
        if search_name.endswith("."):
            search_name = search_name[:-1]
        if not search_name:
            return None

        company_ids = self.execute(
            "res.partner", "search",
            [["name", "ilike", search_name], ["is_company", "=", True]],
            limit=1,
        )
        if not company_ids:
            return None

        company = self.execute("res.partner", "read", company_ids, fields=["child_ids"])[0]
        if not company["child_ids"]:
            return None

        children = self.execute(
            "res.partner", "read", company["child_ids"],
            fields=["name", "type", "phone", "mobile"],
        )
        named = [c for c in children if c["name"]]
        if not named:
            return None
        contact_type = [c for c in named if c["type"] == "contact"]
        child = contact_type[0] if contact_type else named[0]
        return {
            "name": child["name"],
            "phone": child["phone"] or child["mobile"] or None,
        }
