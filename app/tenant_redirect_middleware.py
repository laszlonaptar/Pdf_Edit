from __future__ import annotations
from typing import Iterable, Optional, Set
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import RedirectResponse

class TenantRootRedirectMiddleware(BaseHTTPMiddleware):
    """
    Multitenant gyökér-átirányítás:
      - ha a Host egy tenant-aldomain (*.BASE_DOMAIN), de NEM a fődomain (metori.de) és NEM a www,
      - és az útvonal pontosan "/",
      - akkor átirányít loginra (/login?next=/form).

    Skálázható: nem kell felsorolni az aldomaineket; minden jövőbeni *.metori.de-re működik.
    """

    def __init__(
        self,
        app,
        base_domain: str,
        main_domains: Optional[Iterable[str]] = None,
        login_target: str = "/login?next=/form",
    ):
        super().__init__(app)
        self.base_domain = (base_domain or "").lower().lstrip(".")
        defaults = [self.base_domain, f"www.{self.base_domain}"]
        mains: Set[str] = set(d.lower() for d in (main_domains or defaults) if d)
        # Biztonság kedvéért tisztítjuk a portot is
        self.main_domains = {d.split(":", 1)[0] for d in mains}
        self.login_target = login_target

    async def dispatch(self, request, call_next):
        host = (request.headers.get("host") or request.url.hostname or "").lower()
        host = host.split(":", 1)[0]  # port lecsupaszítása
        path = (request.url.path or "/")
        # Tenant-aldomain detektálása: *.base_domain és nem a fő/main domainek egyike
        is_tenant = host.endswith("." + self.base_domain) and host not in self.main_domains

        if is_tenant and path == "/":
            return RedirectResponse(self.login_target, status_code=302)

        return await call_next(request)
