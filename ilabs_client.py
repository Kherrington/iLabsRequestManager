"""
iLab Solutions API client for https://api.ilab.agilent.com

Authentication: Bearer token in the Authorization header.
Pass token= directly or set the ILAB_TOKEN environment variable.
"""

import os
import pathlib
import requests
from typing import List, Optional, Union

# Core IDs may be numeric (1234) or a slug ("CALM") depending on the iLab instance.
CoreID = Union[int, str]


def _load_dotenv() -> None:
    """
    Load a .env file into os.environ (without requiring python-dotenv).
    Searches the directory containing this file, then the current working directory.
    Existing environment variables are never overwritten.
    """
    candidates = [
        pathlib.Path(__file__).parent / ".env",
        pathlib.Path.cwd() / ".env",
    ]
    for env_path in candidates:
        if env_path.exists():
            with open(env_path, encoding="utf-8") as fh:
                for line in fh:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, _, value = line.partition("=")
                    key   = key.strip()
                    value = value.strip().strip('"').strip("'")
                    os.environ.setdefault(key, value)
            break   # stop after the first .env found


_load_dotenv()


class ILabError(Exception):
    """Raised when the iLab API returns an error response."""

    def __init__(self, status_code: int, message: str):
        self.status_code = status_code
        super().__init__(f"HTTP {status_code}: {message}")


def _as_list(value) -> list:
    """Normalize a value that might be a list, a single dict, or None."""
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


class ILabClient:
    """
    REST client for the iLab Solutions API.

    Usage::

        client = ILabClient()                          # reads ILAB_TOKEN from env
        client = ILabClient(token="abc123")            # explicit token
        client = ILabClient(token="abc123",
                            base_url="https://ucsf.ilab.agilent.com")
    """

    def __init__(
        self,
        token: Optional[str] = None,
        base_url: Optional[str] = None,
    ):
        token = token or os.environ.get("ILAB_TOKEN")
        if not token:
            raise ValueError(
                "An API token is required. Set the ILAB_TOKEN environment variable "
                "or pass token= to ILabClient()."
            )
        # Resolve base URL: explicit arg → ILAB_BASE_URL env var → hard default
        resolved = base_url or os.environ.get("ILAB_BASE_URL", "https://api.ilabsolutions.com")
        self.base_url = resolved.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json",
                "Content-Type": "application/json",
            }
        )

    # ── Internal HTTP helpers ─────────────────────────────────────────────────

    def _url(self, path: str) -> str:
        return f"{self.base_url}/v1{path}"

    def _request(self, method: str, path: str, **kwargs) -> dict:
        import time
        # Retry up to 3 times on 429 (rate limit), backing off 5 → 10 → 20 s
        delays = [5, 10, 20]
        for attempt, delay in enumerate(delays + [None]):
            resp = self.session.request(method, self._url(path), **kwargs)
            if resp.status_code == 429 and delay is not None:
                time.sleep(delay)
                continue
            break
        if resp.status_code == 204:
            return {}
        if not resp.ok:
            try:
                detail = resp.json()
            except Exception:
                detail = resp.text or resp.reason
            raise ILabError(resp.status_code, str(detail))
        return resp.json()

    def _unwrap(self, data: dict) -> dict:
        """Strip the ilab_response envelope if present."""
        return data.get("ilab_response", data)

    def _get_all(
        self, path: str, collection_key: str, params: Optional[dict] = None
    ) -> List[dict]:
        """Fetch every page of a paginated collection and return all items."""
        params = dict(params or {})
        page = 1
        results: List[dict] = []

        while True:
            params["page"] = page
            data = self._unwrap(self._request("GET", path, params=params))
            items = _as_list(data.get(collection_key))
            results.extend(items)

            meta = data.get("ilab_metadata") or {}
            next_page = meta.get("next_page")
            total_pages = meta.get("total_pages")

            if not items:
                break
            if next_page is None:
                break
            if total_pages is not None and page >= int(total_pages):
                break
            page += 1

        return results

    # ── Cores ─────────────────────────────────────────────────────────────────

    def list_cores(self) -> List[dict]:
        """Return all cores this token has access to."""
        data = self._unwrap(self._request("GET", "/cores.json"))
        return _as_list(data.get("cores"))

    def get_core(self, core_id: int) -> dict:
        """Return details of a single core."""
        data = self._request("GET", f"/cores/{core_id}.json")
        return data.get("core", data)

    # ── Services ──────────────────────────────────────────────────────────────

    def list_services(self, core_id: int) -> List[dict]:
        """Return all services offered by a core."""
        return self._get_all(f"/cores/{core_id}/services.json", "services")

    def get_service(self, core_id: int, service_id: int) -> dict:
        data = self._request("GET", f"/cores/{core_id}/services/{service_id}.json")
        return data.get("service", data)

    def update_service(self, core_id: int, service_id: int, **fields) -> dict:
        """
        Update a service.
        Common fields: name, description, public_visibility (0=hidden, 1=visible)
        """
        data = self._request(
            "PUT",
            f"/cores/{core_id}/services/{service_id}.json",
            json={"service": fields},
        )
        return data.get("service", data)

    # ── Prices ────────────────────────────────────────────────────────────────

    def list_prices(self, core_id: int, service_id: int) -> List[dict]:
        return self._get_all(
            f"/cores/{core_id}/services/{service_id}/prices.json", "prices"
        )

    def get_price(self, core_id: int, service_id: int, price_id: int) -> dict:
        data = self._request(
            "GET", f"/cores/{core_id}/services/{service_id}/prices/{price_id}.json"
        )
        return data.get("price", data)

    def update_price(
        self, core_id: int, service_id: int, price_id: int, price: float, **fields
    ) -> dict:
        data = self._request(
            "PUT",
            f"/cores/{core_id}/services/{service_id}/prices/{price_id}.json",
            json={"price": {"price": price, **fields}},
        )
        return data.get("price", data)

    def delete_price(self, core_id: int, service_id: int, price_id: int) -> None:
        self._request(
            "DELETE",
            f"/cores/{core_id}/services/{service_id}/prices/{price_id}.json",
        )

    # ── Service Requests ──────────────────────────────────────────────────────

    def list_service_requests(self, core_id: int, **filters) -> List[dict]:
        """
        Return service requests for a core (default: last 2 years).

        Optional keyword filters
        ------------------------
        from_date   ISO 8601 string, e.g. "2024-01-01"
        to_date     ISO 8601 string
        states      comma-separated, e.g. "processing,completed"
                    Valid values: proposed, needs_financial_reapproval, processing,
                    completed, cancelled, draft, financials_approved,
                    financials_rejected, requested, ...
        q           full-text search string
        name        exact name match
        has_recurring  0 or 1
        per_page    results per page (default 30)
        """
        return self._get_all(
            f"/cores/{core_id}/service_requests.json",
            "service_requests",
            params=filters,
        )

    def get_service_request(self, core_id: int, request_id: int) -> dict:
        data = self._request(
            "GET", f"/cores/{core_id}/service_requests/{request_id}.json"
        )
        return data.get("service_request", data)

    def create_service_request(
        self, core_id: int, owner_email: str, **kwargs
    ) -> dict:
        """
        Create a new service request.

        Required: owner_email
        Optional kwargs: pi_email, name,
          state (proposed | processing | completed, default completed)
        """
        data = self._request(
            "POST",
            f"/cores/{core_id}/service_requests.json",
            json={"service_request": {"owner_email": owner_email, **kwargs}},
        )
        return data.get("service_request", data)

    def update_service_request(
        self, core_id: int, request_id: int, **fields
    ) -> dict:
        """
        Update a service request.

        Updatable fields: name, description, state, completed_on, start_on,
                          end_on, quote_expires_on, projected_cost, summary
        """
        data = self._request(
            "PUT",
            f"/cores/{core_id}/service_requests/{request_id}.json",
            json={"service_request": fields},
        )
        return data.get("service_request", data)



    # ── Custom Forms ──────────────────────────────────────────────────────────

    def list_custom_forms(self, core_id: int, request_id: int) -> List[dict]:
        data = self._unwrap(
            self._request(
                "GET",
                f"/cores/{core_id}/service_requests/{request_id}/custom_forms.json",
            )
        )
        return _as_list(data.get("custom_forms"))

    # ── Milestones ────────────────────────────────────────────────────────────

    def list_milestones(self, core_id: int, request_id: int) -> List[dict]:
        data = self._unwrap(
            self._request(
                "GET",
                f"/cores/{core_id}/service_requests/{request_id}/milestones.json",
            )
        )
        return _as_list(data.get("milestones"))

    def update_milestone(
        self, core_id: int, request_id: int, milestone_id: int, **fields
    ) -> dict:
        """
        Update a milestone.
        Updatable fields: name, description, started_on, completed_on (ISO 8601)
        """
        data = self._request(
            "PUT",
            f"/cores/{core_id}/service_requests/{request_id}/milestones/{milestone_id}.json",
            json={"milestone": fields},
        )
        return data.get("milestone", data)

    # ── Charges ───────────────────────────────────────────────────────────────

    def list_charges(self, core_id: int, request_id: int) -> List[dict]:
        """Return all charges for a service request."""
        return self._get_all(
            f"/cores/{core_id}/service_requests/{request_id}/charges.json",
            "charges",
        )

    def add_charges(
        self, core_id: int, request_id: int, charges: List[dict]
    ) -> dict:
        """
        Add one or more charges to a service request.

        Each charge dict must include:
          quantity   (float)
          price_id   (int)
          service_id (int)
        Optional per charge:
          note       (str)
        """
        return self._request(
            "POST",
            f"/cores/{core_id}/service_requests/{request_id}/charges.json",
            json={"charges": charges},
        )

    def update_charge(
        self, core_id: int, request_id: int, charge_id: int, **fields
    ) -> dict:
        """
        Update a charge.
        Updatable fields: name (if core allows), quantity,
                          billing_status, status, note
        Billing status values: cancelled, not_ready_to_bill, ready_to_bill,
                               not_billable, pro_bono, billed, paid
        Status values: proposed, financials_approved, processing, completed, cancelled
        """
        data = self._request(
            "PUT",
            f"/cores/{core_id}/service_requests/{request_id}/charges/{charge_id}.json",
            json={"charge": fields},
        )
        return data.get("charge", data)

    # ── Equipment ─────────────────────────────────────────────────────────────

    def list_equipment(self, core_id: int) -> List[dict]:
        """Return all equipment registered in a core."""
        return self._get_all(f"/cores/{core_id}/equipment.json", "equipment")

    # ── Attachments ───────────────────────────────────────────────────────────

    def get_attachment(self, attachment_id: int) -> bytes:
        """Download an attachment by ID. Returns raw bytes."""
        resp = self.session.get(self._url(f"/attachments/{attachment_id}"))
        if not resp.ok:
            raise ILabError(resp.status_code, resp.text)
        return resp.content

    def add_attachment(
        self, request_id: int, file_path: str, name: Optional[str] = None
    ) -> dict:
        """
        Upload a file attachment to a service request.

        file_path: local path of the file to upload
        name: optional display name; defaults to the filename
        """
        params = {"object_class": "ServiceItem", "id": request_id}
        with open(file_path, "rb") as fh:
            files = {"attachment[uploaded_data]": fh}
            form_data = {}
            if name:
                form_data["attachment[name]"] = name
            # Remove Content-Type so requests sets the correct multipart boundary
            headers = {
                k: v
                for k, v in self.session.headers.items()
                if k.lower() != "content-type"
            }
            resp = self.session.post(
                self._url("/attachments"),
                params=params,
                files=files,
                data=form_data,
                headers=headers,
            )
        if not resp.ok:
            raise ILabError(resp.status_code, resp.text)
        return resp.json()

    def delete_attachment(self, attachment_id: int) -> dict:
        """Delete an attachment. Returns the updated service request."""
        return self._request("DELETE", f"/attachments/{attachment_id}")
