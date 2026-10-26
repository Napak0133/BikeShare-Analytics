"""Create short-lived Tableau Connected App tokens for embedded views."""

import datetime
import os
import uuid

import jwt


class TableauEmbedConfigError(ValueError):
    """Raised when Connected App credentials for embedding are incomplete."""


def create_embed_token():
    """Create a five-minute JWT for embedding as the configured Tableau user."""
    required = {
        "TABLEAU_EMBED_CLIENT_ID": os.environ.get("TABLEAU_EMBED_CLIENT_ID", "").strip(),
        "TABLEAU_EMBED_SECRET_ID": os.environ.get("TABLEAU_EMBED_SECRET_ID", "").strip(),
        "TABLEAU_EMBED_SECRET": os.environ.get("TABLEAU_EMBED_SECRET", "").strip(),
        "TABLEAU_EMBED_USERNAME": os.environ.get("TABLEAU_EMBED_USERNAME", "").strip(),
    }
    missing = [name for name, value in required.items() if not value]
    if missing:
        raise TableauEmbedConfigError(
            "Set the following local Connected App settings in .env: "
            + ", ".join(missing)
        )

    now = datetime.datetime.now(datetime.timezone.utc)
    client_id = required["TABLEAU_EMBED_CLIENT_ID"]
    return jwt.encode(
        {
            "iss": client_id,
            "sub": required["TABLEAU_EMBED_USERNAME"],
            "aud": "tableau",
            "exp": now + datetime.timedelta(minutes=5),
            "jti": str(uuid.uuid4()),
            "scp": ["tableau:views:embed"],
        },
        required["TABLEAU_EMBED_SECRET"],
        algorithm="HS256",
        headers={
            "kid": required["TABLEAU_EMBED_SECRET_ID"],
            "iss": client_id,
        },
    )
