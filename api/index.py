"""Vercel serverless entrypoint.

Vercel's Python runtime looks for a WSGI-compatible `app` object exported
from a file under api/. This just wires up the same Flask app the local dev
server (run_web.py) uses — no separate code path.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from webapp import create_app

app = create_app()
