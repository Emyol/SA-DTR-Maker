"""Dev server entry point.

Run:
    python run_web.py

Then open http://127.0.0.1:5000 in a browser. For a real deployment, run the
`webapp` app factory under a production WSGI server (waitress, gunicorn,
etc.) instead of this script.
"""

import os

from webapp import create_app

app = create_app()

if __name__ == "__main__":
    debug = os.environ.get("DTR_DEBUG", "0") == "1"
    app.run(host="127.0.0.1", port=5000, debug=debug)
