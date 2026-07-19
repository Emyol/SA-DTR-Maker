"""Flask app factory for the DTR web UI."""

import os

from flask import Flask, render_template

# Kept under Vercel's serverless function request-body limit (4.5 MB) so our
# own check is the one that fires with a friendly message, not a platform 413.
MAX_UPLOAD_MB = 4


def create_app():
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
    app.config["SECRET_KEY"] = os.environ.get("DTR_SECRET_KEY", "dev-only-change-me")

    from . import routes
    app.register_blueprint(routes.bp)

    @app.errorhandler(404)
    def not_found(err):
        message = getattr(err, "description", None) or "That page doesn't exist."
        return render_template("error.html", message=message), 404

    @app.errorhandler(413)
    def too_large(err):
        return render_template(
            "error.html",
            message=f"That file is too large (max {MAX_UPLOAD_MB} MB).",
        ), 413

    @app.errorhandler(500)
    def server_error(err):
        return render_template(
            "error.html",
            message="Something went wrong on our end. Please try again.",
        ), 500

    return app
