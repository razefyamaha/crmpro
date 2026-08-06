"""WSGI entrypoint for cloud platforms (Render, Railway, etc.) and gunicorn.

Local run is unaffected: `python app.py` still works exactly as before.
On a cloud server, the web server imports the Flask `app` object from here.
"""
import os

# Ensure the app binds publicly when started via gunicorn (0.0.0.0).
os.environ.setdefault("CRM_HOST", "0.0.0.0")

from app import app as application  # noqa: E402

app = application
