"""WSGI adapter for hosts (like PythonAnywhere's free tier) that only run WSGI apps.
FastAPI is ASGI-native; a2wsgi bridges it so PythonAnywhere's WSGI-based `wsgi.py`
config can point at a plain callable."""
from a2wsgi import ASGIMiddleware

from .main import app as asgi_app

application = ASGIMiddleware(asgi_app)
