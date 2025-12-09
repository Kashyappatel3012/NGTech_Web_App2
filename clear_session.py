"""Clear session for debugging fingerprint issues"""
from app import app
from flask import session

with app.test_request_context():
    session.clear()
    print("Session cleared. Now try accessing /login again.")

