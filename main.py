"""
Single entrypoint that combines all three services into ONE Flask app,
so it fits PythonAnywhere's free-tier "1 web app" limit:

  /banca/...   -> Banca / Standard Grid Processor Portal
  /agency/...  -> Agency Special Motor Matrix Portal
  /slack/...   -> GridBot Slack Events API endpoint

Local testing:
    python3 main.py            # runs on http://localhost:5000

PythonAnywhere deployment:
    Point your web app's WSGI file at this module's `app` object
    (see wsgi_pythonanywhere_template.py for the exact glue code).
"""

import os
from flask import Flask, redirect

from banca.routes import banca_bp
from agency.routes import agency_bp
from slackbot.routes import slack_bp

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB, matches the larger of the two portals

app.register_blueprint(banca_bp, url_prefix='/banca')
app.register_blueprint(agency_bp, url_prefix='/agency')
app.register_blueprint(slack_bp, url_prefix='/slack')


@app.route('/')
def home():
    # Convenience landing page — sends people to the Banca portal by default.
    return redirect('/banca/')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
