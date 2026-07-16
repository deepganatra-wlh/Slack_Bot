# ═══════════════════════════════════════════════════════════════════════════
# PythonAnywhere WSGI configuration file
#
# On PythonAnywhere: go to the "Web" tab -> click your web app -> under
# "Code" click the link to your WSGI configuration file (something like
# /var/www/yourusername_pythonanywhere_com_wsgi.py) -> DELETE everything in
# it -> paste this in -> fix the two paths marked below -> Save -> hit the
# green "Reload" button at the top of the Web tab.
# ═══════════════════════════════════════════════════════════════════════════

import sys
import os

# 1) The path to the folder containing main.py, banca/, agency/, slackbot/
#    (this is the folder you uploaded/cloned this project into)
project_home = '/home/YOURUSERNAME/gridbot'   # <-- CHANGE THIS
if project_home not in sys.path:
    sys.path.insert(0, project_home)

# 2) Load your environment variables (SLACK_BOT_TOKEN, SLACK_SIGNING_SECRET,
#    GROQ_API_KEY) here so the app can see them. Easiest approach: put them
#    in a .env file in project_home and load it with python-dotenv, OR just
#    set them directly below (fine for a small internal tool):
os.environ.setdefault('SLACK_BOT_TOKEN', 'xoxb-REPLACE-ME')
os.environ.setdefault('SLACK_SIGNING_SECRET', 'REPLACE-ME')
os.environ.setdefault('GROQ_API_KEY', 'gsk-REPLACE-ME')

from main import app as application  # PythonAnywhere looks for `application`
