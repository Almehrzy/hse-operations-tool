# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Project settings shared by the scripts. Edit project.json in the project folder, not the scripts.

The project folder (ROOT) holds project.json, permits.json, equipment.json, monthly.json and templates\\.
It is found in this order, so the same scripts serve every project (the plugin's engine) and the
original tool folder alike:
  1. the HSE_PROJECT_DIR environment variable;
  2. the current folder or the nearest folder above it that holds project.json;
  3. the folder above this script (the original layout: scripts\\ inside the tool folder).
"""
import json, os


def _find_root():
    env = os.environ.get("HSE_PROJECT_DIR")
    if env and os.path.isfile(os.path.join(env, "project.json")):
        return os.path.abspath(env)
    d = os.path.abspath(os.getcwd())
    while True:
        if os.path.isfile(os.path.join(d, "project.json")):
            return d
        up = os.path.dirname(d)
        if up == d:
            break
        d = up
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


ROOT = _find_root()
with open(os.path.join(ROOT, "project.json"), encoding="utf-8") as _f:
    CFG = json.load(_f)


def path(p):
    """Absolute path; relative paths in project.json are relative to the project folder."""
    return p if os.path.isabs(p) else os.path.join(ROOT, p)
