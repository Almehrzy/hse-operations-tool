# HSE Operations Tool. Copyright (c) 2026 Khaled Almehrzy. All rights reserved. Proprietary: see LICENSE.txt.
"""Project settings shared by the scripts. Edit project.json (next to this folder), not the scripts."""
import json, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
with open(os.path.join(ROOT, "project.json"), encoding="utf-8") as _f:
    CFG = json.load(_f)


def path(p):
    """Absolute path; relative paths in project.json are relative to the tool folder."""
    return p if os.path.isabs(p) else os.path.join(ROOT, p)
