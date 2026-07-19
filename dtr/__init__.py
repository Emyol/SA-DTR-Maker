"""Core DTR parsing/generation engine.

Pure, I/O-agnostic functions shared by the CLI (dtr_generator.py) and the
Flask web app (webapp/). Nothing in here touches the filesystem directly
except where a path is explicitly passed in by the caller.
"""
