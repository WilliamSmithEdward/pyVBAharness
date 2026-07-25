"""Worker-process package: all Excel COM access lives here.

The supervisor (session.py) spawns ``python -m pyvbaharness.worker`` and
communicates over the pipe protocol in pyvbaharness.protocol.
"""
