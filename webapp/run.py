"""Run the FactorMaker web server: python run.py [--reload]"""

from __future__ import annotations

import socket
import sys


def _patch_win314_socketpair() -> None:
	"""Python 3.14 on Windows can raise ConnectionError in socket.socketpair().

	Uvicorn/asyncio need socketpair for the event-loop self-pipe. The stdlib
	fallback authenticates peers via getsockname/getpeername and that check
	fails on some Windows setups — patch a simple localhost pair instead.
	"""
	if sys.platform != "win32" or sys.version_info < (3, 14):
		return
	if getattr(socket, "_factormaker_socketpair_patched", False):
		return

	def _socketpair(family=socket.AF_INET, type=socket.SOCK_STREAM, proto=0):
		if family not in (socket.AF_INET, socket.AF_INET6):
			raise ValueError("Only AF_INET and AF_INET6 are supported")
		if type != socket.SOCK_STREAM:
			raise ValueError("Only SOCK_STREAM is supported")
		host = "127.0.0.1" if family == socket.AF_INET else "::1"
		lsock = socket.socket(family, type, proto)
		try:
			lsock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
			lsock.bind((host, 0))
			lsock.listen(1)
			port = lsock.getsockname()[1]
			csock = socket.socket(family, type, proto)
			try:
				csock.connect((host, port))
				ssock, _peer = lsock.accept()
			except Exception:
				csock.close()
				raise
		finally:
			lsock.close()
		return ssock, csock

	socket.socketpair = _socketpair  # type: ignore[assignment]
	socket._factormaker_socketpair_patched = True  # type: ignore[attr-defined]


_patch_win314_socketpair()

import uvicorn


if __name__ == "__main__":
	want_reload = "--reload" in sys.argv
	# Multiprocessing reloader + Py3.14 socketpair is fragile on Windows;
	# default to a single in-process server. Pass --reload to opt in.
	if sys.platform == "win32" and sys.version_info >= (3, 14) and want_reload:
		print("Note: --reload on Windows/Python 3.14 may be unstable; trying anyway.")

	if want_reload:
		uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
	else:
		from app.main import app

		uvicorn.run(app, host="0.0.0.0", port=8000, reload=False)
