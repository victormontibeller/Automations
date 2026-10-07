"""Install only in disposable test environments, never in the runtime venv."""
import sys


def _deny_network(event, args):
    if event in {"socket.connect", "socket.getaddrinfo", "socket.gethostbyname",
                 "socket.gethostbyaddr", "socket.sendto", "socket.sendmsg"}:
        raise RuntimeError("Network disabled during tests")


sys.addaudithook(_deny_network)
