"""Per-application COM hosts, selected by app key."""
from __future__ import annotations

from .access import AccessHost
from .base import HostError, OfficeHost, describe_com_error
from .excel import ExcelHost
from .powerpoint import PowerPointHost
from .word import WordHost

HOSTS: dict[str, type[OfficeHost]] = {
    ExcelHost.app_key: ExcelHost,
    WordHost.app_key: WordHost,
    PowerPointHost.app_key: PowerPointHost,
    AccessHost.app_key: AccessHost,
}


def build_host(app: str) -> OfficeHost:
    try:
        return HOSTS[app]()
    except KeyError:
        raise ValueError(
            f"Unknown app {app!r}; expected one of "
            f"{', '.join(sorted(HOSTS))}.") from None


__all__ = [
    "HOSTS",
    "AccessHost",
    "ExcelHost",
    "HostError",
    "OfficeHost",
    "PowerPointHost",
    "WordHost",
    "build_host",
    "describe_com_error",
]
