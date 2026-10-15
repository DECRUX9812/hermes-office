"""Hermes loads this directory as a plugin package; register() lives in the importable core."""
try:
    from .hermes_office.plugin import register
except ImportError:
    from hermes_office.plugin import register  # type: ignore[no-redef]
