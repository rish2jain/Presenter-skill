"""Shared pytest fixtures for the presentation-skill test suite."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))


@pytest.fixture(autouse=True)
def _reset_global_state():
    """Reset shared module-level state after every test.

    Prevents CJK font-overlay and density settings from leaking between
    tests that run in the same process.  Local autouse fixtures that set
    a specific density before yield are intentionally left in place (they
    run first on setup, this fixture runs last on teardown).
    """
    yield
    import palettes
    import builders
    palettes.set_cjk(False)
    builders.set_density("")
