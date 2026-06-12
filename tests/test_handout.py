"""Tests for gen_handout.py — outline → pre-read markdown handout."""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from gen_handout import gen_handout  # noqa: E402


def test_example_outline_handout():
    text = (ROOT / "assets" / "example-outline.md").read_text(encoding="utf-8")
    out = gen_handout(text)
    assert out.startswith("# AI Infrastructure Strategy\n")
    assert "## 2. Three priorities driving $14.2M in value" in out
    assert "## 1." not in out  # title slide becomes the doc title
    assert "| Label | Enterprise AI spend ($B) |" in out
    assert "| 2024 | 42 |" in out
    assert "> Talk track: Walk through each priority" in out


def test_handout_bullets_and_source():
    out = gen_handout(
        "## Slide 1: Costs have outgrown revenue for three years\n"
        "- icon:trending-up Cost CAGR 12%\n"
        "- Revenue CAGR 4%\n"
        "- Source: Company filings\n"
        "- Notes: Set up the problem.\n")
    assert "# Costs have outgrown revenue for three years" in out
    assert "- Cost CAGR 12%" in out  # icon prefix stripped
    assert "- Revenue CAGR 4%" in out
    assert "*Source: Company filings*" in out
    assert "> Talk track: Set up the problem." in out


def test_handout_table_rows_reemitted():
    out = gen_handout(
        "## Slide 1: Title here\n**Layout:** title\n- Title: Deck\n\n"
        "## Slide 2: Vendor B leads on the criteria that matter\n"
        "**Layout:** table\n"
        "| Criterion | A | B |\n"
        "|---|---|---|\n"
        "| Cost | 2 | 4 |\n"
        "- Notes: n.\n")
    assert "| Criterion | A | B |" in out
    assert "| Cost | 2 | 4 |" in out


def test_handout_multi_series_table_uses_series_headers():
    out = gen_handout(
        "## Slide 1: Revenue outpaces costs from Q1 onward\n"
        "**Visual:** chart:line\n"
        "**Series:** Revenue, Costs\n"
        "**Data:**\n- Q1: 4.2, 3.1\n- Q2: 5.1, 3.0\n"
        "- Notes: n.\n")
    assert "| Label | Revenue | Costs |" in out
    assert "| Q1 | 4.2 | 3.1 |" in out


def test_handout_appendix_divider():
    out = gen_handout(
        "## Slide 1: Deck title goes right here\n**Layout:** title\n\n"
        "## Slide 2: Costs have outgrown revenue for years\n"
        "- a bullet\n- Notes: n.\n\n"
        "## Appendix\n\n"
        "## Slide 3: Backup cost detail by region\n"
        "- detail\n")
    assert "# Appendix" in out
    assert out.index("## 2.") < out.index("# Appendix") < out.index("## 3.")


def test_handout_skips_auto_agenda_slides():
    out = gen_handout(
        "**Auto-Agenda:** on\n\n"
        "## Slide 1: Acme FY27 Strategy\n**Layout:** title\n\n"
        "## Slide 2: Diagnosis\n**Layout:** section-divider\n\n"
        "## Slide 3: Costs have outgrown revenue for years\n"
        "- a bullet\n- Notes: n.\n")
    assert "Agenda" not in out
    # numbering matches deck positions (agenda inserted as slide 2)
    assert "## 3. Diagnosis" in out
    assert "## 4. Costs have outgrown revenue for years" in out


def test_handout_includes_references_slide():
    out = gen_handout(
        "**References:** on\n\n"
        "## Slide 1: Deck title goes right here\n**Layout:** title\n\n"
        "## Slide 2: Costs have outgrown revenue for years\n"
        "- a bullet\n- Source: Gartner 2026\n- Notes: n.\n")
    assert "# Appendix" in out
    assert "- Slide 2 — Gartner 2026" in out


def test_handout_cli_default_output(tmp_path):
    md = tmp_path / "pitch.md"
    md.write_text("## Slide 1: Deck title goes right here\n**Layout:** title\n\n"
                  "## Slide 2: Costs have outgrown revenue for years\n"
                  "- a bullet\n- Notes: n.\n", encoding="utf-8")
    r = subprocess.run([sys.executable, str(SCRIPTS / "gen_handout.py"),
                        str(md)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    out_file = tmp_path / "pitch-handout.md"
    assert out_file.is_file()
    assert "## 2. Costs have outgrown revenue for years" in \
        out_file.read_text(encoding="utf-8")


def test_handout_cli_explicit_output(tmp_path):
    md = tmp_path / "o.md"
    md.write_text("## Slide 1: Deck title goes right here\n", encoding="utf-8")
    dest = tmp_path / "pre-read.md"
    r = subprocess.run([sys.executable, str(SCRIPTS / "gen_handout.py"),
                        str(md), "--output", str(dest)],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert dest.is_file()
