from pathlib import Path

import yaml

HF_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = HF_DIR.parent


def test_render_blueprint():
    config = yaml.safe_load((REPO_ROOT / "render.yaml").read_text(encoding="utf-8"))
    service = config["services"][0]
    assert service["rootDir"] == "hf-space"
    assert "$PORT" in service["startCommand"]
    assert service["healthCheckPath"] == "/health"
    env = {e["key"]: e for e in service["envVars"]}
    assert env["DATABASE_URL"]["sync"] is False


def test_requirements():
    text = (HF_DIR / "requirements.txt").read_text(encoding="utf-8").lower()
    assert "psycopg[binary]" in text
    for banned in ("gradio", "huggingface_hub", "sqlite"):
        assert banned not in text
