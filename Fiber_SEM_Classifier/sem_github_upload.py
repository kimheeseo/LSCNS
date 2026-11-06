"""Package trained SEM models and upload them as a GitHub Release asset.

No upload occurs on import. Tokens are provided by the caller, never persisted.
Original/augmented SEM images, predictions and optimizer checkpoints are excluded.
"""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import uuid
from zipfile import ZipFile, ZIP_DEFLATED

import requests


MODEL_NAMES = ("best_resnet.pth", "best_efficientnet.pth")
REPORT_NAMES = (
    "metrics.csv", "augmentation_effect.csv", "hcf_comparison.csv",
    "confusion_matrix.png", "loss_curve.png", "accuracy_curve.png",
)


def prepare_bundle(pipeline):
    """Validate training outputs and save an allowlisted ZIP beside the project.

    Returns paths and the stable tag stored in the ZIP. Re-uploading this same ZIP
    uses its tag and never replaces an already uploaded asset.
    """
    models = Path(pipeline.models).resolve()
    reports = Path(pipeline.reports).resolve()
    selected_path = models / "selected_model.json"
    required = [models / name for name in MODEL_NAMES]
    required += [selected_path, reports / "metrics.csv"]
    for path in required:
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError(f"학습 완료 파일이 없습니다: {path.name}")

    selected = json.loads(selected_path.read_text(encoding="utf-8"))
    selected_name = Path(selected["best_model_path"]).name
    if selected_name not in MODEL_NAMES:
        raise RuntimeError("선택 모델이 현재 지원하는 최적 모델 파일이 아닙니다.")

    tag = "sem-model-" + datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    tag += "-" + uuid.uuid4().hex[:8]
    out = Path(pipeline.root).resolve() / "07_GitHub_Exports"
    out.mkdir(parents=True, exist_ok=True)
    zip_path = out / f"{tag}.zip"
    partial = zip_path.with_suffix(".zip.partial")
    included = {}
    files = [(models / n, f"03_Models/{n}") for n in MODEL_NAMES]
    files += [(reports / n, f"04_Training_Reports/{n}")
              for n in REPORT_NAMES if (reports / n).is_file()]
    try:
        with ZipFile(partial, "w", compression=ZIP_DEFLATED, compresslevel=1) as z:
            for source, target in files:
                z.write(source, target)
                h = hashlib.sha256()
                with source.open("rb") as stream:
                    for chunk in iter(lambda: stream.read(4 * 1024 * 1024), b""):
                        h.update(chunk)
                included[target] = {"bytes": source.stat().st_size, "sha256": h.hexdigest()}
            # Preserve selection while removing the original Colab absolute path.
            selection = {"model": selected.get("model"),
                         "best_model_path": selected_name,
                         "validation_macro_f1": selected.get("validation_macro_f1")}
            z.writestr("03_Models/selected_model.json",
                       json.dumps(selection, ensure_ascii=False, indent=2))
            z.writestr("bundle.json", json.dumps({
                "format_version": 1, "tag": tag,
                "created_at_utc": datetime.now(timezone.utc).isoformat(),
                "selected_model": selected_name, "files": included,
            }, ensure_ascii=False, indent=2))
            z.writestr("RESTORE.txt", (
                "This archive contains trained weights and aggregate reports.\n"
                "Use the same Pipeline code and preprocessing for inference.\n"
                "With the original Pipeline, pass model_choice='resnet18' or\n"
                "model_choice='efficientnet_b0' to avoid absolute-path selection.\n"
                "Alternatively set selected_model.json best_model_path to the\n"
                "absolute path of the extracted selected model before use.\n"
                "Original training photos are not included.\n"
            ))
        if partial.stat().st_size >= 2 * 1024**3:
            raise RuntimeError("ZIP이 2 GiB 이상입니다. 모델별로 분리해야 합니다.")
        partial.replace(zip_path)
    finally:
        partial.unlink(missing_ok=True)
    return zip_path


def _checked(response, action):
    if not response.ok:
        # Avoid displaying request headers, tokens or private response payloads.
        raise RuntimeError(f"{action} 실패: HTTP {response.status_code}. "
                           "토큰 권한·저장소·네트워크를 확인하세요.")
    return response.json()


def upload_bundle(zip_path, token, repo="kimheeseo/LSCNS"):
    """Publish a prepared ZIP. Retry with the same ZIP after a network failure."""
    if not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repo):
        raise ValueError("repo는 owner/name 형식이어야 합니다.")
    if not token or not token.strip():
        raise ValueError("GITHUB_TOKEN이 필요합니다.")
    zip_path = Path(zip_path)
    with ZipFile(zip_path) as z:
        tag = json.loads(z.read("bundle.json"))["tag"]
    if not re.fullmatch(r"sem-model-[A-Za-z0-9-]+", tag):
        raise ValueError("유효하지 않은 모델 번들 태그입니다.")
    api = f"https://api.github.com/repos/{repo}"
    with requests.Session() as session:
        session.headers.update({"Authorization": f"Bearer {token.strip()}",
                                "Accept": "application/vnd.github+json"})
        found = session.get(f"{api}/releases/tags/{tag}", timeout=(15, 60))
        if found.status_code == 404:
            release = _checked(session.post(f"{api}/releases", json={
                "tag_name": tag, "target_commitish": "main",
                "name": f"SEM trained models — {tag}",
                "body": "Colab trained model weights and aggregate validation reports.",
                "draft": False, "prerelease": False, "make_latest": "false",
            }, timeout=(15, 60)), "Release 생성")
        else:
            release = _checked(found, "Release 조회")
        assets = _checked(session.get(
            f"{api}/releases/{release['id']}/assets", timeout=(15, 60)), "첨부 조회")
        for asset in assets:
            if asset["name"] == zip_path.name:
                if asset.get("state") == "uploaded" and asset["size"] == zip_path.stat().st_size:
                    print("이미 업로드된 번들:", release["html_url"])
                    return release["html_url"]
                raise RuntimeError("동일 이름의 불완전 첨부가 있습니다. Release에서 확인하세요.")
        with zip_path.open("rb") as stream:
            asset = _checked(session.post(
                f"https://uploads.github.com/repos/{repo}/releases/{release['id']}/assets",
                params={"name": zip_path.name}, data=stream,
                headers={"Content-Type": "application/zip"}, timeout=(15, 600),
            ), "모델 ZIP 업로드")
        if asset.get("state") != "uploaded" or asset["size"] != zip_path.stat().st_size:
            raise RuntimeError("서버의 업로드 상태·파일 크기가 일치하지 않습니다.")
        print("GitHub 업로드 완료:", release["html_url"])
        return release["html_url"]
