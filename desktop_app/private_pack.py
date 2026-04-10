from __future__ import annotations

import base64
import hashlib
import json
import os
import shutil
import tempfile
import time
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.scrypt import Scrypt


PACK_EXTENSION = ".amspack"
PACK_METADATA_FILENAME = "pack.json"
PACK_PAYLOAD_FILENAME = "payload.bin"
INNER_MANIFEST_FILENAME = "private-pack.json"
PACK_FORMAT_VERSION = 1
PACK_MAGIC = "AMS_PRIVATE_PACK"
SALT_BYTES = 16
NONCE_BYTES = 12
DERIVED_KEY_BYTES = 32
SCRYPT_N = 2**14
SCRYPT_R = 8
SCRYPT_P = 1

FEATURE_PATHS = {
    "contract_templates": Path("maritime-service") / "assets" / "contract_templates",
    "contract_registry": Path("maritime-service") / "scripts" / "contract_template_registry.json",
    "clearance_site_config": Path("maritime-service") / "scripts" / "clearance_site_config.json",
    "help_overrides": Path("desktop_app") / "release_assets" / "help",
}


class PrivatePackError(Exception):
    pass


@dataclass
class PrivatePackSummary:
    pack_id: str
    display_name: str
    version: str
    created_at: str
    format_version: int
    features: list[str]
    description: str = ""
    source_name: str = ""
    payload_sha256: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class PrivatePackInstallResult:
    summary: PrivatePackSummary
    manifest: dict[str, Any]
    installed_at: str
    unpacked_root: str

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["summary"] = self.summary.to_dict()
        return payload


def _now_iso() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def _encode_b64(raw: bytes) -> str:
    return base64.b64encode(raw).decode("ascii")


def _decode_b64(value: str) -> bytes:
    return base64.b64decode(value.encode("ascii"))


def _derive_key(password: str, salt: bytes) -> bytes:
    if not password:
        raise PrivatePackError("密码不能为空。")
    kdf = Scrypt(
        salt=salt,
        length=DERIVED_KEY_BYTES,
        n=SCRYPT_N,
        r=SCRYPT_R,
        p=SCRYPT_P,
    )
    return kdf.derive(password.encode("utf-8"))


def _manifest_path(root: Path) -> Path:
    return root / INNER_MANIFEST_FILENAME


def _iter_payload_files(source_dir: Path) -> list[Path]:
    files = [
        path
        for path in source_dir.rglob("*")
        if path.is_file() and path.name != PACK_METADATA_FILENAME and "__pycache__" not in path.parts
    ]
    return sorted(files, key=lambda item: item.relative_to(source_dir).as_posix())


def _normalize_manifest(manifest: dict[str, Any], source_dir: Path, features: list[str]) -> dict[str, Any]:
    pack_id = str(manifest.get("pack_id") or source_dir.name).strip()
    display_name = str(manifest.get("display_name") or pack_id).strip()
    version = str(manifest.get("version") or "0.0.0").strip()
    description = str(manifest.get("description") or "").strip()
    normalized = dict(manifest)
    normalized["pack_id"] = pack_id
    normalized["display_name"] = display_name
    normalized["version"] = version
    normalized["description"] = description
    normalized["features"] = features
    normalized.setdefault("created_at", _now_iso())
    return normalized


def detect_private_pack_features(root: Path) -> list[str]:
    features: list[str] = []
    for feature_name, relative_path in FEATURE_PATHS.items():
        if (root / relative_path).exists():
            features.append(feature_name)
    return features


def validate_private_pack_tree(root: Path) -> dict[str, Any]:
    root = root.resolve()
    manifest_path = _manifest_path(root)
    if not manifest_path.exists():
        raise PrivatePackError(f"私密包目录缺少 {INNER_MANIFEST_FILENAME}。")
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise PrivatePackError(f"{INNER_MANIFEST_FILENAME} 不是合法 JSON。") from exc
    features = detect_private_pack_features(root)
    if not features:
        raise PrivatePackError("私密包里没有找到可覆盖的资源。至少需要合同模板、合同注册表、站点配置或帮助页中的一种。")
    normalized_manifest = _normalize_manifest(manifest, root, features)
    return {
        "manifest": normalized_manifest,
        "features": features,
    }


def read_private_pack_summary(pack_path: Path | str) -> PrivatePackSummary:
    pack_path = Path(pack_path).expanduser().resolve()
    if not pack_path.exists():
        raise FileNotFoundError(f"Private pack not found: {pack_path}")
    with zipfile.ZipFile(pack_path, "r") as archive:
        try:
            metadata = json.loads(archive.read(PACK_METADATA_FILENAME).decode("utf-8"))
        except KeyError as exc:
            raise PrivatePackError("私密包缺少 pack.json，无法读取摘要。") from exc
    summary = metadata.get("summary") or {}
    return PrivatePackSummary(
        pack_id=str(summary.get("pack_id") or "").strip(),
        display_name=str(summary.get("display_name") or pack_path.stem).strip(),
        version=str(summary.get("version") or "").strip(),
        created_at=str(summary.get("created_at") or "").strip(),
        format_version=int(metadata.get("format_version") or 0),
        features=list(summary.get("features") or []),
        description=str(summary.get("description") or "").strip(),
        source_name=pack_path.name,
        payload_sha256=str(metadata.get("payload_sha256") or "").strip(),
    )


def build_private_pack(source_dir: Path | str, output_path: Path | str, password: str) -> PrivatePackSummary:
    source_dir = Path(source_dir).expanduser().resolve()
    output_path = Path(output_path).expanduser().resolve()
    if not source_dir.exists():
        raise FileNotFoundError(f"Source directory not found: {source_dir}")
    validation = validate_private_pack_tree(source_dir)
    manifest = validation["manifest"]
    summary = PrivatePackSummary(
        pack_id=manifest["pack_id"],
        display_name=manifest["display_name"],
        version=manifest["version"],
        created_at=manifest["created_at"],
        format_version=PACK_FORMAT_VERSION,
        features=list(validation["features"]),
        description=manifest.get("description", ""),
        source_name=output_path.name,
    )

    payload_sha256 = hashlib.sha256()
    with tempfile.TemporaryDirectory(prefix="ams-private-pack-") as temp_dir:
        payload_zip_path = Path(temp_dir) / "payload.zip"
        with zipfile.ZipFile(payload_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as payload_zip:
            for file_path in _iter_payload_files(source_dir):
                relative = file_path.relative_to(source_dir).as_posix()
                payload_zip.write(file_path, arcname=relative)
        payload_bytes = payload_zip_path.read_bytes()
        payload_sha256.update(payload_bytes)
        summary.payload_sha256 = payload_sha256.hexdigest()

        salt = os.urandom(SALT_BYTES)
        nonce = os.urandom(NONCE_BYTES)
        key = _derive_key(password, salt)
        metadata = {
            "magic": PACK_MAGIC,
            "format_version": PACK_FORMAT_VERSION,
            "kdf": {
                "name": "scrypt",
                "n": SCRYPT_N,
                "r": SCRYPT_R,
                "p": SCRYPT_P,
                "length": DERIVED_KEY_BYTES,
                "salt_b64": _encode_b64(salt),
            },
            "cipher": {
                "name": "AES-256-GCM",
                "nonce_b64": _encode_b64(nonce),
            },
            "payload_sha256": summary.payload_sha256,
            "summary": summary.to_dict(),
        }
        aad = json.dumps(
            {
                "magic": PACK_MAGIC,
                "format_version": PACK_FORMAT_VERSION,
                "pack_id": summary.pack_id,
                "version": summary.version,
            },
            ensure_ascii=False,
            sort_keys=True,
        ).encode("utf-8")
        ciphertext = AESGCM(key).encrypt(nonce, payload_bytes, aad)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(PACK_METADATA_FILENAME, json.dumps(metadata, ensure_ascii=False, indent=2))
            archive.writestr(PACK_PAYLOAD_FILENAME, ciphertext)
    return summary


def unpack_private_pack(pack_path: Path | str, password: str, destination_root: Path | str) -> PrivatePackInstallResult:
    pack_path = Path(pack_path).expanduser().resolve()
    destination_root = Path(destination_root).expanduser().resolve()
    summary = read_private_pack_summary(pack_path)

    with zipfile.ZipFile(pack_path, "r") as archive:
        metadata = json.loads(archive.read(PACK_METADATA_FILENAME).decode("utf-8"))
        ciphertext = archive.read(PACK_PAYLOAD_FILENAME)

    if metadata.get("magic") != PACK_MAGIC:
        raise PrivatePackError("不是合法的 AMS 私密业务包。")
    if int(metadata.get("format_version") or 0) != PACK_FORMAT_VERSION:
        raise PrivatePackError("私密包格式版本不受支持。")

    salt = _decode_b64(metadata["kdf"]["salt_b64"])
    nonce = _decode_b64(metadata["cipher"]["nonce_b64"])
    aad = json.dumps(
        {
            "magic": PACK_MAGIC,
            "format_version": PACK_FORMAT_VERSION,
            "pack_id": summary.pack_id,
            "version": summary.version,
        },
        ensure_ascii=False,
        sort_keys=True,
    ).encode("utf-8")
    key = _derive_key(password, salt)
    try:
        payload_bytes = AESGCM(key).decrypt(nonce, ciphertext, aad)
    except InvalidTag as exc:
        raise PrivatePackError("密码不正确，或者私密包已损坏。") from exc

    payload_hash = hashlib.sha256(payload_bytes).hexdigest()
    if summary.payload_sha256 and payload_hash != summary.payload_sha256:
        raise PrivatePackError("私密包校验失败，内容可能已损坏。")

    destination_root.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ams-private-pack-unpack-") as temp_dir:
        temp_root = Path(temp_dir) / "payload"
        temp_root.mkdir(parents=True, exist_ok=True)
        payload_zip_path = Path(temp_dir) / "payload.zip"
        payload_zip_path.write_bytes(payload_bytes)
        with zipfile.ZipFile(payload_zip_path, "r") as payload_zip:
            payload_zip.extractall(temp_root)
        validation = validate_private_pack_tree(temp_root)
        if destination_root.exists():
            shutil.rmtree(destination_root)
        shutil.copytree(temp_root, destination_root)

    installed_at = _now_iso()
    metadata_payload = {
        "installed_at": installed_at,
        "source_pack_path": str(pack_path),
        "source_pack_name": pack_path.name,
        "summary": summary.to_dict(),
        "manifest": validation["manifest"],
    }
    (destination_root / ".installed-pack.json").write_text(
        json.dumps(metadata_payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return PrivatePackInstallResult(
        summary=summary,
        manifest=validation["manifest"],
        installed_at=installed_at,
        unpacked_root=str(destination_root),
    )


def load_installed_private_pack(root: Path | str) -> dict[str, Any] | None:
    root = Path(root).expanduser().resolve()
    state_path = root / ".installed-pack.json"
    if not state_path.exists():
        manifest_path = _manifest_path(root)
        if not manifest_path.exists():
            return None
        validation = validate_private_pack_tree(root)
        manifest = validation["manifest"]
        return {
            "installed_at": "",
            "source_pack_path": "",
            "source_pack_name": "",
            "summary": {
                "pack_id": manifest["pack_id"],
                "display_name": manifest["display_name"],
                "version": manifest["version"],
                "created_at": manifest.get("created_at", ""),
                "format_version": PACK_FORMAT_VERSION,
                "features": validation["features"],
                "description": manifest.get("description", ""),
                "source_name": "",
                "payload_sha256": "",
            },
            "manifest": manifest,
        }
    return json.loads(state_path.read_text(encoding="utf-8"))
