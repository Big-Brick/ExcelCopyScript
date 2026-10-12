"""User-selectable, atomic profiles containing workbook layout and aliases."""

from __future__ import annotations

import json
import os
import shutil
import tempfile
from pathlib import Path

from .migrations import CURRENT_VERSION, migrate_alias_data
from .profile_migrations import PROFILE_VERSION, build_profile_from_legacy
from .workbook_layout import WorkbookLayout


CONFIG_DIR = Path.home() / ".config" / "excel-copy-script"


def write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", dir=path.parent,
            prefix=".profile-", suffix=".tmp", delete=False,
        ) as stream:
            tmp = Path(stream.name)
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        if path.exists():
            os.chmod(tmp, path.stat().st_mode)
        os.replace(tmp, path)
    finally:
        if tmp is not None:
            tmp.unlink(missing_ok=True)


class Profile:
    def __init__(self, path: Path, data: dict):
        self.path = Path(path).expanduser().resolve()
        if not isinstance(data, dict) or data.get("profile_version") != PROFILE_VERSION:
            raise ValueError(f"Невідома версія профілю: {data.get('profile_version') if isinstance(data, dict) else '?'}")
        self.layout = WorkbookLayout.from_dict(data.get("layout"))
        aliases = data.get("aliases")
        if not isinstance(aliases, dict):
            raise ValueError("Профіль не містить словника аліасів")
        self.aliases_data, migrated = migrate_alias_data(aliases)
        self.data = {
            "profile_version": PROFILE_VERSION,
            "layout": self.layout.to_dict(),
            "aliases": self.aliases_data,
        }
        self._needs_save = migrated

    @classmethod
    def create_empty(
        cls, path: Path, layout: WorkbookLayout | None = None,
    ) -> "Profile":
        """Create an independent profile with no aliases, without replacing files.

        Unlike save(), creation is exclusive: two concurrent windows cannot
        overwrite an existing profile even if both checked its path first.
        """
        config = layout if layout is not None else WorkbookLayout()
        profile = cls(path, {
            "profile_version": PROFILE_VERSION,
            "layout": config.to_dict(),
            "aliases": {"version": CURRENT_VERSION, "templates": {}},
        })
        profile.path.parent.mkdir(parents=True, exist_ok=True)
        created = False
        try:
            with profile.path.open("x", encoding="utf-8") as stream:
                created = True
                json.dump(profile.data, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
        except Exception:
            if created:
                profile.path.unlink(missing_ok=True)
            raise
        return profile

    @classmethod
    def load(cls, path: Path) -> "Profile":
        path = Path(path).expanduser()
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise ValueError(f"Не вдалося відкрити профіль «{path}»: {exc}") from exc
        profile = cls(path, raw)
        if profile._needs_save:
            backup = profile.path.with_suffix(profile.path.suffix + ".before-migration.bak")
            if not backup.exists():
                shutil.copy2(profile.path, backup)
            profile.save()
        return profile

    @classmethod
    def migrate_legacy(cls, path: Path, legacy: Path) -> "Profile":
        # A failure must not erase an existing profile or touch old aliases.
        if Path(path).exists():
            return cls.load(path)
        data = build_profile_from_legacy(legacy)
        profile = cls(path, data)
        profile.save()
        return profile

    def save(self) -> None:
        self.data["layout"] = self.layout.to_dict()
        self.data["aliases"] = self.aliases_data
        write_json(self.path, self.data)

    def copy_as(self, path: Path) -> "Profile":
        destination = Path(path).expanduser().resolve()
        if destination == self.path:
            raise ValueError("Виберіть інше ім'я для копії профілю")
        if destination.exists():
            raise FileExistsError(f"Профіль «{destination}» вже існує")
        clone = Profile(destination, json.loads(json.dumps(self.data, ensure_ascii=False)))
        clone.save()
        return clone


class ProfileManager:
    def __init__(self, config_dir: Path | None = None):
        self.config_dir = Path(config_dir) if config_dir is not None else CONFIG_DIR
        self.state_path = self.config_dir / "profile_state.json"
        self.default_path = self.config_dir / "profiles" / "default.json"
        self.legacy_path = self.config_dir / "aliases.json"

    def load_last(self) -> "Profile":
        if self.state_path.exists():
            try:
                state = json.loads(self.state_path.read_text(encoding="utf-8"))
                if state.get("version") != 1 or not isinstance(state.get("last_profile"), str):
                    raise ValueError("Невірний формат списку останніх профілів")
                return Profile.load(Path(state["last_profile"]))
            except (OSError, ValueError) as exc:
                # The caller shows the error; never replace or overwrite the
                # last profile silently if it is missing/corrupted.
                raise ValueError(f"Не вдалося завантажити останній профіль: {exc}") from exc
        return Profile.migrate_legacy(self.default_path, self.legacy_path)

    def remember(self, profile: Profile) -> None:
        write_json(self.state_path, {
            "version": 1, "last_profile": str(profile.path),
        })
