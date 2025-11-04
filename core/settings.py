from dataclasses import dataclass, asdict
import json
import os
from typing import Dict


DEFAULT_PATH = 'settings.json'


@dataclass
class Settings:
    play_count: int = 1
    infinite: bool = False
    last_action: str = ''

    @staticmethod
    def load(path: str = DEFAULT_PATH) -> 'Settings':
        try:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f) or {}
                    return Settings(
                        play_count=int(data.get('play_count', 1) or 1),
                        infinite=bool(data.get('infinite', False)),
                        last_action=str(data.get('last_action', '') or ''),
                    )
        except Exception:
            pass
        return Settings()

    def save(self, path: str = DEFAULT_PATH) -> None:
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(asdict(self), f, ensure_ascii=False, indent=2)
        except Exception:
            pass


def load_settings(path: str = DEFAULT_PATH) -> Dict:
    s = Settings.load(path)
    return asdict(s)


def save_settings(data: Dict, path: str = DEFAULT_PATH) -> None:
    try:
        s = Settings(
            play_count=int(data.get('play_count', 1) or 1),
            infinite=bool(data.get('infinite', False)),
            last_action=str(data.get('last_action', '') or ''),
        )
        s.save(path)
    except Exception:
        pass


def apply_settings_to_ui(settings: Dict, playCountVar, infiniteVar, actionVar, list_actions_callable):
    try:
        if 'play_count' in settings:
            try:
                playCountVar.set(int(settings.get('play_count', 1)))
            except Exception:
                pass
        if 'infinite' in settings:
            try:
                infiniteVar.set(bool(settings.get('infinite', False)))
            except Exception:
                pass
        if 'last_action' in settings:
            last = str(settings.get('last_action') or '').strip()
            files = list_actions_callable()
            if last and last in files:
                try:
                    actionVar.set(last)
                except Exception:
                    pass
    except Exception:
        pass

