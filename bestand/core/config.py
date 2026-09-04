"""Konfiguration der Bestandsliste - rein, ohne os.environ, dotenv oder argparse.

Die CLI und das Dashboard laden dieselbe ``config.json``; wer Werte ueberschreiben
will (CLI-Flags, Dashboard-Einstellungen), reicht sie als Schluesselwortargumente
durch. Fehler kommen als :class:`ConfigError` zurueck, nie als ``SystemExit`` -
ein Webserver darf sich davon nicht beenden lassen.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


class ConfigError(ValueError):
    """Die Konfiguration ist unbrauchbar; der Aufrufer entscheidet ueber die Folge."""


@dataclass(frozen=True)
class BestandConfig:
    excel_path: Path
    sheet_name: str
    safety_stock: int = 5
    match_overrides: dict[str, str] = field(default_factory=dict)

    @classmethod
    def load(
        cls,
        config_path: Path,
        *,
        excel: str | Path | None = None,
        sheet: str | None = None,
        safety_stock: int | None = None,
    ) -> "BestandConfig":
        """Liest config.json neben dem Skript; Argumente haben Vorrang.

        Ein relativer ``excel``-Wert wird gegen das Verzeichnis der config.json
        aufgeloest, ein absoluter unveraendert uebernommen.
        """
        config_path = Path(config_path)
        try:
            with open(config_path, encoding="utf-8") as handle:
                raw = json.load(handle)
        except FileNotFoundError as exc:
            raise ConfigError(f"config.json nicht gefunden: {config_path}") from exc
        except json.JSONDecodeError as exc:
            raise ConfigError(f"config.json ist kein gueltiges JSON: {exc}") from exc

        excel_value = excel if excel is not None else raw.get("excel_file")
        if not excel_value:
            raise ConfigError("config.json: 'excel_file' fehlt.")
        sheet_name = sheet if sheet is not None else raw.get("sheet_name")
        if not sheet_name:
            raise ConfigError("config.json: 'sheet_name' fehlt.")

        stock = safety_stock if safety_stock is not None else raw.get("safety_stock", 5)
        if not isinstance(stock, int) or isinstance(stock, bool) or stock < 0:
            raise ConfigError("Sicherheitsbestand muss eine Zahl >= 0 sein.")

        overrides = raw.get("match_overrides", {})
        if not isinstance(overrides, dict) or not all(
            isinstance(key, str) and isinstance(value, str) for key, value in overrides.items()
        ):
            raise ConfigError(
                "config.json: match_overrides muss ein Objekt aus String-Schluesseln "
                "und ISBN-Strings sein."
            )

        return cls(
            excel_path=config_path.parent / excel_value,
            sheet_name=str(sheet_name),
            safety_stock=stock,
            match_overrides=dict(overrides),
        )
