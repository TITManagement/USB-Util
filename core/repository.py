"""Persistence layer for USB snapshot data."""

from __future__ import annotations

import json
import os
import sys
from typing import List

from .models import UsbDeviceSnapshot


class UsbSnapshotRepository:
    """Handle JSON persistence for USB device snapshots."""

    def __init__(self, json_path: str) -> None:
        self.json_path = json_path

    def load(self) -> List[UsbDeviceSnapshot]:
        """Load snapshots from JSON storage."""
        if not os.path.exists(self.json_path):
            return [self.placeholder("USBデバイス情報が存在しません")]
        try:
            with open(self.json_path, "r", encoding="utf-8") as infile:
                data = json.load(infile)
        except (OSError, json.JSONDecodeError):
            return [self.placeholder("USBデバイス情報の読み込みに失敗しました")]
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list) or not data:
            return [self.placeholder("USBデバイス情報が空です")]
        return [UsbDeviceSnapshot.from_dict(item) for item in data]

    def save(self, snapshots: List[UsbDeviceSnapshot]) -> None:
        """Persist snapshots to JSON storage."""
        try:
            with open(self.json_path, "w", encoding="utf-8") as outfile:
                json.dump(
                    [snapshot.to_dict() for snapshot in snapshots],
                    outfile,
                    ensure_ascii=False,
                    indent=2,
                )
        except OSError as exc:
            print(f"USB情報の書き込みに失敗しました: {exc}", file=sys.stderr)

    @staticmethod
    def placeholder(message: str) -> UsbDeviceSnapshot:
        """Return placeholder snapshot with error information."""
        return UsbDeviceSnapshot(vid="-", pid="-", error=message)
