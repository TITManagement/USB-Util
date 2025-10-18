"""Domain models for USB-util."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover - circular import guard
    from main import UsbIdsDatabase  # type: ignore


@dataclass
class UsbDeviceSnapshot:
    """Serializable snapshot of a USB device."""

    vid: str
    pid: str
    manufacturer: str = ""
    product: str = ""
    serial: str = ""
    bus: Optional[int] = None
    address: Optional[int] = None
    port_path: List[int] = field(default_factory=list)
    device_descriptor: Dict[str, Any] = field(default_factory=dict)
    configurations: List[Dict[str, Any]] = field(default_factory=list)
    class_guess: str = "-"
    error: Optional[str] = None
    topology_chain: List[str] = field(default_factory=list)
    location_information: str = ""
    location_fallback: str = ""
    usb_controllers: List[str] = field(default_factory=list)

    def key(self) -> str:
        """Return canonical VID/PID pair identifier."""
        return f"{self.vid}:{self.pid}"

    def identity(self) -> str:
        """Return a cross-platform identifier to disambiguate identical devices."""
        parts: List[str] = []
        serial = self.serial.strip() if isinstance(self.serial, str) else ""
        if serial and serial not in {"取得不可", "-"}:
            parts.append(f"SER:{serial}")
        elif self.port_path:
            path = "-".join(str(p) for p in self.port_path)
            parts.append(f"PORT:{path}")
        if self.bus is not None:
            parts.append(f"BUS:{self.bus}")
        if self.address is not None:
            parts.append(f"ADDR:{self.address}")
        if not parts:
            parts.append("UNIDENTIFIED")
        parts.append(f"VIDPID:{self.vid}:{self.pid}")
        return " | ".join(parts)

    def resolve_names(self, ids_db: "UsbIdsDatabase") -> Tuple[str, str]:
        """Resolve vendor/product names using usb.ids database."""
        if self.error:
            return "取得不可", "取得不可"
        if (
            isinstance(self.vid, str)
            and isinstance(self.pid, str)
            and self.vid.startswith("0x")
            and self.pid.startswith("0x")
        ):
            vendor_name, product_name = ids_db.lookup(self.vid, self.pid)
            return vendor_name or "不明", product_name or "不明"
        return "不明", "不明"

    def to_dict(self) -> Dict[str, Any]:
        """Serialize snapshot to dictionary."""
        return {
            "vid": self.vid,
            "pid": self.pid,
            "manufacturer": self.manufacturer,
            "product": self.product,
            "serial": self.serial,
            "bus": self.bus,
            "address": self.address,
            "port_path": self.port_path,
            "device_descriptor": self.device_descriptor,
            "configurations": self.configurations,
            "class_guess": self.class_guess,
            "error": self.error,
            "topology_chain": self.topology_chain,
            "location_information": self.location_information,
            "location_fallback": self.location_fallback,
            "usb_controllers": self.usb_controllers,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "UsbDeviceSnapshot":
        """Deserialize snapshot from dictionary."""
        return cls(
            vid=data.get("vid", "-"),
            pid=data.get("pid", "-"),
            manufacturer=data.get("manufacturer", ""),
            product=data.get("product", ""),
            serial=data.get("serial", ""),
            bus=data.get("bus"),
            address=data.get("address"),
            port_path=data.get("port_path", []),
            device_descriptor=data.get("device_descriptor", {}),
            configurations=data.get("configurations", []),
            class_guess=data.get("class_guess", "-"),
            error=data.get("error"),
            topology_chain=data.get("topology_chain", []),
            location_information=data.get("location_information", ""),
            location_fallback=data.get("location_fallback", ""),
            usb_controllers=data.get("usb_controllers", []),
        )
