"""Tkinter-based application view for USB-util."""

from __future__ import annotations

import sys
from typing import Dict, List, Optional

import customtkinter as ctk

from .view_model import UsbDevicesViewModel


class UsbDevicesApp:
    """Render USB device snapshots via CustomTkinter."""

    def __init__(self, view_model: UsbDevicesViewModel) -> None:
        self.view_model = view_model
        print(f"[DEBUG] UsbDevicesApp initial snapshots count: {self.view_model.device_count()}")

        self.app = ctk.CTk()
        self.app.title("USB Devices Viewer")
        self.app.geometry("900x600")

        self.combo: Optional[ctk.CTkComboBox] = None
        self.device_count_label: Optional[ctk.CTkLabel] = None
        self.device_listbox: Optional[ctk.CTkScrollableFrame] = None
        self.device_list_items: List[ctk.CTkFrame] = []
        self.info_labels: Dict[str, ctk.CTkLabel] = {}
        self.error_label: Optional[ctk.CTkLabel] = None
        self.detail_box: Optional[ctk.CTkTextbox] = None

        self._ui_spacing = self._spacing_config()
        self._ui_fonts = self._font_config()

        self._setup_layout()
        self._apply_view_model(update_combo=True, rebuild_list=True)

    def run(self) -> None:
        self.app.mainloop()

    def _spacing_config(self) -> dict:
        return {
            "layout": {
                "root": (20, 20),             # ウィンドウ全体の外側マージン
                "content_gap": (10, 5),      # 上部コントロールバーと下部エリアの間隔
            },
            "top_controls": {
                "frame_padx": 10,             # 上部コントロールバーの左右余白
                "frame_pady": (10, 5),        # 上部コントロールバーの上下余白
                "title_padx": 10,             # タイトルラベルと左端の距離
                "combo_padx": 20,             # コンボボックス左右余白
                "button_padx": 10,            # 再読み込みボタン左右余白
            },
            "device_summary": {
                "header_pad": (10, 0),        # 左ペイン見出しの余白
                "frame_padx": 10,             # サマリーリスト左右余白
                "frame_pady": 10,             # サマリーリスト上下余白
                "item_padx": 6,               # サマリー項目枠の左右余白
                "item_pady": 4,               # サマリー項目枠の上下余白
                "label_padx": 8,              # サマリー項目テキスト左右余白
                "label_pady": 6,              # サマリー項目テキスト上下余白
            },
            "device_info": {
                "frame_padx": 10,             # デバイス情報/JSONペインの左右余白
                "frame_pady": 10,             # デバイス情報/JSONペインの上下余白
                "label_padx": (10, 6),        # 詳細ラベルと左端との距離
                "value_padx": (0, 10),        # 詳細値と右端との距離
                "row_pady": (0, 0),           # 詳細項目の行間
                "error_pady": (8, 0),         # エラー表示の上下余白
                "json_padx": 10,              # JSONテキストボックス左右余白
                "json_pady": 10,              # JSONテキストボックス上下余白
            },
        }

    def _font_config(self) -> dict:
        return {
            "top_controls": {"title": ("Meiryo", 20, "bold")},
            "section_heading": ("Meiryo", 14, "bold"),
            "device_summary": {"counter": ("Meiryo", 12)},
            "device_info": {"label": ("Meiryo", 12, "bold"), "value": ("Meiryo", 12), "error": ("Meiryo", 11)},
        }

    # ------------------------------------------------------------------ layout
    def _setup_layout(self) -> None:
        spacing = self._ui_spacing

        layout_spacing = spacing["layout"]
        summary_spacing = spacing["device_summary"]
        info_spacing = spacing["device_info"]

        root_padx, root_pady = layout_spacing["root"]
        main_frame = ctk.CTkFrame(self.app)
        main_frame.pack(fill="both", expand=True, padx=root_padx, pady=root_pady)

        self._setup_top_controls(main_frame)

        bottom_frame = ctk.CTkFrame(main_frame)
        bottom_frame.pack(fill="both", expand=True, padx=info_spacing['frame_padx'], pady=layout_spacing["content_gap"])
        bottom_frame.grid_columnconfigure(0, weight=0)
        bottom_frame.grid_columnconfigure(1, weight=1)
        bottom_frame.grid_columnconfigure(2, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        summary_frame = ctk.CTkFrame(bottom_frame)
        summary_frame.grid(row=0, column=0, sticky="nsew", padx=(0, summary_spacing["frame_padx"]))
        self._setup_summary_section(summary_frame)

        detail_parent = ctk.CTkFrame(bottom_frame)
        detail_parent.grid(row=0, column=1, sticky="nsew", padx=(0, info_spacing['frame_padx']))
        self._setup_detail_section(detail_parent)

        json_parent = ctk.CTkFrame(bottom_frame)
        json_parent.grid(row=0, column=2, sticky="nsew", padx=(info_spacing['frame_padx'], 0))
        self._setup_json_section(json_parent)

    def _setup_top_controls(self, parent: ctk.CTkFrame) -> None:
        controls_spacing = self._ui_spacing["top_controls"]
        controls_fonts = self._ui_fonts["top_controls"]

        frame = ctk.CTkFrame(parent)
        frame.pack(fill="x", padx=controls_spacing["frame_padx"], pady=controls_spacing["frame_pady"])

        title = ctk.CTkLabel(frame, text="USBデバイス選択", font=controls_fonts["title"])
        title.pack(side="left", padx=controls_spacing["title_padx"])

        options = self.view_model.get_options()
        self.combo = ctk.CTkComboBox(frame, values=options, width=300)
        self.combo.pack(side="left", padx=controls_spacing["combo_padx"])
        self.combo.configure(command=self._on_selection_change)

        reload_btn = ctk.CTkButton(frame, text="再読み込み", command=self._reload_snapshots, width=120)
        reload_btn.pack(side="right", padx=controls_spacing["button_padx"])

    def _setup_summary_section(self, parent: ctk.CTkFrame) -> None:
        summary_spacing = self._ui_spacing["device_summary"]
        section_heading_font = self._ui_fonts["section_heading"]
        summary_fonts = self._ui_fonts["device_summary"]

        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=summary_spacing["frame_padx"], pady=summary_spacing["header_pad"])

        device_list_label = ctk.CTkLabel(header, text="USBデバイス一覧", font=section_heading_font)
        device_list_label.pack(side="left")

        self.device_count_label = ctk.CTkLabel(header, text="(0)", font=summary_fonts["counter"])
        self.device_count_label.pack(side="right")

        self.device_listbox = ctk.CTkScrollableFrame(parent, width=260, fg_color=("gray12", "gray18"))
        self.device_listbox.pack(
            fill="both",
            expand=True,
            padx=summary_spacing["frame_padx"],
            pady=summary_spacing["frame_pady"],
        )

    def _setup_detail_section(self, parent: ctk.CTkFrame) -> None:
        info_spacing = self._ui_spacing['device_info']
        section_heading_font = self._ui_fonts['section_heading']
        info_fonts = self._ui_fonts['device_info']

        heading = ctk.CTkLabel(parent, text="デバイス情報", font=section_heading_font)
        heading.pack(anchor="w", padx=info_spacing["frame_padx"], pady=(10, 0))

        info_frame = ctk.CTkScrollableFrame(parent)
        info_frame.pack(fill="both", expand=True, padx=info_spacing["frame_padx"], pady=info_spacing["frame_pady"])
        info_frame.grid_columnconfigure(0, weight=1)

        fields = [
            "VID",
            "usb.ids Vendor",
            "PID",
            "usb.ids Product",
            "Manufacturer",
            "Product",
            "Serial",
            "Identity",
            "Bus",
            "Address",
            "Port Path",
            "Class Guess",
            "COMポート",
            "接続経路",
            "LocationInformation",
        ]

        for row, label in enumerate(fields):
            label_kwargs = {"text": f"{label}: ―", "font": info_fonts["label"], "justify": "left", "anchor": "w"}
            if label == "COMポート":
                label_kwargs["text_color"] = "red"
            row_label = ctk.CTkLabel(info_frame, **label_kwargs)
            row_label.grid(
                row=row,
                column=0,
                sticky="w",
                padx=info_spacing["label_padx"],
                pady=info_spacing["row_pady"],
            )
            self.info_labels[label] = row_label

        self.error_label = ctk.CTkLabel(info_frame, text="", font=info_fonts["error"], text_color="red")
        self.error_label.grid(row=len(fields), column=0, columnspan=2, sticky="w", padx=info_spacing["frame_padx"], pady=info_spacing["error_pady"])

    def _setup_json_section(self, parent: ctk.CTkFrame) -> None:
        info_spacing = self._ui_spacing['device_info']
        section_heading_font = self._ui_fonts['section_heading']
        info_fonts = self._ui_fonts['device_info']

        heading = ctk.CTkLabel(parent, text="デバイス詳細 (JSON)", font=section_heading_font)
        heading.pack(anchor="w", padx=info_spacing["frame_padx"], pady=(10, 0))

        self.detail_box = ctk.CTkTextbox(parent, font=info_fonts["value"])
        self.detail_box.pack(fill="both", expand=True, padx=info_spacing["json_padx"], pady=info_spacing["json_pady"])

    # -------------------------------------------------------------- view syncs
    def _apply_view_model(self, update_combo: bool = True, rebuild_list: bool = True) -> None:
        if update_combo and self.combo:
            options = self.view_model.get_options()
            self.combo.configure(values=options)
            selected = self.view_model.selected_option()
            if selected:
                self.combo.set(selected)
            elif options:
                self.combo.set(options[0])
                self.view_model.select_by_key(options[0])
            else:
                self.combo.set("")

        if rebuild_list:
            self._rebuild_device_list()

        self._update_device_count()
        self._update_detail()

    def _reload_snapshots(self) -> None:
        _snapshots, scan_error = self.view_model.refresh()
        if scan_error:
            print(scan_error, file=sys.stderr)
        self._apply_view_model(update_combo=True, rebuild_list=True)

    def _rebuild_device_list(self) -> None:
        if not self.device_listbox:
            return
        for item in self.device_list_items:
            item.destroy()
        self.device_list_items = []

        entries = self.view_model.list_entries()
        selected_index = self.view_model.selected_index
        for idx, (text, dimmed) in enumerate(entries):
            item_frame = ctk.CTkFrame(self.device_listbox, corner_radius=8)
            if idx == selected_index:
                item_frame.configure(fg_color=("#50a7ff", "#2b82d9"))
            elif dimmed:
                item_frame.configure(fg_color=("#6db6ff", "#5a92cc"))
            else:
                item_frame.configure(fg_color=("#9fd3ff", "#75b5f2"))
            item_frame.pack(fill="x", padx=self._ui_spacing['device_summary']['item_padx'], pady=self._ui_spacing['device_summary']['item_pady'])

            text_color = ("black", "black") if not dimmed else ("#1f3c66", "#0f2845")
            item_label = ctk.CTkLabel(item_frame, text=text, justify="left", anchor="w", text_color=text_color)
            item_label.pack(fill="x", padx=self._ui_spacing['device_summary']['label_padx'], pady=self._ui_spacing['device_summary']['label_pady'])

            item_label.bind("<Button-1>", lambda _evt, index=idx: self._on_device_list_select(index))
            item_frame.bind("<Button-1>", lambda _evt, index=idx: self._on_device_list_select(index))
            self.device_list_items.append(item_frame)

    def _update_device_count(self) -> None:
        if self.device_count_label:
            self.device_count_label.configure(text=f"({self.view_model.device_count()})")

    def _update_detail(self) -> None:
        info_values = self.view_model.info_values()
        for key, widget in self.info_labels.items():
            value_text = info_values.get(key, "―")
            widget.configure(text=f"{key}: {value_text}")

        snapshot = self.view_model.current_snapshot()
        if self.error_label:
            if snapshot and snapshot.error:
                self.error_label.configure(text=f"Error: {snapshot.error}")
            else:
                self.error_label.configure(text="")

        if self.detail_box:
            self.detail_box.configure(state="normal")
            self.detail_box.delete("1.0", "end")
            self.detail_box.insert("end", self.view_model.detail_json())
            self.detail_box.configure(state="disabled")

    # ---------------------------------------------------------------- handlers
    def _on_device_list_select(self, index: int) -> None:
        self.view_model.select_by_index(index)
        if self.combo:
            self.combo.set(self.view_model.selected_option())
        self._apply_view_model(update_combo=False, rebuild_list=True)

    def _on_selection_change(self, _: Optional[str] = None) -> None:
        if not self.combo:
            return
        self.view_model.select_by_key(self.combo.get())
        self._apply_view_model(update_combo=False, rebuild_list=True)
