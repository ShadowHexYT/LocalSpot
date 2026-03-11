from array import array
import base64
import csv
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from mutagen.easyid3 import EasyID3
    from mutagen.id3 import APIC, ID3, ID3NoHeaderError
    from mutagen.mp3 import MP3
    MUTAGEN_AVAILABLE = True
except Exception:
    MUTAGEN_AVAILABLE = False

APP_TITLE = "Spotify Local Files Pipeline"
DEFAULT_IMPORT_FOLDER = str(Path.home() / "Music" / "Spotify Local Imports")
DEFAULT_SPOTIFY_FOLDER = str(Path.home() / "Music" / "Spotify Ready")
SETTINGS_FILE = Path.home() / ".spotify_local_pipeline_settings.json"
AUDIO_EXTENSIONS = {".mp3", ".m4a", ".flac", ".wav", ".ogg"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}
INVALID_FILENAME_CHARS = r'[<>:"/\\|?*]+'


def clean_filename(text: str, fallback: str = "Unknown") -> str:
    text = (text or "").strip()
    text = re.sub(INVALID_FILENAME_CHARS, " ", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text or fallback


def guess_image_mime(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if suffix == ".png":
        return "image/png"
    if suffix == ".webp":
        return "image/webp"
    if suffix == ".bmp":
        return "image/bmp"
    return "image/jpeg"


class MetadataEditor(tk.Toplevel):
    def __init__(self, parent, rows, on_save):
        super().__init__(parent)
        self.title("Review Metadata Before Import")
        self.geometry("1320x620")
        self.minsize(1180, 500)
        self.rows = rows
        self.on_save = on_save
        self.tree = None
        self._build_ui()
        self.transient(parent)
        self.grab_set()
        self.focus_force()

    def _build_ui(self):
        container = ttk.Frame(self, padding=12)
        container.pack(fill="both", expand=True)
        ttk.Label(
            container,
            text="Review and edit metadata. Double-click a cell to change it. You can fully rename songs, artists, albums, track numbers, playlist sections, and whether the item should be grouped as a playlist import or a single.",
            wraplength=1260,
        ).pack(anchor="w", pady=(0, 8))
        columns = ("filename", "title", "artist", "album", "track", "playlist", "import_type")
        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        widths = {"filename": 250, "title": 220, "artist": 170, "album": 170, "track": 70, "playlist": 220, "import_type": 110}
        labels = {"filename": "File", "title": "Title", "artist": "Artist", "album": "Album", "track": "Track #", "playlist": "Playlist / Section", "import_type": "Type"}
        for col in columns:
            self.tree.heading(col, text=labels[col])
            self.tree.column(col, width=widths[col], anchor="w")
        ybar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        xbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        for index, row in enumerate(self.rows):
            iid = str(index)
            self.tree.insert("", "end", iid=iid, values=(row.get("filename", ""), row.get("title", ""), row.get("artist", ""), row.get("album", ""), row.get("track", ""), row.get("playlist", ""), row.get("import_type", "playlist")))
        self.tree.bind("<Double-1>", self._begin_edit)
        actions = ttk.Frame(container)
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Auto-clean Text", command=self._auto_clean).pack(side="left")
        ttk.Button(actions, text="Fill Missing Track Numbers", command=self._fill_track_numbers).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Mark All as Playlist Import", command=lambda: self._set_import_type_for_all("playlist")).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Mark All as Singles", command=lambda: self._set_import_type_for_all("single")).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Save Changes", command=self._save).pack(side="right")
        ttk.Button(actions, text="Cancel", command=self.destroy).pack(side="right", padx=(0, 8))

    def _auto_clean(self):
        for iid in self.tree.get_children():
            values = list(self.tree.item(iid, "values"))
            for idx in range(1, 6):
                values[idx] = clean_filename(values[idx], values[idx] or "Unknown")
            values[6] = values[6] if values[6] in {"playlist", "single"} else "playlist"
            self.tree.item(iid, values=values)

    def _fill_track_numbers(self):
        for idx, iid in enumerate(self.tree.get_children(), start=1):
            values = list(self.tree.item(iid, "values"))
            if not str(values[4]).strip():
                values[4] = str(idx)
            self.tree.item(iid, values=values)

    def _set_import_type_for_all(self, import_type):
        for iid in self.tree.get_children():
            values = list(self.tree.item(iid, "values"))
            values[6] = import_type
            self.tree.item(iid, values=values)

    def _begin_edit(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        x, y, width, height = self.tree.bbox(row_id, column_id)
        col_index = int(column_id.replace("#", "")) - 1
        current_values = list(self.tree.item(row_id, "values"))
        if col_index == 6:
            combo = ttk.Combobox(self.tree, values=["playlist", "single"], state="readonly")
            combo.place(x=x, y=y, width=width, height=height)
            combo.set(current_values[col_index] if current_values[col_index] in {"playlist", "single"} else "playlist")
            combo.focus()

            def save_combo(_event=None):
                current_values[col_index] = combo.get().strip() or "playlist"
                self.tree.item(row_id, values=current_values)
                combo.destroy()

            combo.bind("<<ComboboxSelected>>", save_combo)
            combo.bind("<FocusOut>", save_combo)
            combo.bind("<Escape>", lambda _e: combo.destroy())
            return
        entry = ttk.Entry(self.tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_values[col_index])
        entry.focus()
        entry.select_range(0, "end")

        def save_edit(_event=None):
            current_values[col_index] = entry.get().strip()
            self.tree.item(row_id, values=current_values)
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", lambda _e: entry.destroy())

    def _save(self):
        updated_rows = []
        for iid in self.tree.get_children():
            values = list(self.tree.item(iid, "values"))
            updated_rows.append({"filename": values[0], "title": values[1], "artist": values[2], "album": values[3], "track": values[4], "playlist": values[5], "import_type": values[6] if values[6] in {"playlist", "single"} else "playlist"})
        self.on_save(updated_rows)
        self.destroy()


class OutputFileManager(tk.Toplevel):
    def __init__(self, parent, rows, on_delete, on_trim):
        super().__init__(parent)
        self.title("Manage Created Files")
        self.geometry("1420x640")
        self.minsize(1240, 520)
        self.rows = rows
        self.on_delete = on_delete
        self.on_trim = on_trim
        self.tree = None
        self._build_ui()
        self.transient(parent)
        self.grab_set()
        self.focus_force()

    def _build_ui(self):
        container = ttk.Frame(self, padding=12)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text="Review the files created by the pipeline. Select any tracks you do not want and delete them from disk.", wraplength=1360).pack(anchor="w", pady=(0, 8))
        columns = ("title", "artist", "album", "playlist", "status", "spotify_path", "source_path")
        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        widths = {"title": 220, "artist": 160, "album": 180, "playlist": 180, "status": 110, "spotify_path": 360, "source_path": 360}
        labels = {"title": "Title", "artist": "Artist", "album": "Album", "playlist": "Playlist", "status": "Status", "spotify_path": "Spotify File", "source_path": "Downloaded File"}
        for col in columns:
            self.tree.heading(col, text=labels[col])
            self.tree.column(col, width=widths[col], anchor="w")
        ybar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        xbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        for index, row in enumerate(self.rows):
            iid = str(index)
            spotify_path = row.get("spotify_path", "")
            source_path = row.get("source_path", "")
            status = []
            if spotify_path and Path(spotify_path).exists():
                status.append("spotify")
            if source_path and Path(source_path).exists():
                status.append("download")
            self.tree.insert("", "end", iid=iid, values=(row.get("title", ""), row.get("artist", ""), row.get("album", ""), row.get("playlist", ""), " + ".join(status) if status else "missing", spotify_path, source_path))
        actions = ttk.Frame(container)
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Trim Selected", command=self._trim_selected).pack(side="left")
        ttk.Button(actions, text="Delete Selected", command=self._delete_selected).pack(side="left")
        ttk.Button(actions, text="Close", command=self.destroy).pack(side="right")

    def _trim_selected(self):
        selected = self.tree.selection()
        if len(selected) != 1:
            messagebox.showinfo("Manage Created Files", "Select exactly one file to trim.")
            return
        self.on_trim(dict(self.rows[int(selected[0])]))

    def _delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Manage Created Files", "Select one or more files first.")
            return
        rows_to_delete = [self.rows[int(iid)] for iid in selected]
        self.on_delete(rows_to_delete)
        for iid in selected:
            self.tree.delete(iid)


class AudioTrimEditor(tk.Toplevel):
    def __init__(self, parent, row, ffmpeg_path, on_save, logger):
        super().__init__(parent)
        self.title("Trim Audio")
        self.geometry("1220x760")
        self.minsize(1060, 640)
        self.row = row
        self.ffmpeg_path = ffmpeg_path
        self.on_save = on_save
        self.logger = logger
        self.audio_path = Path(row.get("source_path") or row.get("spotify_path") or "")
        self.duration = 0.0
        self.waveform = []
        self.canvas_width = 760
        self.canvas_height = 220
        self.drag_target = None
        self.start_var = tk.DoubleVar(value=0.0)
        self.end_var = tk.DoubleVar(value=0.0)
        self.status_var = tk.StringVar(value="Loading waveform...")
        self.selection_var = tk.StringVar(value="Selection: 0.00s - 0.00s")
        self.segment_tree = None
        self.segments = []
        self._build_ui()
        self.transient(parent)
        self.grab_set()
        self.focus_force()
        self.after(10, self._load_waveform)

    def _build_ui(self):
        container = ttk.Frame(self, padding=12)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)
        title = self.row.get("title") or self.audio_path.name
        ttk.Label(container, text=f"Trim: {title}", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(container, text="Drag the blue handles or use the sliders to define a clip, add it to the list, then edit titles and save all clips as separate files.", wraplength=1120).pack(anchor="w", pady=(4, 10))
        self.canvas = tk.Canvas(container, width=self.canvas_width, height=self.canvas_height, background="#101418", highlightthickness=1, highlightbackground="#405060")
        self.canvas.pack(fill="x")
        self.canvas.bind("<Button-1>", self._on_canvas_press)
        self.canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        ttk.Label(container, textvariable=self.status_var).pack(anchor="w", pady=(8, 0))
        ttk.Label(container, textvariable=self.selection_var).pack(anchor="w", pady=(2, 8))
        slider_frame = ttk.Frame(container)
        slider_frame.pack(fill="x", pady=(6, 0))
        slider_frame.columnconfigure(1, weight=1)
        ttk.Label(slider_frame, text="Start").grid(row=0, column=0, sticky="w")
        self.start_scale = tk.Scale(slider_frame, from_=0, to=0, orient="horizontal", resolution=0.1, variable=self.start_var, command=lambda _v: self._on_scale_change("start"))
        self.start_scale.grid(row=0, column=1, sticky="ew")
        ttk.Label(slider_frame, text="End").grid(row=1, column=0, sticky="w")
        self.end_scale = tk.Scale(slider_frame, from_=0, to=0, orient="horizontal", resolution=0.1, variable=self.end_var, command=lambda _v: self._on_scale_change("end"))
        self.end_scale.grid(row=1, column=1, sticky="ew")
        clip_actions = ttk.Frame(container)
        clip_actions.pack(fill="x", pady=(8, 8))
        ttk.Button(clip_actions, text="Add Current Clip", command=self._add_current_segment).pack(side="left")
        ttk.Button(clip_actions, text="Update Selected Clip", command=self._update_selected_segment).pack(side="left", padx=(8, 0))
        ttk.Button(clip_actions, text="Remove Selected Clip", command=self._remove_selected_segment).pack(side="left", padx=(8, 0))
        ttk.Button(clip_actions, text="Reset Selection", command=self._reset_selection).pack(side="left", padx=(8, 0))

        columns = ("start", "end", "title", "artist", "album", "track", "playlist", "import_type")
        tree_frame = ttk.Frame(container)
        tree_frame.pack(fill="both", expand=True, pady=(0, 8))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.segment_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        widths = {"start": 80, "end": 80, "title": 220, "artist": 150, "album": 160, "track": 70, "playlist": 200, "import_type": 100}
        labels = {"start": "Start", "end": "End", "title": "Title", "artist": "Artist", "album": "Album", "track": "Track #", "playlist": "Playlist", "import_type": "Type"}
        for col in columns:
            self.segment_tree.heading(col, text=labels[col])
            self.segment_tree.column(col, width=widths[col], anchor="w")
        ybar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.segment_tree.yview)
        xbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.segment_tree.xview)
        self.segment_tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.segment_tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.segment_tree.bind("<<TreeviewSelect>>", self._load_selected_segment)
        self.segment_tree.bind("<Double-1>", self._begin_segment_edit)

        actions = ttk.Frame(container)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Button(actions, text="Save All Clips", command=self._save_trim).pack(side="right")
        ttk.Button(actions, text="Cancel", command=self.destroy).pack(side="right", padx=(0, 8))

    def _load_waveform(self):
        if not self.audio_path.exists():
            messagebox.showerror(APP_TITLE, f"Audio file not found: {self.audio_path}")
            self.destroy()
            return
        cmd = [self.ffmpeg_path, "-v", "error", "-i", str(self.audio_path), "-ac", "1", "-ar", "2000", "-f", "s16le", "-"]
        result = subprocess.run(cmd, capture_output=True, timeout=120)
        if result.returncode != 0 or not result.stdout:
            messagebox.showerror(APP_TITLE, "Could not decode audio samples for trimming.")
            self.destroy()
            return
        samples = array("h")
        samples.frombytes(result.stdout)
        if not samples:
            messagebox.showerror(APP_TITLE, "Could not decode audio samples for trimming.")
            self.destroy()
            return
        self.duration = len(samples) / 2000.0
        step = max(1, len(samples) // 600)
        self.waveform = [max(abs(samples[idx]) for idx in range(start, min(len(samples), start + step))) / 32768.0 for start in range(0, len(samples), step)][:600]
        if len(self.waveform) < 600:
            self.waveform.extend([0.0] * (600 - len(self.waveform)))
        self.start_var.set(0.0)
        self.end_var.set(round(self.duration, 2))
        self.start_scale.configure(to=self.duration)
        self.end_scale.configure(to=self.duration)
        self.status_var.set(f"{self.audio_path.name} | Duration: {self.duration:.2f}s")
        self._draw_waveform()

    def _draw_waveform(self):
        self.canvas.delete("all")
        mid = self.canvas_height / 2
        x_step = self.canvas_width / len(self.waveform)
        for index, amplitude in enumerate(self.waveform):
            height = max(1, amplitude * (self.canvas_height * 0.45))
            x = index * x_step
            self.canvas.create_line(x, mid - height, x, mid + height, fill="#7fb3d5")
        start_x = int((self.start_var.get() / self.duration) * self.canvas_width) if self.duration else 0
        end_x = int((self.end_var.get() / self.duration) * self.canvas_width) if self.duration else self.canvas_width
        self.canvas.create_rectangle(0, 0, start_x, self.canvas_height, fill="#0b0f12", stipple="gray50", outline="")
        self.canvas.create_rectangle(end_x, 0, self.canvas_width, self.canvas_height, fill="#0b0f12", stipple="gray50", outline="")
        for segment in self.segments:
            seg_start = int((segment["start"] / self.duration) * self.canvas_width) if self.duration else 0
            seg_end = int((segment["end"] / self.duration) * self.canvas_width) if self.duration else self.canvas_width
            self.canvas.create_rectangle(seg_start, 0, seg_end, self.canvas_height, outline="#6dd17c", width=1)
        self.canvas.create_line(start_x, 0, start_x, self.canvas_height, fill="#00c2ff", width=3)
        self.canvas.create_line(end_x, 0, end_x, self.canvas_height, fill="#00c2ff", width=3)
        self.selection_var.set(f"Selection: {self.start_var.get():.2f}s - {self.end_var.get():.2f}s")

    def _set_from_canvas(self, which, x_value):
        if not self.duration:
            return
        time_value = round(max(0, min(self.canvas_width, x_value)) / self.canvas_width * self.duration, 2)
        if which == "start":
            self.start_var.set(min(time_value, max(0.0, self.end_var.get() - 0.1)))
        else:
            self.end_var.set(max(time_value, min(self.duration, self.start_var.get() + 0.1)))
        self._draw_waveform()

    def _on_canvas_press(self, event):
        if not self.duration:
            return
        start_x = int((self.start_var.get() / self.duration) * self.canvas_width)
        end_x = int((self.end_var.get() / self.duration) * self.canvas_width)
        self.drag_target = "start" if abs(event.x - start_x) <= abs(event.x - end_x) else "end"
        self._set_from_canvas(self.drag_target, event.x)

    def _on_canvas_drag(self, event):
        if self.drag_target:
            self._set_from_canvas(self.drag_target, event.x)

    def _on_canvas_release(self, _event):
        self.drag_target = None

    def _on_scale_change(self, changed):
        if changed == "start" and self.start_var.get() >= self.end_var.get():
            self.start_var.set(max(0.0, self.end_var.get() - 0.1))
        if changed == "end" and self.end_var.get() <= self.start_var.get():
            self.end_var.set(min(self.duration, self.start_var.get() + 0.1))
        self._draw_waveform()

    def _reset_selection(self):
        self.start_var.set(0.0)
        self.end_var.set(round(self.duration, 2))
        self._draw_waveform()

    def _segment_defaults(self, index):
        base_title = clean_filename(self.row.get("title", ""), "Clip")
        return {
            "start": round(self.start_var.get(), 2),
            "end": round(self.end_var.get(), 2),
            "title": f"{base_title} Part {index}",
            "artist": clean_filename(self.row.get("artist", ""), "Unknown Artist"),
            "album": clean_filename(self.row.get("album", ""), "Singles"),
            "track": str(index),
            "playlist": clean_filename(self.row.get("playlist", ""), "Unknown Playlist"),
            "import_type": self.row.get("import_type", "playlist"),
        }

    def _refresh_segment_tree(self):
        self.segment_tree.delete(*self.segment_tree.get_children())
        for index, segment in enumerate(self.segments):
            self.segment_tree.insert("", "end", iid=str(index), values=(
                f"{segment['start']:.2f}",
                f"{segment['end']:.2f}",
                segment["title"],
                segment["artist"],
                segment["album"],
                segment["track"],
                segment["playlist"],
                segment["import_type"],
            ))
        self._draw_waveform()

    def _add_current_segment(self):
        if self.end_var.get() - self.start_var.get() < 0.2:
            messagebox.showerror(APP_TITLE, "Clip selection must be at least 0.2 seconds long.")
            return
        self.segments.append(self._segment_defaults(len(self.segments) + 1))
        self._refresh_segment_tree()
        self.segment_tree.selection_set(str(len(self.segments) - 1))

    def _update_selected_segment(self):
        selected = self.segment_tree.selection()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Select a clip in the list first.")
            return
        if self.end_var.get() - self.start_var.get() < 0.2:
            messagebox.showerror(APP_TITLE, "Clip selection must be at least 0.2 seconds long.")
            return
        idx = int(selected[0])
        segment = dict(self.segments[idx])
        segment["start"] = round(self.start_var.get(), 2)
        segment["end"] = round(self.end_var.get(), 2)
        self.segments[idx] = segment
        self._refresh_segment_tree()
        self.segment_tree.selection_set(str(idx))

    def _remove_selected_segment(self):
        selected = self.segment_tree.selection()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Select a clip in the list first.")
            return
        del self.segments[int(selected[0])]
        self._refresh_segment_tree()

    def _load_selected_segment(self, _event=None):
        selected = self.segment_tree.selection()
        if not selected:
            return
        segment = self.segments[int(selected[0])]
        self.start_var.set(segment["start"])
        self.end_var.set(segment["end"])
        self._draw_waveform()

    def _begin_segment_edit(self, event):
        region = self.segment_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.segment_tree.identify_row(event.y)
        column_id = self.segment_tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        x, y, width, height = self.segment_tree.bbox(row_id, column_id)
        col_index = int(column_id.replace("#", "")) - 1
        keys = ["start", "end", "title", "artist", "album", "track", "playlist", "import_type"]
        key = keys[col_index]
        current_segment = dict(self.segments[int(row_id)])

        if key == "import_type":
            combo = ttk.Combobox(self.segment_tree, values=["playlist", "single"], state="readonly")
            combo.place(x=x, y=y, width=width, height=height)
            combo.set(current_segment[key] if current_segment[key] in {"playlist", "single"} else "playlist")
            combo.focus()

            def save_combo(_event=None):
                current_segment[key] = combo.get().strip() or "playlist"
                self.segments[int(row_id)] = current_segment
                self._refresh_segment_tree()
                combo.destroy()

            combo.bind("<<ComboboxSelected>>", save_combo)
            combo.bind("<FocusOut>", save_combo)
            combo.bind("<Escape>", lambda _e: combo.destroy())
            return

        entry = ttk.Entry(self.segment_tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, self.segment_tree.item(row_id, "values")[col_index])
        entry.focus()
        entry.select_range(0, "end")

        def save_edit(_event=None):
            value = entry.get().strip()
            if key in {"start", "end"}:
                try:
                    value = round(float(value), 2)
                except Exception:
                    value = current_segment[key]
            current_segment[key] = value
            if current_segment["end"] <= current_segment["start"]:
                current_segment["end"] = round(current_segment["start"] + 0.1, 2)
            self.segments[int(row_id)] = current_segment
            self._refresh_segment_tree()
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", lambda _e: entry.destroy())

    def _save_trim(self):
        if not self.segments:
            messagebox.showerror(APP_TITLE, "Add at least one clip before saving.")
            return
        normalized = []
        for index, segment in enumerate(self.segments, start=1):
            clip = dict(segment)
            clip["title"] = clean_filename(clip.get("title", ""), f"Clip {index}")
            clip["artist"] = clean_filename(clip.get("artist", ""), "Unknown Artist")
            clip["album"] = clean_filename(clip.get("album", ""), "Singles")
            clip["playlist"] = clean_filename(clip.get("playlist", ""), "Unknown Playlist")
            clip["track"] = str(clip.get("track", "")).strip() or str(index)
            clip["import_type"] = clip.get("import_type", "playlist") if clip.get("import_type", "playlist") in {"playlist", "single"} else "playlist"
            normalized.append(clip)
        self.on_save(self.row, normalized)
        self.logger(f"Saved {len(normalized)} clip(s) from {self.audio_path.name}")
        self.destroy()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1020x820")
        self.root.minsize(940, 740)
        self.root.protocol("WM_DELETE_WINDOW", self._force_close_app)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.log_queue = queue.Queue()
        self.worker = None
        self.download_process = None
        self.stop_requested = threading.Event()
        self.last_downloaded_files = []
        self.last_metadata_rows = []
        self.last_processed_rows = []
        saved = self._load_settings()
        self.url_var = tk.StringVar()
        self.import_folder_var = tk.StringVar(value=saved.get("import_folder", DEFAULT_IMPORT_FOLDER))
        self.spotify_folder_var = tk.StringVar(value=saved.get("spotify_folder", DEFAULT_SPOTIFY_FOLDER))
        self.yt_dlp_path_var = tk.StringVar(value=saved.get("yt_dlp_path", ""))
        self.ffmpeg_path_var = tk.StringVar(value=saved.get("ffmpeg_path", ""))
        self.mp3_quality_var = tk.StringVar(value=saved.get("mp3_quality", "320"))
        self.preview_volume_var = tk.IntVar(value=int(saved.get("preview_volume", 100)))
        self.theme_mode_var = tk.StringVar(value=saved.get("theme_mode", "light"))
        self.playlist_var = tk.BooleanVar(value=saved.get("playlist_downloads", True))
        self.open_folder_var = tk.BooleanVar(value=saved.get("open_folder_when_finished", True))
        self.keep_temp_var = tk.BooleanVar(value=saved.get("keep_temp_files", False))
        self.review_before_import_var = tk.BooleanVar(value=saved.get("review_before_import", True))
        self.copy_to_spotify_folder_var = tk.BooleanVar(value=saved.get("copy_to_spotify_folder", True))
        self.write_tags_var = tk.BooleanVar(value=saved.get("write_embedded_tags", True))
        self.download_history = list(saved.get("download_history", []))
        self.workspace_rows = []
        self.workspace_mode = "metadata"
        self.trim_row = None
        self.trim_segments = []
        self.trim_waveform = []
        self.trim_duration = 0.0
        self.trim_drag_target = None
        self.audio_preview_process = None
        self.audio_preview_range = None
        self.audio_preview_restart_after = None
        self.trim_start_var = tk.DoubleVar(value=0.0)
        self.trim_end_var = tk.DoubleVar(value=0.0)
        self.trim_status_var = tk.StringVar(value="No file selected.")
        self.trim_selection_var = tk.StringVar(value="Selection: 0.00s - 0.00s")
        self.trim_selected_clip_var = tk.StringVar(value="Editing: new clip")
        self.metadata_selected_index = None
        self.metadata_thumbnail_image = None
        self.status_animation_after = None
        self.current_theme_colors = {}
        self.download_state_var = tk.StringVar(value="Ready to download.")
        self.trim_preview_state_var = tk.StringVar(value="Preview stopped.")
        self.workspace_mode_var = tk.StringVar(value="Metadata")
        self.trim_title_var = tk.StringVar()
        self.trim_artist_meta_var = tk.StringVar()
        self.trim_album_meta_var = tk.StringVar()
        self.trim_track_meta_var = tk.StringVar()
        self.trim_playlist_meta_var = tk.StringVar()
        self.trim_import_type_meta_var = tk.StringVar(value="playlist")
        self.meta_filename_var = tk.StringVar()
        self.meta_title_var = tk.StringVar()
        self.meta_artist_var = tk.StringVar()
        self.meta_album_var = tk.StringVar()
        self.meta_track_var = tk.StringVar()
        self.meta_playlist_var = tk.StringVar()
        self.meta_import_type_var = tk.StringVar(value="playlist")
        self.meta_source_path_var = tk.StringVar()
        self.meta_artwork_path_var = tk.StringVar()
        self.files_selected_index = None
        self.file_title_var = tk.StringVar()
        self.file_artist_var = tk.StringVar()
        self.file_album_var = tk.StringVar()
        self.file_track_var = tk.StringVar()
        self.file_playlist_var = tk.StringVar()
        self.file_import_type_var = tk.StringVar(value="playlist")
        self.file_source_path_var = tk.StringVar()
        self.file_spotify_path_var = tk.StringVar()
        self.file_artwork_path_var = tk.StringVar()
        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self):
        self.main = ttk.Frame(self.root, padding=12)
        self.main.grid(row=0, column=0, sticky="nsew")
        self.main.columnconfigure(0, weight=1)
        self.main.rowconfigure(2, weight=1)

        header = ttk.Frame(self.main, style="Toolbar.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text=APP_TITLE, font=("Segoe UI", 18, "bold"), style="Hero.TLabel").grid(row=0, column=0, sticky="w")
        theme_toggle = ttk.Frame(header, style="Toolbar.TFrame")
        theme_toggle.grid(row=0, column=1, sticky="e")
        self.light_theme_button = ttk.Button(theme_toggle, text="☀", command=lambda: self._set_theme("light"), style="Theme.TButton", width=3)
        self.light_theme_button.pack(side="left")
        self.dark_theme_button = ttk.Button(theme_toggle, text="☾", command=lambda: self._set_theme("dark"), style="Theme.TButton", width=3)
        self.dark_theme_button.pack(side="left", padx=(6, 0))

        ttk.Label(
            self.main,
            text="Paste a YouTube link, pull down clean MP3s, review the track data, and send the final files into your Spotify-ready folder.",
            wraplength=960,
            style="Subtle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(6, 10))

        notebook_shell = ttk.Frame(self.main, style="Shell.TFrame", padding=8)
        notebook_shell.grid(row=2, column=0, sticky="nsew")
        notebook_shell.columnconfigure(0, weight=1)
        notebook_shell.rowconfigure(0, weight=1)

        self.main_notebook = ttk.Notebook(notebook_shell)
        self.main_notebook.grid(row=0, column=0, sticky="nsew")
        self.main_notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        home = ttk.Frame(self.main_notebook, padding=12)
        self.workspace = ttk.Frame(self.main_notebook, padding=12)
        settings = ttk.Frame(self.main_notebook, padding=12)
        self.main_notebook.add(home, text="Home")
        self.main_notebook.add(self.workspace, text="Workspace")
        self.main_notebook.add(settings, text="Settings")

        home.columnconfigure(0, weight=1)
        home.rowconfigure(2, weight=1)
        self.workspace.columnconfigure(0, weight=1)
        self.workspace.rowconfigure(2, weight=1)
        settings.columnconfigure(0, weight=1)

        form_card = ttk.LabelFrame(home, text="Download Setup", padding=14)
        form_card.grid(row=0, column=0, sticky="ew")
        form_card.columnconfigure(0, weight=1)
        form = ttk.Frame(form_card, style="Card.TFrame")
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="YouTube URL").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        ttk.Entry(form, textvariable=self.url_var).grid(row=0, column=1, sticky="ew", pady=(0, 8))
        ttk.Label(form, text="Download folder").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        import_frame = ttk.Frame(form)
        import_frame.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        import_frame.columnconfigure(0, weight=1)
        ttk.Entry(import_frame, textvariable=self.import_folder_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(import_frame, text="Browse", command=self._choose_import_folder).grid(row=0, column=1, padx=(8, 0))
        ttk.Label(form, text="Spotify-ready folder").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        spotify_frame = ttk.Frame(form)
        spotify_frame.grid(row=2, column=1, sticky="ew", pady=(0, 8))
        spotify_frame.columnconfigure(0, weight=1)
        ttk.Entry(spotify_frame, textvariable=self.spotify_folder_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(spotify_frame, text="Browse", command=self._choose_spotify_folder).grid(row=0, column=1, padx=(8, 0))
        ttk.Label(form, text="MP3 quality").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        ttk.Combobox(form, textvariable=self.mp3_quality_var, state="readonly", values=["320", "256", "192", "128"]).grid(row=3, column=1, sticky="w", pady=(0, 8))

        actions = ttk.LabelFrame(home, text="Download Controls", padding=14)
        actions.grid(row=1, column=0, sticky="ew", pady=(6, 8))
        action_row = ttk.Frame(actions, style="Toolbar.TFrame")
        action_row.pack(fill="x")
        self.download_button = ttk.Button(action_row, text="Start Download", command=self._start_download, style="Primary.TButton")
        self.download_button.pack(side="left")
        self.stop_button = ttk.Button(action_row, text="Stop", command=self._stop_download, state="disabled", style="Danger.TButton")
        self.stop_button.pack(side="left", padx=(8, 0))
        self.download_progress = ttk.Progressbar(action_row, mode="indeterminate", length=140, style="Accent.Horizontal.TProgressbar")
        self.download_progress.pack(side="left", padx=(12, 0))
        ttk.Label(action_row, textvariable=self.download_state_var, style="Chip.TLabel").pack(side="left", padx=(12, 0))

        log_frame = ttk.LabelFrame(home, text="Activity", padding=10)
        log_frame.grid(row=2, column=0, sticky="nsew")
        self.log_text = tk.Text(log_frame, wrap="word", height=16)
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")

        ttk.Label(home, text="Point Spotify Desktop at the Spotify-ready folder through Settings -> Local Files -> Add a source.", wraplength=960, style="Subtle.TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))

        workspace_header = ttk.Frame(self.workspace)
        workspace_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        workspace_header.columnconfigure(0, weight=1)
        ttk.Label(workspace_header, text="Workspace", font=("Segoe UI", 14, "bold"), style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(workspace_header, textvariable=self.workspace_mode_var, style="Badge.TLabel").grid(row=0, column=1, sticky="e")
        self.workspace_status_label = ttk.Label(workspace_header, text="Open metadata, recent files, or trimming tools here.", style="Subtle.TLabel")
        self.workspace_status_label.grid(row=1, column=0, sticky="w", pady=(4, 0))

        workspace_actions = ttk.LabelFrame(self.workspace, text="Workspace Tools", padding=12)
        workspace_actions.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        workspace_toolbar = ttk.Frame(workspace_actions, style="Toolbar.TFrame")
        workspace_toolbar.pack(fill="x")
        ttk.Button(workspace_toolbar, text="Metadata", command=self._open_metadata_editor, style="Workspace.TButton").pack(side="left")
        ttk.Button(workspace_toolbar, text="Recent Files", command=self._open_output_file_manager, style="Workspace.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(workspace_toolbar, text="Trim", command=self._open_trim_from_selection, style="Workspace.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(workspace_toolbar, text="Reload", command=self._reload_recent_files_action, style="Secondary.TButton").pack(side="left", padx=(12, 0))

        self.workspace_stack = ttk.Frame(self.workspace, style="Shell.TFrame", padding=8)
        self.workspace_stack.grid(row=2, column=0, sticky="nsew")
        self.workspace_stack.columnconfigure(0, weight=1)
        self.workspace_stack.rowconfigure(0, weight=1)

        self.metadata_panel = ttk.Frame(self.workspace_stack, padding=10, style="Card.TFrame")
        self.files_panel = ttk.Frame(self.workspace_stack, padding=10, style="Card.TFrame")
        self.trim_panel = ttk.Frame(self.workspace_stack, padding=10, style="Card.TFrame")
        for panel in (self.metadata_panel, self.files_panel, self.trim_panel):
            panel.grid(row=0, column=0, sticky="nsew")

        self._build_metadata_panel()
        self._build_files_panel()
        self._build_trim_panel()
        self._show_workspace_panel("metadata")

        install_frame = ttk.LabelFrame(settings, text="Install Locations", padding=12)
        install_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        install_frame.columnconfigure(1, weight=1)
        ttk.Label(install_frame, text="Custom yt-dlp path").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        yt_dlp_frame = ttk.Frame(install_frame)
        yt_dlp_frame.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        yt_dlp_frame.columnconfigure(0, weight=1)
        ttk.Entry(yt_dlp_frame, textvariable=self.yt_dlp_path_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(yt_dlp_frame, text="Browse", command=self._choose_yt_dlp_path).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(yt_dlp_frame, text="Auto-detect", command=lambda: self._autodetect_tool_path("yt-dlp")).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(yt_dlp_frame, text="Clear", command=self._clear_yt_dlp_path).grid(row=0, column=3, padx=(8, 0))
        ttk.Label(install_frame, text="Custom ffmpeg path").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        ffmpeg_frame = ttk.Frame(install_frame)
        ffmpeg_frame.grid(row=1, column=1, sticky="ew", pady=(0, 8))
        ffmpeg_frame.columnconfigure(0, weight=1)
        ttk.Entry(ffmpeg_frame, textvariable=self.ffmpeg_path_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(ffmpeg_frame, text="Browse", command=self._choose_ffmpeg_path).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(ffmpeg_frame, text="Auto-detect", command=lambda: self._autodetect_tool_path("ffmpeg")).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(ffmpeg_frame, text="Clear", command=self._clear_ffmpeg_path).grid(row=0, column=3, padx=(8, 0))

        options = ttk.LabelFrame(settings, text="Pipeline Options", padding=12)
        options.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        ttk.Checkbutton(options, text="Allow playlist downloads", variable=self.playlist_var).pack(anchor="w")
        ttk.Checkbutton(options, text="Review metadata before sending to Spotify folder", variable=self.review_before_import_var).pack(anchor="w")
        ttk.Checkbutton(options, text="Copy cleaned songs into Spotify-ready folder automatically", variable=self.copy_to_spotify_folder_var).pack(anchor="w")
        ttk.Checkbutton(options, text="Write corrected embedded MP3 tags with mutagen", variable=self.write_tags_var).pack(anchor="w")
        ttk.Checkbutton(options, text="Open output folder when finished", variable=self.open_folder_var).pack(anchor="w")
        ttk.Checkbutton(options, text="Keep temp files created during processing", variable=self.keep_temp_var).pack(anchor="w")

        deps = ttk.LabelFrame(settings, text="Required Tools", padding=12)
        deps.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        deps.columnconfigure(0, weight=1)

        status_frame = ttk.LabelFrame(deps, text="Connection", padding=10)
        status_frame.pack(fill="x", anchor="w")
        self.tool_status_rows = {}
        self.dependency_link_labels = []
        for service_name, key in (("yt-dlp", "yt-dlp"), ("ffmpeg", "ffmpeg"), ("mutagen", "mutagen"), ("pyinstaller", "pyinstaller")):
            row = ttk.Frame(status_frame)
            row.pack(fill="x", anchor="w", pady=2)
            ttk.Label(row, text=f"{service_name}:").pack(side="left")
            status_label = tk.Label(row, text="Checking...", anchor="w", font=("Segoe UI", 10, "bold"))
            status_label.pack(side="left", padx=(8, 0))
            self.tool_status_rows[key] = status_label

        instructions_frame = ttk.LabelFrame(deps, text="Instructions", padding=10)
        instructions_frame.pack(fill="x", anchor="w", pady=(10, 0))
        ttk.Label(
            instructions_frame,
            text="Install the missing tools below, or point the app at the executable in Install Locations. mutagen and pyinstaller are optional Python packages.",
            wraplength=920,
            justify="left",
        ).pack(anchor="w")
        self._add_dependency_link(instructions_frame, "Download yt-dlp", "https://github.com/yt-dlp/yt-dlp/releases/latest")
        self._add_dependency_link(instructions_frame, "Download FFmpeg", "https://www.gyan.dev/ffmpeg/builds/")
        self._add_dependency_link(instructions_frame, "mutagen on PyPI", "https://pypi.org/project/mutagen/")
        self._add_dependency_link(instructions_frame, "PyInstaller docs", "https://pyinstaller.org/en/stable/")
        ttk.Label(
            instructions_frame,
            text="Quick install commands: winget install yt-dlp.yt-dlp | winget install Gyan.FFmpeg | python -m pip install mutagen pyinstaller",
            wraplength=920,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

        ttk.Button(deps, text="Refresh dependency check", command=self._refresh_dependency_check).pack(anchor="w", pady=(10, 0))
        self._refresh_dependency_check()

        utility_actions = ttk.LabelFrame(settings, text="Utilities", padding=12)
        utility_actions.grid(row=3, column=0, sticky="ew")
        ttk.Button(utility_actions, text="Reload Recent Files", command=self._reload_recent_files_action).pack(side="left")
        ttk.Button(utility_actions, text="Open Metadata Editor", command=self._open_metadata_editor).pack(side="left", padx=(8, 0))
        ttk.Button(utility_actions, text="Manage Created Files", command=self._open_output_file_manager).pack(side="left", padx=(8, 0))
        ttk.Button(utility_actions, text="Send Last Download to Spotify Folder", command=self._send_last_download_to_spotify_folder).pack(side="left", padx=(8, 0))
        ttk.Button(utility_actions, text="Export Build Script", command=self._export_build_script).pack(side="left", padx=(8, 0))
        ttk.Button(utility_actions, text="Clear Log", command=self._clear_log).pack(side="right")
        self._apply_theme(self.theme_mode_var.get())
        self._set_download_state("idle")

    def _set_theme(self, mode: str):
        self.theme_mode_var.set(mode)
        self._apply_theme(mode)
        self._save_settings()

    def _apply_theme(self, mode: str):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        if mode == "dark":
            colors = {"bg": "#11151b", "panel": "#1b2129", "panel_alt": "#252d37", "panel_soft": "#303948", "fg": "#f3f5f8", "muted": "#9ba7b4", "accent": "#6ec1ff", "accent_soft": "#1f405a", "field": "#141920", "danger": "#d56a6a", "border": "#394453"}
        else:
            colors = {"bg": "#edf1f5", "panel": "#ffffff", "panel_alt": "#e6edf5", "panel_soft": "#f6f8fb", "fg": "#19202a", "muted": "#5f6c7b", "accent": "#0f6cbd", "accent_soft": "#d8e8f8", "field": "#ffffff", "danger": "#c14646", "border": "#ccd6e0"}
        self.current_theme_colors = colors
        self.root.configure(bg=colors["bg"])
        style.configure(".", background=colors["bg"], foreground=colors["fg"])
        style.configure("TFrame", background=colors["bg"])
        style.configure("Shell.TFrame", background=colors["panel_soft"])
        style.configure("Card.TFrame", background=colors["panel"])
        style.configure("Toolbar.TFrame", background=colors["panel"])
        style.configure("TLabel", background=colors["bg"], foreground=colors["fg"])
        style.configure("Hero.TLabel", background=colors["bg"], foreground=colors["fg"])
        style.configure("Section.TLabel", background=colors["bg"], foreground=colors["fg"])
        style.configure("Subtle.TLabel", background=colors["bg"], foreground=colors["muted"])
        style.configure("CardTitle.TLabel", background=colors["panel"], foreground=colors["fg"])
        style.configure("CardSubtle.TLabel", background=colors["panel"], foreground=colors["muted"])
        style.configure("CardBadge.TLabel", background=colors["accent_soft"], foreground=colors["accent"], padding=(10, 5), font=("Segoe UI", 9, "bold"))
        style.configure("CardChip.TLabel", background=colors["panel_alt"], foreground=colors["fg"], padding=(10, 6))
        style.configure("Badge.TLabel", background=colors["accent_soft"], foreground=colors["accent"], padding=(10, 5), font=("Segoe UI", 9, "bold"))
        style.configure("Chip.TLabel", background=colors["panel_alt"], foreground=colors["fg"], padding=(10, 6))
        style.configure("TLabelframe", background=colors["panel"], foreground=colors["fg"], bordercolor=colors["border"], relief="solid")
        style.configure("TLabelframe.Label", background=colors["panel"], foreground=colors["fg"])
        style.configure("TButton", background=colors["panel"], foreground=colors["fg"], padding=(12, 8), borderwidth=0)
        style.map("TButton", background=[("active", colors["panel_alt"])], foreground=[("active", colors["fg"])])
        style.configure("Primary.TButton", background=colors["accent"], foreground=colors["panel"], padding=(14, 8), borderwidth=0)
        style.map("Primary.TButton", background=[("active", colors["fg"]), ("disabled", colors["panel_alt"])], foreground=[("active", colors["accent"]), ("disabled", colors["muted"])])
        style.configure("Secondary.TButton", background=colors["accent_soft"], foreground=colors["accent"], padding=(12, 8), borderwidth=0)
        style.map("Secondary.TButton", background=[("active", colors["accent"]), ("disabled", colors["panel_alt"])], foreground=[("active", colors["panel"]), ("disabled", colors["muted"])])
        style.configure("Nav.TButton", background=colors["panel"], foreground=colors["fg"], padding=(12, 7), borderwidth=0)
        style.map("Nav.TButton", background=[("active", colors["panel_alt"])], foreground=[("active", colors["fg"])])
        style.configure("Workspace.TButton", background=colors["panel_alt"], foreground=colors["fg"], padding=(14, 8), borderwidth=0)
        style.map("Workspace.TButton", background=[("active", colors["accent_soft"])], foreground=[("active", colors["accent"])])
        style.configure("Transport.TButton", background=colors["panel_alt"], foreground=colors["fg"], padding=(14, 8), borderwidth=0)
        style.map("Transport.TButton", background=[("active", colors["accent_soft"])], foreground=[("active", colors["accent"])])
        style.configure("Danger.TButton", background=colors["danger"], foreground=colors["panel"], padding=(12, 8), borderwidth=0)
        style.map("Danger.TButton", background=[("active", colors["fg"]), ("disabled", colors["panel_alt"])], foreground=[("active", colors["danger"]), ("disabled", colors["muted"])])
        style.configure("Theme.TButton", background=colors["panel_alt"], foreground=colors["muted"], padding=(10, 6), borderwidth=0, font=("Segoe UI", 9, "bold"))
        style.map("Theme.TButton", background=[("active", colors["panel"])], foreground=[("active", colors["fg"])])
        style.configure("TCheckbutton", background=colors["bg"], foreground=colors["fg"])
        style.configure("TNotebook", background=colors["panel_soft"], borderwidth=0, tabmargins=(6, 6, 6, 0))
        style.configure("TNotebook.Tab", background=colors["panel_alt"], foreground=colors["muted"], padding=(20, 11), borderwidth=0, font=("Segoe UI", 10, "bold"))
        style.map(
            "TNotebook.Tab",
            background=[("selected", colors["accent"]), ("active", colors["panel"])],
            foreground=[("selected", colors["panel"]), ("active", colors["fg"])],
            padding=[("selected", (24, 12)), ("!selected", (20, 11))],
            expand=[("selected", [1, 1, 1, 0])],
        )
        style.configure("TEntry", fieldbackground=colors["field"], foreground=colors["fg"])
        style.configure("TCombobox", fieldbackground=colors["field"], foreground=colors["fg"])
        style.configure("Treeview", background=colors["panel"], fieldbackground=colors["panel"], foreground=colors["fg"], rowheight=32, borderwidth=0)
        style.configure("Treeview.Heading", background=colors["panel_alt"], foreground=colors["fg"], padding=(8, 8), borderwidth=0)
        style.map("Treeview", background=[("selected", colors["accent_soft"])], foreground=[("selected", colors["fg"])])
        style.map("Treeview.Heading", background=[("active", colors["panel"])])
        style.configure("Accent.Horizontal.TProgressbar", background=colors["accent"], troughcolor=colors["panel_alt"], bordercolor=colors["panel_alt"], lightcolor=colors["accent"], darkcolor=colors["accent"])
        if hasattr(self, "log_text") and self.log_text:
            self.log_text.configure(bg=colors["field"], fg=colors["fg"], insertbackground=colors["fg"], selectbackground=colors["accent_soft"], relief="flat", font=("Segoe UI", 10), padx=10, pady=10, spacing1=4, spacing3=6)
        if hasattr(self, "trim_canvas") and self.trim_canvas:
            self.trim_canvas.configure(
                background=colors["field"],
                highlightbackground=colors["border"],
                highlightcolor=colors["border"],
            )
        if hasattr(self, "trim_range_canvas") and self.trim_range_canvas:
            self.trim_range_canvas.configure(
                background=colors["panel_soft"],
                highlightbackground=colors["border"],
                highlightcolor=colors["border"],
            )
        if hasattr(self, "tool_status_rows"):
            for label in self.tool_status_rows.values():
                label.configure(bg=colors["panel"])
        if hasattr(self, "dependency_link_labels"):
            for label in self.dependency_link_labels:
                label.configure(bg=colors["panel"], fg=colors["accent"])
        self._update_theme_toggle_buttons()

    def _update_theme_toggle_buttons(self):
        if not hasattr(self, "light_theme_button"):
            return
        active_mode = self.theme_mode_var.get().strip() or "light"
        colors = self.current_theme_colors or {}
        active_bg = colors.get("accent", "#0f6cbd")
        active_fg = colors.get("panel", "#ffffff")
        inactive_bg = colors.get("panel_alt", "#e6edf5")
        inactive_fg = colors.get("muted", "#5f6c7b")
        self.light_theme_button.configure(style="Theme.TButton")
        self.dark_theme_button.configure(style="Theme.TButton")
        if active_mode == "light":
            self.light_theme_button.configure(style="Primary.TButton")
            self.dark_theme_button.configure(style="Theme.TButton")
        else:
            self.dark_theme_button.configure(style="Primary.TButton")
            self.light_theme_button.configure(style="Theme.TButton")

    def _build_metadata_panel(self):
        self.metadata_panel.columnconfigure(0, weight=2)
        self.metadata_panel.columnconfigure(1, weight=1)
        self.metadata_panel.rowconfigure(1, weight=1)
        ttk.Label(self.metadata_panel, text="Metadata Editor", font=("Segoe UI", 12, "bold"), style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(self.metadata_panel, text="Select a track, adjust the fields on the right, then apply or save the full set.", style="CardSubtle.TLabel").grid(row=0, column=1, sticky="e", pady=(0, 4))
        tree_frame = ttk.Frame(self.metadata_panel)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        columns = ("filename", "title", "artist", "album", "track", "playlist", "import_type")
        self.metadata_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        widths = {"filename": 220, "title": 180, "artist": 140, "album": 140, "track": 56, "playlist": 180, "import_type": 88}
        for col, label in [("filename", "File"), ("title", "Title"), ("artist", "Artist"), ("album", "Album"), ("track", "Track"), ("playlist", "Playlist"), ("import_type", "Type")]:
            self.metadata_tree.heading(col, text=label)
            self.metadata_tree.column(col, width=widths[col], minwidth=widths[col], stretch=(col in {"filename", "title", "playlist"}), anchor="w")
        ybar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.metadata_tree.yview)
        xbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.metadata_tree.xview)
        self.metadata_tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.metadata_tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.metadata_tree.bind("<<TreeviewSelect>>", self._load_metadata_selection)
        self.metadata_tree.bind("<Double-1>", self._begin_metadata_edit)
        actions = ttk.Frame(self.metadata_panel, style="Toolbar.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(actions, text="Auto-clean", command=self._metadata_auto_clean, style="Nav.TButton").pack(side="left")
        ttk.Button(actions, text="Fill Track Numbers", command=self._metadata_fill_tracks, style="Nav.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Apply Selected Changes", command=self._apply_metadata_form, style="Secondary.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Save Metadata", command=self._save_metadata_from_workspace, style="Primary.TButton").pack(side="right")

        inspector = ttk.LabelFrame(self.metadata_panel, text="Selected Track", padding=12)
        inspector.grid(row=1, column=1, sticky="nsew", padx=(12, 0))
        inspector.columnconfigure(1, weight=1)
        self.metadata_thumbnail_label = ttk.Label(inspector, text="No artwork", anchor="center")
        self.metadata_thumbnail_label.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        ttk.Label(inspector, text="Filename").grid(row=1, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.meta_filename_var).grid(row=1, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Title").grid(row=2, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.meta_title_var).grid(row=2, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Artist").grid(row=3, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.meta_artist_var).grid(row=3, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Album").grid(row=4, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.meta_album_var).grid(row=4, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Track #").grid(row=5, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.meta_track_var).grid(row=5, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Playlist").grid(row=6, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.meta_playlist_var).grid(row=6, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Type").grid(row=7, column=0, sticky="w", pady=(0, 6))
        ttk.Combobox(inspector, textvariable=self.meta_import_type_var, state="readonly", values=["playlist", "single"]).grid(row=7, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Artwork").grid(row=8, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.meta_artwork_path_var, state="readonly").grid(row=8, column=1, sticky="ew", pady=(0, 6))
        meta_artwork_actions = ttk.Frame(inspector, style="Card.TFrame")
        meta_artwork_actions.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Button(meta_artwork_actions, text="Use YouTube Thumbnail", command=self._use_metadata_thumbnail_artwork, style="Nav.TButton").pack(side="left")
        ttk.Button(meta_artwork_actions, text="Choose Artwork", command=self._choose_metadata_artwork, style="Nav.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(meta_artwork_actions, text="Clear Artwork", command=self._clear_metadata_artwork, style="Nav.TButton").pack(side="left", padx=(8, 0))
        ttk.Label(inspector, text="Source file").grid(row=10, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.meta_source_path_var, state="readonly").grid(row=10, column=1, sticky="ew", pady=(0, 6))

    def _build_files_panel(self):
        self.files_panel.columnconfigure(0, weight=1)
        self.files_panel.columnconfigure(1, weight=1)
        self.files_panel.rowconfigure(1, weight=1)
        ttk.Label(self.files_panel, text="Recent Files", font=("Segoe UI", 12, "bold"), style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(self.files_panel, text="Use this list to review what the app created and jump straight into trimming.", style="CardSubtle.TLabel").grid(row=0, column=1, sticky="e", pady=(0, 4))
        tree_frame = ttk.Frame(self.files_panel)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        columns = ("title", "artist", "album", "playlist", "status", "spotify_path", "source_path")
        self.files_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        widths = {"title": 180, "artist": 130, "album": 130, "playlist": 150, "status": 84, "spotify_path": 260, "source_path": 260}
        for col, label in [("title", "Title"), ("artist", "Artist"), ("album", "Album"), ("playlist", "Playlist"), ("status", "Status"), ("spotify_path", "Spotify File"), ("source_path", "Source File")]:
            self.files_tree.heading(col, text=label)
            self.files_tree.column(col, width=widths[col], minwidth=widths[col], stretch=(col in {"title", "spotify_path", "source_path"}), anchor="w")
        ybar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.files_tree.yview)
        xbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.files_tree.xview)
        self.files_tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.files_tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        self.files_tree.bind("<<TreeviewSelect>>", self._load_files_selection)
        actions = ttk.Frame(self.files_panel, style="Toolbar.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(actions, text="Trim Selected", command=self._open_trim_from_selection, style="Primary.TButton").pack(side="left")
        ttk.Button(actions, text="Delete Selected", command=self._delete_selected_from_workspace, style="Danger.TButton").pack(side="left", padx=(8, 0))

        inspector = ttk.LabelFrame(self.files_panel, text="Selected File Metadata", padding=12)
        inspector.grid(row=1, column=1, sticky="nsew", padx=(12, 0))
        inspector.columnconfigure(1, weight=1)
        ttk.Label(inspector, text="Title").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.file_title_var).grid(row=0, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Artist").grid(row=1, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.file_artist_var).grid(row=1, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Album").grid(row=2, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.file_album_var).grid(row=2, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Track #").grid(row=3, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.file_track_var).grid(row=3, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Playlist").grid(row=4, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.file_playlist_var).grid(row=4, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Type").grid(row=5, column=0, sticky="w", pady=(0, 6))
        ttk.Combobox(inspector, textvariable=self.file_import_type_var, state="readonly", values=["playlist", "single"]).grid(row=5, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Source file").grid(row=6, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.file_source_path_var, state="readonly").grid(row=6, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Spotify file").grid(row=7, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.file_spotify_path_var, state="readonly").grid(row=7, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(inspector, text="Artwork").grid(row=8, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(inspector, textvariable=self.file_artwork_path_var, state="readonly").grid(row=8, column=1, sticky="ew", pady=(0, 6))
        file_artwork_actions = ttk.Frame(inspector, style="Card.TFrame")
        file_artwork_actions.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Button(file_artwork_actions, text="Use YouTube Thumbnail", command=self._use_files_thumbnail_artwork, style="Nav.TButton").pack(side="left")
        ttk.Button(file_artwork_actions, text="Choose Artwork", command=self._choose_files_artwork, style="Nav.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(file_artwork_actions, text="Clear Artwork", command=self._clear_files_artwork, style="Nav.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(inspector, text="Save File Metadata", command=self._apply_files_form, style="Primary.TButton").grid(row=10, column=0, columnspan=2, sticky="ew", pady=(10, 0))

    def _build_trim_panel(self):
        self.trim_panel.columnconfigure(0, weight=5)
        self.trim_panel.columnconfigure(1, weight=3)
        self.trim_panel.rowconfigure(5, weight=1)
        ttk.Label(self.trim_panel, text="Trim Workspace", font=("Segoe UI", 12, "bold"), style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(self.trim_panel, textvariable=self.trim_preview_state_var, style="CardBadge.TLabel").grid(row=0, column=1, sticky="e", pady=(0, 4))
        ttk.Label(self.trim_panel, textvariable=self.trim_status_var, style="CardSubtle.TLabel").grid(row=1, column=0, sticky="w")
        self.trim_canvas = tk.Canvas(self.trim_panel, width=840, height=118, background="#101418", highlightthickness=1, highlightbackground="#405060")
        self.trim_canvas.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 6))
        self.trim_canvas.bind("<Button-1>", self._on_trim_canvas_press)
        self.trim_canvas.bind("<B1-Motion>", self._on_trim_canvas_drag)
        self.trim_canvas.bind("<ButtonRelease-1>", self._on_trim_canvas_release)
        self.trim_canvas.bind("<Configure>", lambda _event: self._draw_trim_waveform())
        range_frame = ttk.Frame(self.trim_panel, style="Card.TFrame")
        range_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        range_frame.columnconfigure(0, weight=1)
        ttk.Label(range_frame, textvariable=self.trim_selection_var, style="CardChip.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.trim_range_canvas = tk.Canvas(range_frame, width=840, height=38, background="#dfe6ee", highlightthickness=1, highlightbackground="#b9c4d0")
        self.trim_range_canvas.grid(row=1, column=0, sticky="ew")
        self.trim_range_canvas.bind("<Button-1>", self._on_trim_canvas_press)
        self.trim_range_canvas.bind("<B1-Motion>", self._on_trim_canvas_drag)
        self.trim_range_canvas.bind("<ButtonRelease-1>", self._on_trim_canvas_release)
        self.trim_range_canvas.bind("<Configure>", lambda _event: self._draw_trim_waveform())
        transport = ttk.Frame(self.trim_panel, style="Toolbar.TFrame")
        transport.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(2, 4))
        ttk.Button(transport, text="Play Selection", command=self._play_trim_selection, style="Transport.TButton").pack(side="left")
        ttk.Button(transport, text="Play Clip", command=self._play_selected_trim_clip, style="Transport.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(transport, text="Play Full", command=self._play_trim_full_track, style="Transport.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(transport, text="Stop", command=self._stop_audio_preview, style="Danger.TButton").pack(side="left", padx=(8, 14))
        ttk.Label(transport, text="Volume").pack(side="left", padx=(0, 6))
        preview_volume_scale = tk.Scale(
            transport,
            from_=0,
            to=200,
            variable=self.preview_volume_var,
            orient="horizontal",
            resolution=1,
            showvalue=False,
            length=120,
            command=self._on_preview_volume_change,
        )
        preview_volume_scale.pack(side="left")
        ttk.Label(transport, textvariable=self._preview_volume_text_var()).pack(side="left", padx=(8, 14))
        ttk.Separator(transport, orient="vertical").pack(side="left", fill="y", padx=(0, 14))
        ttk.Button(transport, text="Add", command=self._add_trim_segment, style="Primary.TButton").pack(side="left")
        ttk.Button(transport, text="Update", command=self._update_trim_segment, style="Secondary.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(transport, text="Remove", command=self._remove_trim_segment, style="Nav.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(transport, text="Save All Clips", command=self._save_trim_workspace, style="Primary.TButton").pack(side="right")
        columns = ("start", "end", "title", "track")
        tree_frame = ttk.Frame(self.trim_panel)
        tree_frame.grid(row=5, column=0, sticky="nsew", padx=(0, 12))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.trim_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        widths = {"start": 54, "end": 54, "title": 150, "track": 34}
        for col, label in [("start", "Start"), ("end", "End"), ("title", "Title"), ("track", "#")]:
            self.trim_tree.heading(col, text=label)
            self.trim_tree.column(col, width=widths[col], minwidth=widths[col], stretch=(col == "title"), anchor="w")
        ybar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.trim_tree.yview)
        self.trim_tree.configure(yscrollcommand=ybar.set)
        self.trim_tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        self.trim_tree.bind("<<TreeviewSelect>>", self._load_selected_trim_segment)
        self.trim_tree.bind("<Double-1>", self._begin_trim_segment_edit)

        trim_details = ttk.LabelFrame(self.trim_panel, text="Selected Clip Details", padding=12)
        trim_details.grid(row=5, column=1, sticky="nsew")
        trim_details.columnconfigure(1, weight=1)
        ttk.Label(trim_details, textvariable=self.trim_selected_clip_var, style="CardSubtle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(trim_details, text="Title").grid(row=1, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(trim_details, textvariable=self.trim_title_var).grid(row=1, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(trim_details, text="Artist").grid(row=2, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(trim_details, textvariable=self.trim_artist_meta_var).grid(row=2, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(trim_details, text="Album").grid(row=3, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(trim_details, textvariable=self.trim_album_meta_var).grid(row=3, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(trim_details, text="Track #").grid(row=4, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(trim_details, textvariable=self.trim_track_meta_var).grid(row=4, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(trim_details, text="Playlist").grid(row=5, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(trim_details, textvariable=self.trim_playlist_meta_var).grid(row=5, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(trim_details, text="Type").grid(row=6, column=0, sticky="w", pady=(0, 6))
        ttk.Combobox(trim_details, textvariable=self.trim_import_type_meta_var, state="readonly", values=["playlist", "single"]).grid(row=6, column=1, sticky="ew", pady=(0, 6))
        self.trim_save_button = ttk.Button(trim_details, text="Save Clip Settings", command=self._save_selected_trim_segment, style="Primary.TButton", state="disabled")
        self.trim_save_button.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(8, 8))
        ttk.Label(trim_details, text="Use this panel to give each saved clip a clear name and complete metadata before exporting.", style="CardSubtle.TLabel", wraplength=320, justify="left").grid(row=8, column=0, columnspan=2, sticky="nw")

    def _show_workspace_panel(self, mode: str):
        self.workspace_mode = mode
        panel_map = {"metadata": self.metadata_panel, "files": self.files_panel, "trim": self.trim_panel}
        panel_map[mode].tkraise()
        self.main_notebook.select(self.workspace)
        self._animate_workspace_status()

    def _on_tab_changed(self, _event=None):
        current = self.main_notebook.tab(self.main_notebook.select(), "text")
        if current == "Home":
            self.download_state_var.set(self.download_state_var.get() or "Ready to download.")
        elif current == "Workspace":
            self._animate_workspace_status()

    def _set_download_state(self, state: str):
        states = {
            "idle": ("Ready to download.", "disabled"),
            "running": ("Downloading and processing...", "normal"),
            "stopping": ("Stopping current download...", "normal"),
            "done": ("Download finished.", "disabled"),
            "error": ("Download stopped with an error.", "disabled"),
        }
        message, stop_state = states.get(state, states["idle"])
        self.download_state_var.set(message)
        if hasattr(self, "stop_button"):
            self.stop_button.configure(state=stop_state if state in {"running", "stopping"} else "disabled")
        if hasattr(self, "download_button"):
            self.download_button.configure(state="disabled" if state in {"running", "stopping"} else "normal")
        if hasattr(self, "download_progress"):
            if state in {"running", "stopping"}:
                self.download_progress.start(12)
            else:
                self.download_progress.stop()

    def _set_workspace_status(self, mode_label: str, message: str):
        self.workspace_mode_var.set(mode_label)
        self.workspace_status_label.configure(text=message)
        self._animate_workspace_status()

    def _animate_workspace_status(self, steps: int = 6, tick: int = 120):
        if not hasattr(self, "workspace_status_label"):
            return
        base_text = self.workspace_status_label.cget("text").rstrip(".")
        if self.status_animation_after:
            self.root.after_cancel(self.status_animation_after)
            self.status_animation_after = None

        def pulse(index: int = 0):
            dots = "." * (index % 4)
            self.workspace_status_label.configure(text=f"{base_text}{dots}")
            if index < steps:
                self.status_animation_after = self.root.after(tick, lambda: pulse(index + 1))
            else:
                self.workspace_status_label.configure(text=base_text)
                self.status_animation_after = None

        pulse(0)

    def _normalize_source_url(self, url: str) -> str:
        normalized = (url or "").strip()
        normalized = re.sub(r"#.*$", "", normalized)
        normalized = normalized.rstrip("/")
        return normalized

    def _confirm_duplicate_download(self, url: str) -> bool:
        dialog = tk.Toplevel(self.root)
        dialog.title("Download Again?")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        colors = self.current_theme_colors or {}
        dialog.configure(bg=colors.get("bg", "#edf1f5"))

        shell = ttk.Frame(dialog, padding=16, style="Card.TFrame")
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=1)

        ttk.Label(shell, text="This source was downloaded before.", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(shell, text="Are you sure you want to download this again?", style="CardSubtle.TLabel", wraplength=360, justify="left").grid(row=1, column=0, sticky="w", pady=(8, 4))
        ttk.Label(shell, text=url, style="CardSubtle.TLabel", wraplength=360, justify="left").grid(row=2, column=0, sticky="w")

        result = {"value": False}

        def close_with(value: bool):
            result["value"] = value
            dialog.destroy()

        actions = ttk.Frame(shell, style="Toolbar.TFrame")
        actions.grid(row=3, column=0, sticky="e", pady=(14, 0))
        ttk.Button(actions, text="No", command=lambda: close_with(False), style="Nav.TButton").pack(side="right")
        ttk.Button(actions, text="Yes", command=lambda: close_with(True), style="Primary.TButton").pack(side="right", padx=(0, 8))

        dialog.update_idletasks()
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        dialog_w = dialog.winfo_width()
        dialog_h = dialog.winfo_height()
        pos_x = root_x + max(0, (root_w - dialog_w) // 2)
        pos_y = root_y + max(0, (root_h - dialog_h) // 2)
        dialog.geometry(f"+{pos_x}+{pos_y}")
        dialog.protocol("WM_DELETE_WINDOW", lambda: close_with(False))
        dialog.wait_window()
        return result["value"]

    def _remember_download_url(self, url: str):
        normalized = self._normalize_source_url(url)
        if not normalized:
            return
        self.download_history = [entry for entry in self.download_history if entry != normalized]
        self.download_history.append(normalized)
        self.download_history = self.download_history[-200:]
        self._save_settings()

    def _choose_yt_dlp_path(self):
        selected = filedialog.askopenfilename(title="Select yt-dlp.exe", filetypes=[("yt-dlp executable", "yt-dlp.exe"), ("Executable files", "*.exe"), ("All files", "*.*")])
        if not selected:
            selected = filedialog.askdirectory(title="Or select the folder containing yt-dlp.exe")
        if selected:
            normalized = self._normalize_tool_candidate("yt-dlp", selected)
            self.yt_dlp_path_var.set(normalized or selected)
            self._save_settings()
            self._refresh_dependency_check()

    def _clear_yt_dlp_path(self):
        self.yt_dlp_path_var.set("")
        self._save_settings()
        self._refresh_dependency_check()

    def _autodetect_tool_path(self, name: str):
        resolved = self._resolve_tool_path(name)
        if resolved:
            if name == "yt-dlp":
                self.yt_dlp_path_var.set(resolved)
            else:
                self.ffmpeg_path_var.set(resolved)
            self._save_settings()
            self._refresh_dependency_check()
            self._log(f"Detected {name} at: {resolved}")
            return
        seen = set()
        for candidate in self._tool_candidates(name):
            candidate_text = str(candidate)
            if candidate_text in seen:
                continue
            seen.add(candidate_text)
            try:
                if candidate.exists():
                    if name == "yt-dlp":
                        self.yt_dlp_path_var.set(candidate_text)
                    else:
                        self.ffmpeg_path_var.set(candidate_text)
                    self._save_settings()
                    self._refresh_dependency_check()
                    self._log(f"Detected {name} at: {candidate_text}")
                    return
            except Exception:
                continue
        messagebox.showerror(APP_TITLE, f"Could not auto-detect {name}. Use Browse to select it manually.")

    def _load_settings(self):
        try:
            if SETTINGS_FILE.exists():
                return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
        return {}

    def _save_settings(self):
        payload = {
            "import_folder": self.import_folder_var.get().strip(),
            "spotify_folder": self.spotify_folder_var.get().strip(),
            "yt_dlp_path": self.yt_dlp_path_var.get().strip(),
            "ffmpeg_path": self.ffmpeg_path_var.get().strip(),
            "mp3_quality": self.mp3_quality_var.get().strip(),
            "preview_volume": int(self.preview_volume_var.get()),
            "theme_mode": self.theme_mode_var.get().strip(),
            "playlist_downloads": self.playlist_var.get(),
            "open_folder_when_finished": self.open_folder_var.get(),
            "keep_temp_files": self.keep_temp_var.get(),
            "review_before_import": self.review_before_import_var.get(),
            "copy_to_spotify_folder": self.copy_to_spotify_folder_var.get(),
            "write_embedded_tags": self.write_tags_var.get(),
            "download_history": self.download_history[-200:],
        }
        try:
            SETTINGS_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except Exception as exc:
            self._log(f"Could not save settings: {exc}")

    def _add_dependency_link(self, parent, text: str, url: str):
        link = tk.Label(parent, text=text, fg="#0f6cbd", cursor="hand2", anchor="w", font=("Segoe UI", 10, "underline"))
        link.pack(anchor="w", pady=(6, 0))
        link.bind("<Button-1>", lambda _event, target=url: self._open_url(target))
        self.dependency_link_labels.append(link)

    def _open_url(self, url: str):
        try:
            webbrowser.open(url)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not open link:\n{url}\n\n{exc}")

    def _refresh_dependency_check(self):
        colors = {"connected": "#1f7a1f", "missing": "#b42318"}
        statuses = {
            "yt-dlp": self._probe_tool("yt-dlp")["found"],
            "ffmpeg": self._probe_tool("ffmpeg")["found"],
            "mutagen": MUTAGEN_AVAILABLE,
            "pyinstaller": self._tool_exists("pyinstaller"),
        }
        for key, is_connected in statuses.items():
            label = self.tool_status_rows.get(key)
            if not label:
                continue
            label.configure(
                text="Connected" if is_connected else "Not Connected",
                fg=colors["connected"] if is_connected else colors["missing"],
                bg=self.current_theme_colors.get("panel", self.root.cget("bg")) if getattr(self, "current_theme_colors", None) else self.root.cget("bg"),
            )

    def _probe_tool(self, name: str):
        resolved = self._resolve_tool_path(name)
        if not resolved:
            return {"found": False, "path": None, "detail": "not found"}
        version_arg = "-version" if name == "ffmpeg" else "--version"
        try:
            result = subprocess.run([resolved, version_arg], capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                output = (result.stdout or result.stderr or "").strip().splitlines()
                return {"found": True, "path": resolved, "detail": output[0] if output else "ok"}
            error_output = (result.stderr or result.stdout or "").strip()
            return {"found": False, "path": resolved, "detail": error_output or f"exit code {result.returncode}"}
        except Exception as exc:
            return {"found": False, "path": resolved, "detail": str(exc)}

    def _normalize_tool_candidate(self, name: str, value: str | Path | None):
        if not value:
            return None
        candidate = Path(value).expanduser()
        if candidate.is_dir():
            if name == "ffmpeg":
                candidate = candidate / "ffmpeg.exe"
            elif name == "yt-dlp":
                candidate = candidate / "yt-dlp.exe"
        try:
            return str(candidate.resolve(strict=False))
        except Exception:
            return str(candidate)

    def _tool_candidates(self, name: str):
        home = Path.home()
        if name == "yt-dlp":
            candidates = []
            custom_ytdlp = self._normalize_tool_candidate(name, self.yt_dlp_path_var.get().strip())
            if custom_ytdlp:
                candidates.append(Path(custom_ytdlp))
            candidates.extend([
                home / "AppData" / "Local" / "Microsoft" / "WinGet" / "Packages" / "yt-dlp.yt-dlp_Microsoft.Winget.Source_8wekyb3d8bbwe" / "yt-dlp.exe",
                home / "AppData" / "Local" / "Microsoft" / "WinGet" / "Links" / "yt-dlp.exe",
                home / "AppData" / "Local" / "Microsoft" / "WindowsApps" / "yt-dlp.exe",
            ])
            return candidates
        if name != "ffmpeg":
            return []

        candidates = []
        custom_ffmpeg = self._normalize_tool_candidate(name, self.ffmpeg_path_var.get().strip())
        if custom_ffmpeg:
            candidates.append(Path(custom_ffmpeg))

        candidates.extend([
            home / "AppData" / "Local" / "Microsoft" / "WinGet" / "Links" / "ffmpeg.exe",
            home / "AppData" / "Local" / "Microsoft" / "WindowsApps" / "ffmpeg.exe",
            home / "AppData" / "Local" / "Microsoft" / "WinGet" / "Packages" / "Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe" / "ffmpeg.exe",
            home / "AppData" / "Local" / "Microsoft" / "WinGet" / "Packages" / "Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe" / "ffmpeg" / "bin" / "ffmpeg.exe",
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "ffmpeg" / "bin" / "ffmpeg.exe",
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "ffmpeg" / "bin" / "ffmpeg.exe",
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "ShareX" / "ffmpeg.exe",
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "SteelSeries" / "GG" / "apps" / "moments" / "ffmpeg.exe",
            home / "scoop" / "shims" / "ffmpeg.exe",
        ])

        chocolatey_install = os.environ.get("ChocolateyInstall")
        if chocolatey_install:
            candidates.append(Path(chocolatey_install) / "bin" / "ffmpeg.exe")

        ffmpeg_root = home / "AppData" / "Local" / "Microsoft" / "WinGet" / "Packages" / "Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe"
        try:
            for child in ffmpeg_root.iterdir():
                candidates.append(child / "bin" / "ffmpeg.exe")
                candidates.append(child / "ffmpeg.exe")
        except Exception:
            pass

        return candidates

    def _resolve_tool_path(self, name: str):
        resolved = shutil.which(name)
        if resolved:
            return resolved

        candidates = self._tool_candidates(name)

        for candidate in candidates:
            try:
                if candidate.exists():
                    return str(candidate)
            except Exception:
                continue
        return None

    def _tool_exists(self, name: str) -> bool:
        return self._probe_tool(name)["found"]

    def _choose_import_folder(self):
        folder = filedialog.askdirectory(initialdir=self.import_folder_var.get() or DEFAULT_IMPORT_FOLDER)
        if folder:
            self.import_folder_var.set(folder)
            self._save_settings()

    def _choose_spotify_folder(self):
        folder = filedialog.askdirectory(initialdir=self.spotify_folder_var.get() or DEFAULT_SPOTIFY_FOLDER)
        if folder:
            self.spotify_folder_var.set(folder)
            self._save_settings()

    def _choose_ffmpeg_path(self):
        selected = filedialog.askopenfilename(
            title="Select ffmpeg.exe",
            filetypes=[("ffmpeg executable", "ffmpeg.exe"), ("Executable files", "*.exe"), ("All files", "*.*")],
        )
        if not selected:
            selected = filedialog.askdirectory(title="Or select the folder containing ffmpeg.exe")
        if selected:
            normalized = self._normalize_tool_candidate("ffmpeg", selected)
            self.ffmpeg_path_var.set(normalized or selected)
            self._save_settings()
            self._refresh_dependency_check()

    def _autodetect_ffmpeg_path(self):
        ffmpeg_path = self._resolve_tool_path("ffmpeg")
        if ffmpeg_path:
            self.ffmpeg_path_var.set(ffmpeg_path)
            self._save_settings()
            self._refresh_dependency_check()
            self._log(f"Detected ffmpeg at: {ffmpeg_path}")
            return

        seen = set()
        for candidate in self._tool_candidates("ffmpeg"):
            candidate_text = str(candidate)
            if candidate_text in seen:
                continue
            seen.add(candidate_text)
            try:
                if candidate.exists():
                    self.ffmpeg_path_var.set(candidate_text)
                    self._save_settings()
                    self._refresh_dependency_check()
                    self._log(f"Detected ffmpeg at: {candidate_text}")
                    return
            except Exception:
                continue

        messagebox.showerror(APP_TITLE, "Could not auto-detect ffmpeg.exe. Use Browse to select it manually.")

    def _clear_ffmpeg_path(self):
        self.ffmpeg_path_var.set("")
        self._save_settings()
        self._refresh_dependency_check()

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _format_log_message(self, message: str) -> str:
        text = (message or "").strip()
        if not text:
            return ""
        replacements = [
            ("[download] ", ""),
            ("[youtube] ", ""),
            ("[ExtractAudio] ", "Converting: "),
            ("[Metadata] ", "Metadata: "),
            ("[Merger] ", "Merging: "),
            ("[EmbedThumbnail] ", "Artwork: "),
        ]
        for source, target in replacements:
            text = text.replace(source, target)
        if text.startswith("Destination: "):
            text = f"Saving: {text[len('Destination: '):]}"
        if text.startswith("Deleting original file "):
            text = f"Cleanup: {text[len('Deleting original file '):]}"
        if text.startswith("ERROR: "):
            text = f"Error: {text[len('ERROR: '):]}"
        return text

    def _log(self, message: str):
        formatted = self._format_log_message(message)
        if formatted:
            self.log_queue.put(formatted)

    def _poll_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", message + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log_queue)

    def _kill_process_hard(self, process):
        if not process:
            return
        try:
            if process.poll() is not None:
                return
        except Exception:
            pass
        try:
            if sys.platform.startswith("win"):
                subprocess.run(
                    ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
            else:
                process.kill()
        except Exception:
            try:
                process.kill()
            except Exception:
                pass

    def _force_close_app(self):
        self.stop_requested.set()
        self._kill_process_hard(self.download_process)
        self._kill_process_hard(self.audio_preview_process)
        self.download_process = None
        self.audio_preview_process = None
        try:
            self.root.destroy()
        except Exception:
            pass
        os._exit(1)

    def _stop_download(self):
        self.stop_requested.set()
        self._set_download_state("stopping")
        process = self.download_process
        if process and process.poll() is None:
            try:
                process.terminate()
                self._log("Stopping download...")
            except Exception as exc:
                self._log(f"Could not stop download cleanly: {exc}")

    def _start_download(self):
        url = self.url_var.get().strip()
        import_folder = self.import_folder_var.get().strip()
        spotify_folder = self.spotify_folder_var.get().strip()

        if not url:
            messagebox.showerror(APP_TITLE, "Paste a YouTube video or playlist URL first.")
            return

        normalized_url = self._normalize_source_url(url)
        if normalized_url and normalized_url in self.download_history:
            if not self._confirm_duplicate_download(normalized_url):
                self._log("Download cancelled because the source was already downloaded before.")
                return

        yt_dlp_probe = self._probe_tool("yt-dlp")
        ffmpeg_probe = self._probe_tool("ffmpeg")
        if not yt_dlp_probe["found"] or not ffmpeg_probe["found"]:
            self._log(f"Tool check: yt-dlp -> {yt_dlp_probe['detail']}")
            self._log(f"Tool check: ffmpeg -> {ffmpeg_probe['detail']}")
            details = []
            if not yt_dlp_probe["found"]:
                details.append(f"yt-dlp: {yt_dlp_probe['detail']}")
            if not ffmpeg_probe["found"]:
                details.append(f"ffmpeg: {ffmpeg_probe['detail']}")
            self._refresh_dependency_check()
            messagebox.showerror(APP_TITLE, "Required tools check failed:\n\n" + "\n".join(details))
            return

        Path(import_folder).mkdir(parents=True, exist_ok=True)
        Path(spotify_folder).mkdir(parents=True, exist_ok=True)
        self._save_settings()
        self.stop_requested.clear()

        if self.worker and self.worker.is_alive():
            messagebox.showinfo(APP_TITLE, "A download is already running.")
            return

        self.last_downloaded_files = []
        self.last_metadata_rows = []
        self.last_processed_rows = []
        self._set_download_state("running")
        self._log("Starting download pipeline...")
        self.worker = threading.Thread(target=self._download_worker, args=(url, import_folder, spotify_folder), daemon=True)
        self.worker.start()
        self.root.after(500, self._watch_worker)

    def _watch_worker(self):
        if self.worker and self.worker.is_alive():
            self.root.after(500, self._watch_worker)
            return
        if self.stop_requested.is_set():
            self._set_download_state("idle")
        elif self.last_downloaded_files:
            self._set_download_state("done")
        else:
            self._set_download_state("idle")
        self.download_process = None

    def _download_worker(self, url: str, import_folder: str, spotify_folder: str):
        try:
            files = self._run_download(url, import_folder)
            self.last_downloaded_files = files
            rows = self._build_metadata_rows(files)
            self.last_metadata_rows = rows
            if self.stop_requested.is_set():
                self._log("Download stopped.")
                return
            self._log(f"Downloaded {len(files)} file(s).")
            if files:
                self._remember_download_url(url)

            if self.review_before_import_var.get() and rows:
                self.root.after(0, self._open_metadata_editor)
            elif self.copy_to_spotify_folder_var.get() and rows:
                self._process_and_send_rows(rows, spotify_folder)

            self._log("Pipeline finished.")
            if self.open_folder_var.get():
                self._open_folder(spotify_folder if self.copy_to_spotify_folder_var.get() else import_folder)
        except Exception as exc:
            if self.stop_requested.is_set():
                self._log("Download stopped.")
            else:
                self._set_download_state("error")
                self._log(f"ERROR: {exc}")
                self.root.after(0, lambda: messagebox.showerror(APP_TITLE, f"Download failed.\n\n{exc}"))

    def _open_folder(self, folder: str):
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as exc:
            self._log(f"Could not open folder automatically: {exc}")

    def _run_download(self, url: str, folder: str):
        yt_dlp_path = self._resolve_tool_path("yt-dlp")
        if not yt_dlp_path:
            raise RuntimeError("yt-dlp is not installed or could not be located.")
        ffmpeg_path = self._resolve_tool_path("ffmpeg")
        if not ffmpeg_path:
            raise RuntimeError("ffmpeg is not installed or could not be located.")

        before = {}
        for path in Path(folder).rglob("*"):
            if path.is_file() and path.suffix.lower() in AUDIO_EXTENSIONS:
                resolved = str(path.resolve())
                try:
                    before[resolved] = path.stat().st_mtime_ns
                except Exception:
                    before[resolved] = None
        output_template = os.path.join(folder, "%(playlist_title,Unknown Playlist)s", "%(artist,Unknown Artist)s", "%(album,Singles)s", "%(track_number,0>2)s - %(title)s.%(ext)s")

        cmd = [
            yt_dlp_path,
            "--no-warnings",
            "--newline",
            "--ignore-errors",
            "--extract-audio",
            "--audio-format", "mp3",
            "--audio-quality", self.mp3_quality_var.get(),
            "--write-thumbnail",
            "--embed-metadata",
            "--embed-thumbnail",
            "--convert-thumbnails", "jpg",
            "--add-metadata",
            "--ffmpeg-location", str(Path(ffmpeg_path).parent),
            "--parse-metadata", "%(uploader)s:%(artist)s",
            "--parse-metadata", "%(playlist_title)s:%(album)s",
            "--output", output_template,
            url,
        ]

        if not self.playlist_var.get():
            cmd.insert(-1, "--no-playlist")
        if not self.keep_temp_var.get():
            cmd.insert(-1, "--no-keep-video")

        self._log("Connecting to source...")

        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
        )
        self.download_process = process

        for line in iter(process.stdout.readline, ""):
            line = line.rstrip()
            if line:
                self._log(line)
            if self.stop_requested.is_set() and process.poll() is None:
                try:
                    process.terminate()
                except Exception:
                    pass
                break

        process.stdout.close()
        return_code = process.wait()
        self.download_process = None
        if self.stop_requested.is_set():
            raise RuntimeError("Download stopped by user.")
        if return_code != 0:
            raise RuntimeError(f"yt-dlp exited with code {return_code}")

        touched_files = []
        for path in Path(folder).rglob("*"):
            if not path.is_file() or path.suffix.lower() not in AUDIO_EXTENSIONS:
                continue
            resolved = str(path.resolve())
            try:
                mtime_ns = path.stat().st_mtime_ns
            except Exception:
                mtime_ns = None
            if resolved not in before or before.get(resolved) != mtime_ns:
                touched_files.append(resolved)

        if touched_files:
            return sorted(set(touched_files))

        fallback_files = []
        try:
            cutoff_ns = time.time_ns() - (15 * 60 * 1_000_000_000)
        except Exception:
            cutoff_ns = None
        for path in Path(folder).rglob("*"):
            if not path.is_file() or path.suffix.lower() not in AUDIO_EXTENSIONS:
                continue
            try:
                if cutoff_ns is None or path.stat().st_mtime_ns >= cutoff_ns:
                    fallback_files.append(str(path.resolve()))
            except Exception:
                continue
        return sorted(set(fallback_files))

    def _build_metadata_rows(self, files):
        rows = []
        for path_str in files:
            path = Path(path_str)
            stem = path.stem
            track = ""
            title = stem
            match = re.match(r"^(\d{1,3})\s*-\s*(.+)$", stem)
            if match:
                track = match.group(1)
                title = match.group(2)

            parts = path.parts
            playlist = parts[-4] if len(parts) >= 4 else "Unknown Playlist"
            artist = parts[-3] if len(parts) >= 3 else "Unknown Artist"
            album = parts[-2] if len(parts) >= 2 else "Singles"

            rows.append(
                {
                    "source_path": str(path),
                    "filename": path.name,
                    "title": title,
                    "artist": artist,
                    "album": album,
                    "track": track,
                    "playlist": playlist,
                    "artwork_path": self._find_thumbnail_for_audio(path),
                    "import_type": "playlist" if playlist and playlist != "Unknown Playlist" else "single",
                }
            )
        return rows

    def _load_rows_from_manifest(self, manifest_path: Path):
        rows = []
        if not manifest_path.exists():
            return rows
        try:
            with manifest_path.open("r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    rows.append(
                        {
                            "title": row.get("title", ""),
                            "artist": row.get("artist", ""),
                            "album": row.get("album", ""),
                            "track": row.get("track", ""),
                            "playlist": row.get("playlist", ""),
                            "source_path": row.get("source_path", ""),
                            "spotify_path": row.get("spotify_path", ""),
                            "artwork_path": row.get("artwork_path", ""),
                            "filename": Path(row.get("source_path", "")).name if row.get("source_path", "") else "",
                            "import_type": "playlist" if row.get("playlist", "") and row.get("playlist", "") != "Unknown Playlist" else "single",
                        }
                    )
        except Exception as exc:
            self._log(f"Could not read manifest {manifest_path}: {exc}")
        return rows

    def _reload_recent_state(self):
        spotify_folder = self.spotify_folder_var.get().strip()
        import_folder = self.import_folder_var.get().strip()

        processed_rows = []
        if spotify_folder:
            manifest_path = Path(spotify_folder) / "spotify_local_manifest.csv"
            processed_rows = self._load_rows_from_manifest(manifest_path)
            if processed_rows:
                self.last_processed_rows = processed_rows
                source_rows = []
                for row in processed_rows:
                    source_path = row.get("source_path", "")
                    if source_path and Path(source_path).exists():
                        source_rows.append(dict(row))
                self.last_metadata_rows = source_rows or [dict(r) for r in processed_rows]
                self._log(f"Reloaded {len(self.last_processed_rows)} recent file(s) from manifest.")
                return True

        if import_folder and Path(import_folder).exists():
            recent_files = []
            cutoff_ns = time.time_ns() - (24 * 60 * 60 * 1_000_000_000)
            for path in Path(import_folder).rglob("*"):
                if not path.is_file() or path.suffix.lower() not in AUDIO_EXTENSIONS:
                    continue
                try:
                    if path.stat().st_mtime_ns >= cutoff_ns:
                        recent_files.append(str(path.resolve()))
                except Exception:
                    continue
            if recent_files:
                self.last_metadata_rows = self._build_metadata_rows(sorted(set(recent_files)))
                self._log(f"Reloaded {len(self.last_metadata_rows)} recent source file(s) from download folder.")
                return True

        return False

    def _reload_recent_files_action(self):
        if self._reload_recent_state():
            self._populate_workspace_views()
            messagebox.showinfo(APP_TITLE, "Recent files were reloaded from the configured folders.")
            return
        messagebox.showinfo(APP_TITLE, "No recent files were found. Point the app at the correct download and Spotify folders in Settings, then try again.")

    def _populate_workspace_views(self):
        self._populate_metadata_tree()
        self._populate_files_tree()

    def _populate_metadata_tree(self):
        self.metadata_tree.delete(*self.metadata_tree.get_children())
        for index, row in enumerate(self.last_metadata_rows):
            self.metadata_tree.insert("", "end", iid=str(index), values=(
                row.get("filename", ""),
                row.get("title", ""),
                row.get("artist", ""),
                row.get("album", ""),
                row.get("track", ""),
                row.get("playlist", ""),
                row.get("import_type", "playlist"),
            ))
        if self.last_metadata_rows:
            if self.metadata_selected_index is None or self.metadata_selected_index >= len(self.last_metadata_rows):
                self.metadata_selected_index = 0
            self.metadata_tree.selection_set(str(self.metadata_selected_index))
            self._load_metadata_selection()
        else:
            self.metadata_selected_index = None
            self._clear_metadata_form()

    def _populate_files_tree(self):
        self.files_tree.delete(*self.files_tree.get_children())
        rows = self.last_processed_rows or self.last_metadata_rows
        self.workspace_rows = [dict(r) for r in rows]
        for index, row in enumerate(self.workspace_rows):
            spotify_path = row.get("spotify_path", "")
            source_path = row.get("source_path", "")
            status = []
            if spotify_path and Path(spotify_path).exists():
                status.append("spotify")
            if source_path and Path(source_path).exists():
                status.append("download")
            self.files_tree.insert("", "end", iid=str(index), values=(
                row.get("title", ""),
                row.get("artist", ""),
                row.get("album", ""),
                row.get("playlist", ""),
                " + ".join(status) if status else "missing",
                spotify_path,
                source_path,
            ))
        if self.workspace_rows:
            if self.files_selected_index is None or self.files_selected_index >= len(self.workspace_rows):
                self.files_selected_index = 0
            self.files_tree.selection_set(str(self.files_selected_index))
            self._load_files_selection()
        else:
            self.files_selected_index = None
            self._clear_files_form()

    def _clear_files_form(self):
        for var in (
            self.file_title_var,
            self.file_artist_var,
            self.file_album_var,
            self.file_track_var,
            self.file_playlist_var,
            self.file_source_path_var,
            self.file_spotify_path_var,
            self.file_artwork_path_var,
        ):
            var.set("")
        self.file_import_type_var.set("playlist")

    def _load_files_selection(self, _event=None):
        selected = self.files_tree.selection()
        if not selected:
            self.files_selected_index = None
            self._clear_files_form()
            return
        idx = int(selected[0])
        self.files_selected_index = idx
        row = self.workspace_rows[idx]
        self.file_title_var.set(row.get("title", ""))
        self.file_artist_var.set(row.get("artist", ""))
        self.file_album_var.set(row.get("album", ""))
        self.file_track_var.set(str(row.get("track", "")))
        self.file_playlist_var.set(row.get("playlist", ""))
        self.file_import_type_var.set(row.get("import_type", "playlist"))
        self.file_source_path_var.set(row.get("source_path", ""))
        self.file_spotify_path_var.set(row.get("spotify_path", ""))
        self.file_artwork_path_var.set(row.get("artwork_path", ""))

    def _sync_workspace_row(self, updated_row: dict):
        source_path = updated_row.get("source_path", "")
        spotify_path = updated_row.get("spotify_path", "")

        def matches(existing):
            return (
                (source_path and existing.get("source_path", "") == source_path)
                or (spotify_path and existing.get("spotify_path", "") == spotify_path)
            )

        for rows in (self.workspace_rows, self.last_processed_rows, self.last_metadata_rows):
            for index, existing in enumerate(rows):
                if matches(existing):
                    rows[index] = dict(existing) | dict(updated_row)

    def _apply_files_form(self):
        if self.files_selected_index is None or self.files_selected_index >= len(self.workspace_rows):
            messagebox.showinfo(APP_TITLE, "Select a recent file first.")
            return
        row = dict(self.workspace_rows[self.files_selected_index])
        row["title"] = clean_filename(self.file_title_var.get().strip(), "Unknown Title")
        row["artist"] = clean_filename(self.file_artist_var.get().strip(), "Unknown Artist")
        row["album"] = clean_filename(self.file_album_var.get().strip(), "Singles")
        row["track"] = self.file_track_var.get().strip()
        row["playlist"] = clean_filename(self.file_playlist_var.get().strip(), "Unknown Playlist")
        row["import_type"] = self.file_import_type_var.get().strip() if self.file_import_type_var.get().strip() in {"playlist", "single"} else "playlist"
        row["artwork_path"] = self.file_artwork_path_var.get().strip()
        self._sync_workspace_row(row)
        for path_key in ("source_path", "spotify_path"):
            target = row.get(path_key, "")
            if target and Path(target).exists() and self.write_tags_var.get():
                self._write_mp3_tags(Path(target), row)
        spotify_folder = self.spotify_folder_var.get().strip()
        if spotify_folder and self.last_processed_rows:
            self._write_manifest(self.last_processed_rows, spotify_folder)
        self._populate_files_tree()

    def _begin_metadata_edit(self, event):
        region = self.metadata_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.metadata_tree.identify_row(event.y)
        column_id = self.metadata_tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        x, y, width, height = self.metadata_tree.bbox(row_id, column_id)
        col_index = int(column_id.replace("#", "")) - 1
        keys = ["filename", "title", "artist", "album", "track", "playlist", "import_type"]
        key = keys[col_index]
        current_row = dict(self.last_metadata_rows[int(row_id)])
        if key == "filename":
            return
        if key == "import_type":
            combo = ttk.Combobox(self.metadata_tree, values=["playlist", "single"], state="readonly")
            combo.place(x=x, y=y, width=width, height=height)
            combo.set(current_row.get(key, "playlist"))
            combo.focus()
            def save_combo(_event=None):
                current_row[key] = combo.get().strip() or "playlist"
                self.last_metadata_rows[int(row_id)] = current_row
                self._populate_metadata_tree()
                combo.destroy()
            combo.bind("<<ComboboxSelected>>", save_combo)
            combo.bind("<FocusOut>", save_combo)
            return
        entry = ttk.Entry(self.metadata_tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, self.metadata_tree.item(row_id, "values")[col_index])
        entry.focus()
        entry.select_range(0, "end")
        def save_edit(_event=None):
            current_row[key] = entry.get().strip()
            self.last_metadata_rows[int(row_id)] = current_row
            self._populate_metadata_tree()
            entry.destroy()
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

    def _clear_metadata_form(self):
        for var in (
            self.meta_filename_var,
            self.meta_title_var,
            self.meta_artist_var,
            self.meta_album_var,
            self.meta_track_var,
            self.meta_playlist_var,
            self.meta_artwork_path_var,
            self.meta_source_path_var,
        ):
            var.set("")
        self.meta_import_type_var.set("playlist")
        self.metadata_thumbnail_image = None
        self.metadata_thumbnail_label.configure(image="", text="No artwork")

    def _load_metadata_selection(self, _event=None):
        selected = self.metadata_tree.selection()
        if not selected:
            self._clear_metadata_form()
            return
        idx = int(selected[0])
        self.metadata_selected_index = idx
        row = self.last_metadata_rows[idx]
        self.meta_filename_var.set(row.get("filename", ""))
        self.meta_title_var.set(row.get("title", ""))
        self.meta_artist_var.set(row.get("artist", ""))
        self.meta_album_var.set(row.get("album", ""))
        self.meta_track_var.set(str(row.get("track", "")))
        self.meta_playlist_var.set(row.get("playlist", ""))
        self.meta_import_type_var.set(row.get("import_type", "playlist"))
        self.meta_artwork_path_var.set(row.get("artwork_path", ""))
        self.meta_source_path_var.set(row.get("source_path", ""))
        self._load_metadata_thumbnail(row)

    def _load_metadata_thumbnail(self, row):
        artwork_path = row.get("artwork_path", "")
        if artwork_path and Path(artwork_path).exists():
            try:
                image = tk.PhotoImage(file=artwork_path)
                self.metadata_thumbnail_image = image
                self.metadata_thumbnail_label.configure(image=image, text="")
                return
            except Exception:
                pass
        source_path = row.get("source_path", "")
        if not source_path or not Path(source_path).exists() or not MUTAGEN_AVAILABLE:
            self.metadata_thumbnail_image = None
            self.metadata_thumbnail_label.configure(image="", text="No artwork")
            return
        try:
            tags = ID3(str(source_path))
            frames = tags.getall("APIC")
            if not frames:
                raise ValueError("No APIC frame")
            image_data = base64.b64encode(frames[0].data).decode("ascii")
            image = tk.PhotoImage(data=image_data)
            self.metadata_thumbnail_image = image
            self.metadata_thumbnail_label.configure(image=image, text="")
        except Exception:
            self.metadata_thumbnail_image = None
            self.metadata_thumbnail_label.configure(image="", text="No artwork")

    def _find_thumbnail_for_audio(self, audio_path: Path | str) -> str:
        if not audio_path:
            return ""
        path = Path(audio_path)
        if not path.exists() and not path.parent.exists():
            return ""
        for suffix in (".jpg", ".jpeg", ".png", ".webp", ".bmp"):
            candidate = path.with_suffix(suffix)
            if candidate.exists():
                return str(candidate)
        return ""

    def _choose_artwork_file(self) -> str:
        return filedialog.askopenfilename(
            title="Choose artwork image",
            filetypes=[
                ("Image files", "*.jpg;*.jpeg;*.png;*.webp;*.bmp"),
                ("JPEG files", "*.jpg;*.jpeg"),
                ("PNG files", "*.png"),
                ("All files", "*.*"),
            ],
        )

    def _use_metadata_thumbnail_artwork(self):
        if self.metadata_selected_index is None or self.metadata_selected_index >= len(self.last_metadata_rows):
            messagebox.showinfo(APP_TITLE, "Select a track first.")
            return
        row = self.last_metadata_rows[self.metadata_selected_index]
        thumbnail = self._find_thumbnail_for_audio(row.get("source_path", ""))
        if not thumbnail:
            messagebox.showinfo(APP_TITLE, "No saved YouTube thumbnail was found for this track.")
            return
        self.meta_artwork_path_var.set(thumbnail)
        self._load_metadata_thumbnail(dict(row) | {"artwork_path": thumbnail})

    def _choose_metadata_artwork(self):
        selected = self._choose_artwork_file()
        if selected:
            self.meta_artwork_path_var.set(selected)
            current = {}
            if self.metadata_selected_index is not None and self.metadata_selected_index < len(self.last_metadata_rows):
                current = dict(self.last_metadata_rows[self.metadata_selected_index])
            self._load_metadata_thumbnail(current | {"artwork_path": selected})

    def _clear_metadata_artwork(self):
        self.meta_artwork_path_var.set("")
        if self.metadata_selected_index is not None and self.metadata_selected_index < len(self.last_metadata_rows):
            self._load_metadata_thumbnail(dict(self.last_metadata_rows[self.metadata_selected_index]) | {"artwork_path": ""})
        else:
            self.metadata_thumbnail_image = None
            self.metadata_thumbnail_label.configure(image="", text="No artwork")

    def _use_files_thumbnail_artwork(self):
        if self.files_selected_index is None or self.files_selected_index >= len(self.workspace_rows):
            messagebox.showinfo(APP_TITLE, "Select a recent file first.")
            return
        row = self.workspace_rows[self.files_selected_index]
        thumbnail = self._find_thumbnail_for_audio(row.get("source_path", "") or row.get("spotify_path", ""))
        if not thumbnail:
            messagebox.showinfo(APP_TITLE, "No saved YouTube thumbnail was found for this file.")
            return
        self.file_artwork_path_var.set(thumbnail)

    def _choose_files_artwork(self):
        selected = self._choose_artwork_file()
        if selected:
            self.file_artwork_path_var.set(selected)

    def _clear_files_artwork(self):
        self.file_artwork_path_var.set("")

    def _apply_metadata_form(self):
        if self.metadata_selected_index is None or self.metadata_selected_index >= len(self.last_metadata_rows):
            messagebox.showinfo(APP_TITLE, "Select a track first.")
            return
        row = dict(self.last_metadata_rows[self.metadata_selected_index])
        row["filename"] = clean_filename(self.meta_filename_var.get().strip(), row.get("filename", "Track"))
        row["title"] = clean_filename(self.meta_title_var.get().strip(), "Unknown Title")
        row["artist"] = clean_filename(self.meta_artist_var.get().strip(), "Unknown Artist")
        row["album"] = clean_filename(self.meta_album_var.get().strip(), "Singles")
        row["track"] = self.meta_track_var.get().strip()
        row["playlist"] = clean_filename(self.meta_playlist_var.get().strip(), "Unknown Playlist")
        row["import_type"] = self.meta_import_type_var.get().strip() if self.meta_import_type_var.get().strip() in {"playlist", "single"} else "playlist"
        row["artwork_path"] = self.meta_artwork_path_var.get().strip()
        self.last_metadata_rows[self.metadata_selected_index] = row
        self._populate_metadata_tree()

    def _metadata_auto_clean(self):
        cleaned = []
        for row in self.last_metadata_rows:
            updated = dict(row)
            for key, fallback in (("title", "Unknown Title"), ("artist", "Unknown Artist"), ("album", "Singles"), ("playlist", "Unknown Playlist")):
                updated[key] = clean_filename(updated.get(key, ""), fallback)
            cleaned.append(updated)
        self.last_metadata_rows = cleaned
        self._populate_metadata_tree()

    def _metadata_fill_tracks(self):
        for idx, row in enumerate(self.last_metadata_rows, start=1):
            if not str(row.get("track", "")).strip():
                row["track"] = str(idx)
        self._populate_metadata_tree()

    def _save_metadata_from_workspace(self):
        self._apply_metadata_form()
        self._populate_metadata_tree()
        self._log("Metadata changes saved.")
        if self.copy_to_spotify_folder_var.get() and self.last_metadata_rows:
            self._process_and_send_rows(self.last_metadata_rows, self.spotify_folder_var.get().strip())

    def _delete_selected_from_workspace(self):
        selected = self.files_tree.selection()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Select one or more recent files first.")
            return
        rows = [self.workspace_rows[int(iid)] for iid in selected]
        self._delete_output_rows(rows)
        self._populate_workspace_views()

    def _open_trim_from_selection(self):
        rows = self.workspace_rows or (self.last_processed_rows or self.last_metadata_rows)
        selected = self.files_tree.selection()
        if not selected and rows:
            selected = ("0",)
        if not selected:
            if not self._reload_recent_state():
                messagebox.showinfo(APP_TITLE, "No recent files are available to trim.")
                return
            self._populate_workspace_views()
            rows = self.workspace_rows or (self.last_processed_rows or self.last_metadata_rows)
            if not rows:
                messagebox.showinfo(APP_TITLE, "No recent files are available to trim.")
                return
            selected = ("0",)
        row = rows[int(selected[0])]
        self._load_trim_workspace(row)

    def _resolve_ffplay_path(self):
        ffmpeg_path = self._resolve_tool_path("ffmpeg")
        if not ffmpeg_path:
            return None
        ffplay_path = Path(ffmpeg_path).with_name("ffplay.exe")
        return str(ffplay_path) if ffplay_path.exists() else None

    def _stop_audio_preview(self):
        if self.audio_preview_restart_after:
            self.root.after_cancel(self.audio_preview_restart_after)
            self.audio_preview_restart_after = None
        if self.audio_preview_process and self.audio_preview_process.poll() is None:
            try:
                self.audio_preview_process.terminate()
            except Exception:
                pass
        self.audio_preview_process = None
        self.audio_preview_range = None
        self.trim_preview_state_var.set("Preview stopped.")

    def _play_audio_range(self, start_time=None, end_time=None):
        if not self.trim_row:
            return
        ffplay_path = self._resolve_ffplay_path()
        audio_path = self.trim_row.get("source_path") or self.trim_row.get("spotify_path")
        if not ffplay_path or not audio_path or not Path(audio_path).exists():
            messagebox.showinfo(APP_TITLE, "ffplay was not found next to ffmpeg, so preview playback is unavailable.")
            self.trim_preview_state_var.set("Preview unavailable.")
            return
        self._stop_audio_preview()
        cmd = [ffplay_path, "-nodisp", "-autoexit", "-loglevel", "quiet"]
        preview_gain = max(0.0, float(self.preview_volume_var.get())) / 100.0
        cmd.extend(["-af", f"volume={preview_gain:.2f}"])
        if start_time is not None:
            cmd.extend(["-ss", f"{start_time:.2f}"])
        if end_time is not None and start_time is not None:
            cmd.extend(["-t", f"{max(0.1, end_time - start_time):.2f}"])
        cmd.append(audio_path)
        try:
            self.audio_preview_process = subprocess.Popen(cmd)
            self.audio_preview_range = (start_time, end_time)
            if start_time is None:
                self.trim_preview_state_var.set("Previewing full track.")
            else:
                self.trim_preview_state_var.set(f"Previewing {start_time:.2f}s to {end_time:.2f}s.")
        except Exception as exc:
            self.audio_preview_range = None
            self.trim_preview_state_var.set("Preview failed.")
            messagebox.showerror(APP_TITLE, f"Could not start preview playback.\n\n{exc}")

    def _preview_volume_text_var(self):
        if not hasattr(self, "_preview_volume_label_var"):
            self._preview_volume_label_var = tk.StringVar()
            self.preview_volume_var.trace_add("write", self._update_preview_volume_label)
            self._update_preview_volume_label()
        return self._preview_volume_label_var

    def _update_preview_volume_label(self, *_args):
        if hasattr(self, "_preview_volume_label_var"):
            self._preview_volume_label_var.set(f"{int(self.preview_volume_var.get())}%")

    def _on_preview_volume_change(self, value):
        self.preview_volume_var.set(int(round(float(value))))
        self._save_settings()
        if self.audio_preview_process and self.audio_preview_process.poll() is None and self.audio_preview_range:
            if self.audio_preview_restart_after:
                self.root.after_cancel(self.audio_preview_restart_after)
            self.audio_preview_restart_after = self.root.after(150, self._restart_audio_preview_with_volume)

    def _restart_audio_preview_with_volume(self):
        self.audio_preview_restart_after = None
        if not self.audio_preview_range:
            return
        start_time, end_time = self.audio_preview_range
        self._play_audio_range(start_time, end_time)

    def _play_trim_selection(self):
        self._play_audio_range(self.trim_start_var.get(), self.trim_end_var.get())

    def _play_selected_trim_clip(self):
        selected = self.trim_tree.selection()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Select a clip in the trim list first.")
            return
        segment = self.trim_segments[int(selected[0])]
        self._play_audio_range(segment["start"], segment["end"])

    def _play_trim_full_track(self):
        self._play_audio_range()

    def _open_metadata_editor(self):
        if not self.last_metadata_rows:
            self._reload_recent_state()
        if not self.last_metadata_rows:
            messagebox.showinfo(APP_TITLE, "No downloaded files are loaded yet. Point the app at your download/Spotify folders in Settings, then try again.")
            return
        self._populate_metadata_tree()
        self._set_workspace_status("Metadata", f"Editing {len(self.last_metadata_rows)} track(s) inside the app")
        self._show_workspace_panel("metadata")

    def _open_output_file_manager(self):
        if not self.last_processed_rows and not self.last_metadata_rows:
            self._reload_recent_state()
        if not self.last_processed_rows and not self.last_metadata_rows:
            messagebox.showinfo(APP_TITLE, "No recent files are loaded yet. Point the app at your download/Spotify folders in Settings, then try again.")
            return
        self._populate_files_tree()
        self._set_workspace_status("Recent Files", f"Browsing {len(self.workspace_rows)} recent file(s) inside the app")
        self._show_workspace_panel("files")

    def _open_trim_editor(self, row):
        ffmpeg_path = self._resolve_tool_path("ffmpeg")
        if not ffmpeg_path:
            messagebox.showerror(APP_TITLE, "ffmpeg must be installed before trimming audio.")
            self._refresh_dependency_check()
            return

        audio_path = Path(row.get("source_path") or row.get("spotify_path") or "")
        if not audio_path.exists():
            messagebox.showerror(APP_TITLE, "The selected file no longer exists.")
            return
        self._load_trim_workspace(row)

    def _safe_rel_path(self, row):
        playlist = clean_filename(row.get("playlist"), "Unknown Playlist")
        artist = clean_filename(row.get("artist"), "Unknown Artist")
        album = clean_filename(row.get("album"), "Singles")
        title = clean_filename(row.get("title"), "Unknown Title")
        track_value = str(row.get("track", "")).strip()
        track = clean_filename(track_value.zfill(2) if track_value else "00", "00")
        ext = Path(row["source_path"]).suffix or ".mp3"
        filename = f"{track} - {title}{ext}"

        import_type = row.get("import_type", "playlist")
        if import_type == "single":
            return Path(artist) / album / filename
        return Path(playlist) / artist / album / filename

    def _load_trim_workspace(self, row):
        self.trim_row = dict(row)
        self.trim_segments = []
        self._stop_audio_preview()
        audio_path = Path(row.get("source_path") or row.get("spotify_path") or "")
        ffmpeg_path = self._resolve_tool_path("ffmpeg")
        if not audio_path.exists() or not ffmpeg_path:
            messagebox.showerror(APP_TITLE, "Could not prepare the selected file for trimming.")
            return
        cmd = [ffmpeg_path, "-v", "error", "-i", str(audio_path), "-ac", "1", "-ar", "2000", "-f", "s16le", "-"]
        result = subprocess.run(cmd, capture_output=True, timeout=120)
        if result.returncode != 0 or not result.stdout:
            messagebox.showerror(APP_TITLE, "Could not decode audio samples for trimming.")
            return
        from array import array as _array
        samples = _array("h")
        samples.frombytes(result.stdout)
        if not samples:
            messagebox.showerror(APP_TITLE, "Could not decode audio samples for trimming.")
            return
        self.trim_duration = len(samples) / 2000.0
        step = max(1, len(samples) // 600)
        self.trim_waveform = [max(abs(samples[idx]) for idx in range(start, min(len(samples), start + step))) / 32768.0 for start in range(0, len(samples), step)][:600]
        if len(self.trim_waveform) < 600:
            self.trim_waveform.extend([0.0] * (600 - len(self.trim_waveform)))
        self.trim_start_var.set(0.0)
        self.trim_end_var.set(round(self.trim_duration, 2))
        self.trim_status_var.set(f"Trim file: {audio_path.name} | Duration: {self.trim_duration:.2f}s")
        self.trim_preview_state_var.set("Preview stopped.")
        self._load_trim_form(self._trim_segment_defaults(1))
        self._set_workspace_status("Trim", f"Trimming {audio_path.name} inside the app")
        self._refresh_trim_tree()
        self._draw_trim_waveform()
        self._show_workspace_panel("trim")

    def _draw_trim_waveform(self):
        self.trim_canvas.delete("all")
        self.trim_range_canvas.delete("all")
        if not self.trim_waveform:
            return
        width = int(self.trim_canvas["width"])
        height = int(self.trim_canvas["height"])
        mid = height / 2
        x_step = width / len(self.trim_waveform)
        for index, amplitude in enumerate(self.trim_waveform):
            amp_height = max(1, amplitude * (height * 0.45))
            x = index * x_step
            self.trim_canvas.create_line(x, mid - amp_height, x, mid + amp_height, fill="#7fb3d5")
        start_x = int((self.trim_start_var.get() / self.trim_duration) * width) if self.trim_duration else 0
        end_x = int((self.trim_end_var.get() / self.trim_duration) * width) if self.trim_duration else width
        self.trim_canvas.create_rectangle(0, 0, start_x, height, fill="#0b0f12", stipple="gray50", outline="")
        self.trim_canvas.create_rectangle(end_x, 0, width, height, fill="#0b0f12", stipple="gray50", outline="")
        for segment in self.trim_segments:
            seg_start = int((segment["start"] / self.trim_duration) * width) if self.trim_duration else 0
            seg_end = int((segment["end"] / self.trim_duration) * width) if self.trim_duration else width
            self.trim_canvas.create_rectangle(seg_start, 0, seg_end, height, outline="#6dd17c", width=1)
        self.trim_canvas.create_line(start_x, 0, start_x, height, fill="#00c2ff", width=3)
        self.trim_canvas.create_line(end_x, 0, end_x, height, fill="#00c2ff", width=3)
        self._draw_trim_range_bar()
        self.trim_selection_var.set(f"Selection: {self.trim_start_var.get():.2f}s - {self.trim_end_var.get():.2f}s")

    def _draw_trim_range_bar(self):
        colors = self.current_theme_colors or {}
        trough = colors.get("panel_alt", "#c7d2df")
        fill = colors.get("accent", "#4ca8ff")
        handle = colors.get("fg", "#19202a")
        text = colors.get("fg", "#19202a")
        width = int(self.trim_range_canvas["width"])
        height = int(self.trim_range_canvas["height"])
        start_x = int((self.trim_start_var.get() / self.trim_duration) * width) if self.trim_duration else 0
        end_x = int((self.trim_end_var.get() / self.trim_duration) * width) if self.trim_duration else width
        line_y1 = 18
        line_y2 = height - 10
        self.trim_range_canvas.create_rectangle(0, line_y1, width, line_y2, fill=trough, outline="")
        self.trim_range_canvas.create_rectangle(start_x, line_y1, end_x, line_y2, fill=fill, outline="")
        self.trim_range_canvas.create_rectangle(max(0, start_x - 2), line_y1 - 4, min(width, start_x + 2), line_y2 + 4, fill=handle, outline="")
        self.trim_range_canvas.create_rectangle(max(0, end_x - 2), line_y1 - 4, min(width, end_x + 2), line_y2 + 4, fill=handle, outline="")
        self.trim_range_canvas.create_text(6, 10, text=f"{self.trim_start_var.get():.2f}s", anchor="w", fill=text, font=("Segoe UI", 9, "bold"))
        self.trim_range_canvas.create_text(width - 6, 10, text=f"{self.trim_end_var.get():.2f}s", anchor="e", fill=text, font=("Segoe UI", 9, "bold"))

    def _set_trim_from_canvas(self, which, x_value, width):
        if not self.trim_duration:
            return
        time_value = round(max(0, min(width, x_value)) / width * self.trim_duration, 2)
        if which == "start":
            self.trim_start_var.set(min(time_value, max(0.0, self.trim_end_var.get() - 0.1)))
        else:
            self.trim_end_var.set(max(time_value, min(self.trim_duration, self.trim_start_var.get() + 0.1)))
        self._draw_trim_waveform()

    def _on_trim_canvas_press(self, event):
        if not self.trim_duration:
            return
        width = int(event.widget["width"])
        start_x = int((self.trim_start_var.get() / self.trim_duration) * width)
        end_x = int((self.trim_end_var.get() / self.trim_duration) * width)
        self.trim_drag_target = "start" if abs(event.x - start_x) <= abs(event.x - end_x) else "end"
        self._set_trim_from_canvas(self.trim_drag_target, event.x, width)

    def _on_trim_canvas_drag(self, event):
        if self.trim_drag_target:
            self._set_trim_from_canvas(self.trim_drag_target, event.x, int(event.widget["width"]))

    def _on_trim_canvas_release(self, _event):
        self.trim_drag_target = None

    def _trim_segment_defaults(self, index):
        base_title = clean_filename(self.trim_row.get("title", ""), "Clip")
        return {
            "start": round(self.trim_start_var.get(), 2),
            "end": round(self.trim_end_var.get(), 2),
            "title": f"{base_title} Part {index}",
            "artist": clean_filename(self.trim_row.get("artist", ""), "Unknown Artist"),
            "album": clean_filename(self.trim_row.get("album", ""), "Singles"),
            "track": str(index),
            "playlist": clean_filename(self.trim_row.get("playlist", ""), "Unknown Playlist"),
            "import_type": self.trim_row.get("import_type", "playlist"),
        }

    def _update_trim_selection_status(self, selected_index=None):
        if selected_index is None or selected_index < 0 or selected_index >= len(self.trim_segments):
            self.trim_selected_clip_var.set("Editing: new clip")
            if hasattr(self, "trim_save_button"):
                self.trim_save_button.configure(text="Save Clip Settings", state="disabled")
            return
        segment = self.trim_segments[selected_index]
        clip_name = segment.get("title", "").strip() or f"Clip {selected_index + 1}"
        self.trim_selected_clip_var.set(f"Editing clip {selected_index + 1}: {clip_name}")
        if hasattr(self, "trim_save_button"):
            self.trim_save_button.configure(text=f"Save Clip Settings for Clip {selected_index + 1}", state="normal")

    def _clear_trim_form(self):
        self.trim_title_var.set("")
        self.trim_artist_meta_var.set("")
        self.trim_album_meta_var.set("")
        self.trim_track_meta_var.set("")
        self.trim_playlist_meta_var.set("")
        self.trim_import_type_meta_var.set("playlist")
        self._update_trim_selection_status()

    def _load_trim_form(self, segment: dict):
        self.trim_title_var.set(segment.get("title", ""))
        self.trim_artist_meta_var.set(segment.get("artist", ""))
        self.trim_album_meta_var.set(segment.get("album", ""))
        self.trim_track_meta_var.set(str(segment.get("track", "")))
        self.trim_playlist_meta_var.set(segment.get("playlist", ""))
        self.trim_import_type_meta_var.set(segment.get("import_type", "playlist"))

    def _trim_segment_from_form(self, fallback_index: int):
        base = self._trim_segment_defaults(fallback_index)
        base["title"] = clean_filename(self.trim_title_var.get().strip(), base["title"])
        base["artist"] = clean_filename(self.trim_artist_meta_var.get().strip(), base["artist"])
        base["album"] = clean_filename(self.trim_album_meta_var.get().strip(), base["album"])
        base["track"] = self.trim_track_meta_var.get().strip() or base["track"]
        base["playlist"] = clean_filename(self.trim_playlist_meta_var.get().strip(), base["playlist"])
        selected_type = self.trim_import_type_meta_var.get().strip()
        base["import_type"] = selected_type if selected_type in {"playlist", "single"} else base["import_type"]
        return base

    def _refresh_trim_tree(self):
        selected = self.trim_tree.selection()
        selected_id = selected[0] if selected else None
        self.trim_tree.delete(*self.trim_tree.get_children())
        for index, segment in enumerate(self.trim_segments):
            self.trim_tree.insert("", "end", iid=str(index), values=(
                f"{segment['start']:.2f}",
                f"{segment['end']:.2f}",
                segment["title"],
                segment["track"],
            ))
        if self.trim_segments:
            if selected_id is not None and selected_id.isdigit():
                new_index = min(int(selected_id), len(self.trim_segments) - 1)
            else:
                new_index = 0
            self.trim_tree.selection_set(str(new_index))
            self.trim_tree.focus(str(new_index))
            self._update_trim_selection_status(new_index)
        else:
            self._update_trim_selection_status()
        self._draw_trim_waveform()

    def _add_trim_segment(self):
        if self.trim_end_var.get() - self.trim_start_var.get() < 0.2:
            messagebox.showerror(APP_TITLE, "Clip selection must be at least 0.2 seconds long.")
            return
        segment = self._trim_segment_from_form(len(self.trim_segments) + 1)
        self.trim_segments.append(segment)
        self._refresh_trim_tree()
        new_index = len(self.trim_segments) - 1
        self.trim_tree.selection_set(str(new_index))
        self.trim_tree.focus(str(new_index))
        self._load_trim_form(segment)

    def _update_trim_segment(self):
        selected = self.trim_tree.selection()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Select a clip in the trim list first.")
            return
        idx = int(selected[0])
        updated = self._trim_segment_from_form(idx + 1)
        self.trim_segments[idx] = updated
        self._refresh_trim_tree()
        self.trim_tree.selection_set(str(idx))
        self.trim_tree.focus(str(idx))
        self._load_selected_trim_segment()

    def _save_selected_trim_segment(self):
        self._update_trim_segment()

    def _remove_trim_segment(self):
        selected = self.trim_tree.selection()
        if not selected:
            messagebox.showinfo(APP_TITLE, "Select a clip in the trim list first.")
            return
        idx = int(selected[0])
        del self.trim_segments[idx]
        self._refresh_trim_tree()
        if self.trim_segments:
            new_idx = min(idx, len(self.trim_segments) - 1)
            self.trim_tree.selection_set(str(new_idx))
            self.trim_tree.focus(str(new_idx))
            self._load_trim_form(self.trim_segments[new_idx])
        else:
            self._load_trim_form(self._trim_segment_defaults(1))

    def _load_selected_trim_segment(self, _event=None):
        selected = self.trim_tree.selection()
        if not selected:
            self._update_trim_selection_status()
            return
        idx = int(selected[0])
        segment = self.trim_segments[idx]
        self.trim_start_var.set(segment["start"])
        self.trim_end_var.set(segment["end"])
        self._load_trim_form(segment)
        self._update_trim_selection_status(idx)
        self._draw_trim_waveform()

    def _begin_trim_segment_edit(self, event):
        region = self.trim_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.trim_tree.identify_row(event.y)
        column_id = self.trim_tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        x, y, width, height = self.trim_tree.bbox(row_id, column_id)
        col_index = int(column_id.replace("#", "")) - 1
        keys = ["start", "end", "title", "track"]
        key = keys[col_index]
        current = dict(self.trim_segments[int(row_id)])
        entry = ttk.Entry(self.trim_tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, self.trim_tree.item(row_id, "values")[col_index])
        entry.focus()
        entry.select_range(0, "end")
        def save_edit(_event=None):
            value = entry.get().strip()
            if key in {"start", "end"}:
                try:
                    value = round(float(value), 2)
                except Exception:
                    value = current[key]
            current[key] = value
            if current["end"] <= current["start"]:
                current["end"] = round(current["start"] + 0.1, 2)
            self.trim_segments[int(row_id)] = current
            self._refresh_trim_tree()
            self.trim_tree.selection_set(row_id)
            self.trim_tree.focus(row_id)
            self._load_selected_trim_segment()
            entry.destroy()
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)

    def _save_trim_workspace(self):
        if not self.trim_row:
            return
        if not self.trim_segments:
            messagebox.showerror(APP_TITLE, "Add at least one clip before saving.")
            return
        normalized = []
        for index, segment in enumerate(self.trim_segments, start=1):
            clip = dict(segment)
            clip["title"] = clean_filename(clip.get("title", ""), f"Clip {index}")
            clip["artist"] = clean_filename(clip.get("artist", ""), "Unknown Artist")
            clip["album"] = clean_filename(clip.get("album", ""), "Singles")
            clip["playlist"] = clean_filename(clip.get("playlist", ""), "Unknown Playlist")
            clip["track"] = str(clip.get("track", "")).strip() or str(index)
            clip["import_type"] = clip.get("import_type", "playlist") if clip.get("import_type", "playlist") in {"playlist", "single"} else "playlist"
            normalized.append(clip)
        self._trim_output_row(self.trim_row, normalized)
        self._stop_audio_preview()
        self._populate_workspace_views()
        self._set_workspace_status("Trim", f"Saved {len(normalized)} clip(s) inside the app")

    def _copy_with_cleanup(self, source_path: str, destination_path: Path):
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, destination_path)

    def _run_ffmpeg_trim(self, input_path: Path, output_path: Path, start_time: float, end_time: float):
        ffmpeg_path = self._resolve_tool_path("ffmpeg")
        if not ffmpeg_path:
            raise RuntimeError("ffmpeg is not installed or could not be located.")

        bitrate = self.mp3_quality_var.get().strip() or "320"
        cmd = [
            ffmpeg_path,
            "-y",
            "-hide_banner",
            "-loglevel", "error",
            "-ss", f"{start_time:.2f}",
            "-to", f"{end_time:.2f}",
            "-i", str(input_path),
            "-vn",
            "-c:a", "libmp3lame",
            "-b:a", f"{bitrate}k",
            str(output_path),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or "ffmpeg trim command failed.")

    def _trim_output_row(self, row, segments):
        source_path = Path(row.get("source_path", ""))
        spotify_path = Path(row.get("spotify_path", ""))
        existing_targets = [path for path in (source_path, spotify_path) if str(path).strip() and path.exists()]
        if not existing_targets:
            raise RuntimeError("No output files exist for the selected row.")
        input_path = source_path if source_path.exists() else spotify_path
        if not input_path.exists():
            raise RuntimeError("The selected source file does not exist.")
        if not segments:
            raise RuntimeError("No clips were provided.")

        segment_rows = []
        for index, segment in enumerate(segments, start=1):
            segment_row = dict(row)
            segment_row.update(segment)
            segment_row["track"] = str(segment_row.get("track", "")).strip() or str(index)
            segment_row["artwork_path"] = segment_row.get("artwork_path", "") or self._find_thumbnail_for_audio(input_path)
            segment_title = clean_filename(segment_row.get("title", ""), f"Clip {index}")
            ext = source_path.suffix or input_path.suffix or ".mp3"
            source_parent = source_path.parent if source_path.exists() else input_path.parent
            output_source = source_parent / f"{str(segment_row['track']).zfill(2)} - {segment_title}{ext}"
            temp_source = output_source.with_name(f"{output_source.stem}.tmp{output_source.suffix}")
            if temp_source.exists():
                temp_source.unlink()
            self._run_ffmpeg_trim(input_path, temp_source, segment_row["start"], segment_row["end"])
            temp_source.replace(output_source)
            segment_row["source_path"] = str(output_source)
            if self.write_tags_var.get():
                self._write_mp3_tags(output_source, segment_row)
            self._log(f"Saved clip source file: {output_source}")

            if spotify_path.exists():
                spotify_destination = Path(self.spotify_folder_var.get().strip()) / self._safe_rel_path(segment_row)
                self._copy_with_cleanup(str(output_source), spotify_destination)
                if self.write_tags_var.get():
                    self._write_mp3_tags(spotify_destination, segment_row)
                segment_row["spotify_path"] = str(spotify_destination)
                self._log(f"Saved clip Spotify file: {spotify_destination}")
            else:
                segment_row["spotify_path"] = ""
            segment_rows.append(segment_row)

        preserved_paths = {r.get("source_path", "") for r in segment_rows} | {r.get("spotify_path", "") for r in segment_rows}
        for target in existing_targets:
            try:
                if target.exists() and str(target) not in preserved_paths:
                    target.unlink()
                    self._log(f"Removed original trimmed source: {target}")
            except Exception as exc:
                self._log(f"Could not remove original file {target}: {exc}")

        self.last_metadata_rows = [r for r in self.last_metadata_rows if r.get("source_path") != row.get("source_path")]
        self.last_metadata_rows.extend(segment_rows)
        self.last_processed_rows = [
            r for r in self.last_processed_rows
            if r.get("source_path") != row.get("source_path") and r.get("spotify_path") != row.get("spotify_path")
        ]
        processed_segments = [r for r in segment_rows if r.get("spotify_path")]
        if processed_segments:
            self.last_processed_rows.extend(processed_segments)
        spotify_folder = self.spotify_folder_var.get().strip()
        if spotify_folder and self.last_processed_rows:
            self._write_manifest(self.last_processed_rows, spotify_folder)
        elif spotify_folder:
            manifest = Path(spotify_folder) / "spotify_local_manifest.csv"
            if manifest.exists():
                manifest.unlink()
        messagebox.showinfo(APP_TITLE, f"Saved {len(segment_rows)} clip(s).")

    def _write_mp3_tags(self, mp3_path: Path, row: dict):
        if not MUTAGEN_AVAILABLE:
            self._log("mutagen is not installed, so embedded tag writing was skipped.")
            return
        if mp3_path.suffix.lower() != ".mp3":
            self._log(f"Skipping tag write for non-MP3 file: {mp3_path.name}")
            return

        title = row.get("title", "") or ""
        artist = row.get("artist", "") or ""
        album = row.get("album", "") or ""
        track = str(row.get("track", "") or "")

        try:
            try:
                audio = EasyID3(str(mp3_path))
            except Exception:
                audio_file = MP3(str(mp3_path))
                audio_file.add_tags()
                audio_file.save()
                audio = EasyID3(str(mp3_path))

            audio["title"] = [title]
            audio["artist"] = [artist]
            audio["album"] = [album]
            if track.strip():
                audio["tracknumber"] = [track.strip()]
            audio.save(v2_version=3)
        except Exception as exc:
            self._log(f"Could not write EasyID3 tags for {mp3_path.name}: {exc}")

        artwork_path = row.get("artwork_path", "")
        if artwork_path and Path(artwork_path).exists():
            self._write_cover_art_from_image(Path(artwork_path), mp3_path)
            return
        source_art = Path(row.get("source_path", ""))
        if source_art.exists():
            self._copy_cover_art_between_files(source_art, mp3_path)

    def _write_cover_art_from_image(self, image_path: Path, target_mp3: Path):
        if not MUTAGEN_AVAILABLE or not image_path.exists():
            return
        try:
            try:
                dest_tags = ID3(str(target_mp3))
            except ID3NoHeaderError:
                dest_tags = ID3()
            dest_tags.delall("APIC")
            dest_tags.add(APIC(
                encoding=3,
                mime=guess_image_mime(image_path),
                type=3,
                desc="Cover",
                data=image_path.read_bytes(),
            ))
            dest_tags.save(str(target_mp3), v2_version=3)
        except Exception as exc:
            self._log(f"Could not write cover art to {target_mp3.name}: {exc}")

    def _copy_cover_art_between_files(self, source_mp3: Path, target_mp3: Path):
        if not MUTAGEN_AVAILABLE:
            return
        try:
            try:
                src_tags = ID3(str(source_mp3))
            except ID3NoHeaderError:
                return

            apic_frames = src_tags.getall("APIC")
            if not apic_frames:
                return

            try:
                dest_tags = ID3(str(target_mp3))
            except ID3NoHeaderError:
                dest_tags = ID3()

            for frame in dest_tags.getall("APIC"):
                dest_tags.delall("APIC")
                break
            for frame in apic_frames:
                dest_tags.add(APIC(
                    encoding=frame.encoding,
                    mime=frame.mime,
                    type=frame.type,
                    desc=frame.desc,
                    data=frame.data,
                ))
            dest_tags.save(str(target_mp3), v2_version=3)
        except Exception as exc:
            self._log(f"Could not copy cover art to {target_mp3.name}: {exc}")

    def _write_manifest(self, rows, spotify_folder: str):
        manifest = Path(spotify_folder) / "spotify_local_manifest.csv"
        with manifest.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["title", "artist", "album", "track", "playlist", "artwork_path", "source_path", "spotify_path"])
            for row in rows:
                writer.writerow([
                    row.get("title", ""),
                    row.get("artist", ""),
                    row.get("album", ""),
                    row.get("track", ""),
                    row.get("playlist", ""),
                    row.get("artwork_path", ""),
                    row.get("source_path", ""),
                    row.get("spotify_path", str(Path(spotify_folder) / self._safe_rel_path(row))),
                ])

    def _process_and_send_rows(self, rows, spotify_folder: str):
        Path(spotify_folder).mkdir(parents=True, exist_ok=True)
        processed = []
        for row in rows:
            source = row.get("source_path")
            if not source or not Path(source).exists():
                self._log(f"Skipping missing source file: {source}")
                continue
            destination = Path(spotify_folder) / self._safe_rel_path(row)
            self._copy_with_cleanup(source, destination)
            if self.write_tags_var.get():
                self._write_mp3_tags(destination, row)
            processed_row = dict(row)
            processed_row["spotify_path"] = str(destination)
            processed.append(processed_row)
            self._log(f"Sent to Spotify folder: {destination}")
        if processed:
            self.last_processed_rows = processed
            self._write_manifest(processed, spotify_folder)
            self._log(f"Wrote manifest: {Path(spotify_folder) / 'spotify_local_manifest.csv'}")

    def _delete_output_rows(self, rows):
        deleted_count = 0
        for row in rows:
            for key in ("spotify_path", "source_path"):
                file_path = row.get(key)
                if not file_path:
                    continue
                path = Path(file_path)
                try:
                    if path.exists():
                        path.unlink()
                        deleted_count += 1
                        self._log(f"Deleted file: {path}")
                except Exception as exc:
                    self._log(f"Could not delete {path}: {exc}")

        source_paths = {r.get("source_path", "") for r in rows}
        spotify_paths = {r.get("spotify_path", "") for r in rows}
        self.last_processed_rows = [
            row for row in self.last_processed_rows
            if row.get("source_path", "") not in source_paths
            and row.get("spotify_path", "") not in spotify_paths
        ]
        self.last_metadata_rows = [
            row for row in self.last_metadata_rows
            if row.get("source_path", "") not in source_paths
        ]

        spotify_folder = self.spotify_folder_var.get().strip()
        if spotify_folder and self.last_processed_rows:
            self._write_manifest(self.last_processed_rows, spotify_folder)
        elif spotify_folder:
            manifest = Path(spotify_folder) / "spotify_local_manifest.csv"
            try:
                if manifest.exists():
                    manifest.unlink()
                    self._log(f"Deleted manifest: {manifest}")
            except Exception as exc:
                self._log(f"Could not delete manifest: {exc}")

        messagebox.showinfo(APP_TITLE, f"Deleted {deleted_count} file(s).")

    def _send_last_download_to_spotify_folder(self):
        if not self.last_metadata_rows:
            messagebox.showinfo(APP_TITLE, "There is no recent download loaded yet.")
            return
        self._process_and_send_rows(self.last_metadata_rows, self.spotify_folder_var.get().strip())
        messagebox.showinfo(APP_TITLE, "Last download copied into the Spotify-ready folder.")

    def _export_build_script(self):
        target = filedialog.asksaveasfilename(
            title="Save Windows build script",
            defaultextension=".bat",
            filetypes=[("Batch files", "*.bat")],
            initialfile="build_spotify_pipeline_exe.bat",
        )
        if not target:
            return
        script_name = Path(sys.argv[0]).name
        script = f'''@echo off
setlocal
python -m pip install --upgrade pip
python -m pip install --upgrade yt-dlp ffmpeg-python mutagen pyinstaller
pyinstaller --noconfirm --onefile --windowed --name "SpotifyLocalFilesPipeline" "{script_name}"
echo.
echo Build finished. Check the dist folder for SpotifyLocalFilesPipeline.exe
pause
'''
        Path(target).write_text(script, encoding="utf-8")
        self._log(f"Exported build script: {target}")
        messagebox.showinfo(APP_TITLE, "Build script exported.")


def main():
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    style = ttk.Style()
    if sys.platform.startswith("win"):
        try:
            style.theme_use("vista")
        except Exception:
            pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
