import ctypes
import logging
import queue
import re
import subprocess
import threading
import time
import tkinter as tk
from dataclasses import dataclass, field
from itertools import combinations
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import cv2
import numpy as np
import win32gui
import win32con
import win32ui


LOG_PATH = Path.cwd() / "log.log"
logger = logging.getLogger("sudoku_scrcpy_solver")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    file_handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    )
    logger.addHandler(file_handler)
    logger.propagate = False


try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


@dataclass
class Rect:
    x: int
    y: int
    w: int
    h: int

    @property
    def x2(self) -> int:
        return self.x + self.w

    @property
    def y2(self) -> int:
        return self.y + self.h


@dataclass
class GridGeometry:
    rect: Rect
    x_lines: List[int]
    y_lines: List[int]


@dataclass
class ReadResult:
    board: List[List[int]]
    filled_cells: set
    templates: Dict[int, np.ndarray]
    geometry: GridGeometry
    keypad_rect: Rect
    content_image: np.ndarray
    missing_templates: Tuple[int, ...] = ()
    unknown_digit_cells: set = field(default_factory=set)


@dataclass
class StepMove:
    row: int
    col: int
    value: int
    reason: str


class ScrcpyManager:
    def __init__(self, status_callback: Callable[[str], None]):
        self.status_callback = status_callback
        self.proc: Optional[subprocess.Popen] = None
        self.reader_thread: Optional[threading.Thread] = None
        self.lock = threading.Lock()
        self.adb_serial: Optional[str] = None
        self.window_title: Optional[str] = None
        self.device_size: Optional[Tuple[int, int]] = None
        self.hwnd: Optional[int] = None
        self.ready = False
        self.running = False

    def start(self) -> None:
        with self.lock:
            if self.running:
                return
            self.running = True
        logger.info("scrcpy_start_requested")
        self.proc = subprocess.Popen(
            ["scrcpy"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="ignore",
            bufsize=1,
        )
        self.reader_thread = threading.Thread(target=self._read_output, daemon=True)
        self.reader_thread.start()
        self.status_callback("正在启动 scrcpy...")

    def stop(self) -> None:
        proc = self.proc
        if not proc:
            return
        if proc.poll() is None:
            logger.info("scrcpy_stop_requested")
            proc.terminate()

    def _read_output(self) -> None:
        if not self.proc or not self.proc.stdout:
            return
        for raw_line in self.proc.stdout:
            line = raw_line.strip()
            if not line:
                continue
            self._parse_line(line)
        with self.lock:
            self.running = False
            self.ready = False
        self.status_callback("scrcpy 已退出")

    def _parse_line(self, line: str) -> None:
        logger.info("scrcpy_output line=%s", line)
        device_match = re.search(r"INFO:\s+-->\s+\(\w+\)\s+(\S+)\s+\w+\s+(\S+)$", line)
        if device_match:
            self.adb_serial = device_match.group(1)
            self.window_title = device_match.group(2)
            self.status_callback(f"检测到设备 {self.window_title}")
            self._refresh_hwnd()
            return
        texture_match = re.search(r"INFO:\s+Texture:\s+(\d+)x(\d+)", line)
        if texture_match:
            self.device_size = (int(texture_match.group(1)), int(texture_match.group(2)))
            self._refresh_hwnd()
            if self.hwnd and self.device_size:
                self.ready = True
                self.status_callback(
                    f"scrcpy 已就绪，分辨率 {self.device_size[0]}x{self.device_size[1]}"
                )
            return
        if "Renderer:" in line:
            self.status_callback(line)

    def _refresh_hwnd(self) -> None:
        if not self.window_title:
            return
        title = self.window_title
        found: List[int] = []

        def enum_handler(hwnd: int, _: int) -> None:
            if not win32gui.IsWindowVisible(hwnd):
                return
            window_text = win32gui.GetWindowText(hwnd)
            if not window_text:
                return
            if window_text == title or title in window_text:
                found.append(hwnd)

        win32gui.EnumWindows(enum_handler, 0)
        self.hwnd = found[0] if found else None
        if self.hwnd and self.device_size:
            self.ready = True

    def ensure_ready(self) -> None:
        if not self.running:
            self.start()
        deadline = time.time() + 20
        while time.time() < deadline:
            self._refresh_hwnd()
            if self.ready and self.hwnd and self.device_size:
                return
            time.sleep(0.2)
        raise RuntimeError("scrcpy 启动超时，未找到设备窗口")

    def capture_content(self) -> np.ndarray:
        self.ensure_ready()
        if not self.hwnd or not self.device_size:
            raise RuntimeError("scrcpy 尚未准备好")
        if win32gui.IsIconic(self.hwnd):
            raise RuntimeError("scrcpy 窗口已最小化，无法截图")
        image = self._print_window(self.hwnd)
        if float(image.mean()) < 2.0:
            raise RuntimeError("离屏截图结果接近全黑，请确认 scrcpy 窗口正常显示")
        left, top, right, bottom = win32gui.GetWindowRect(self.hwnd)
        window_w = right - left
        window_h = bottom - top
        client_left_top = win32gui.ClientToScreen(self.hwnd, (0, 0))
        client_right_bottom = win32gui.ClientToScreen(
            self.hwnd, win32gui.GetClientRect(self.hwnd)[2:]
        )
        rel_x1 = max(0, client_left_top[0] - left)
        rel_y1 = max(0, client_left_top[1] - top)
        rel_x2 = min(window_w, client_right_bottom[0] - left)
        rel_y2 = min(window_h, client_right_bottom[1] - top)
        client = image[rel_y1:rel_y2, rel_x1:rel_x2]
        return self._crop_black_bars(client)

    def _crop_black_bars(self, image: np.ndarray) -> np.ndarray:
        if not self.device_size:
            return image
        h, w = image.shape[:2]
        if h == 0 or w == 0:
            return image
        target_ratio = self.device_size[0] / self.device_size[1]
        current_ratio = w / h
        if current_ratio > target_ratio:
            content_h = h
            content_w = int(round(content_h * target_ratio))
            x = max(0, (w - content_w) // 2)
            return image[:, x : x + content_w]
        content_w = w
        content_h = int(round(content_w / target_ratio))
        y = max(0, (h - content_h) // 2)
        return image[y : y + content_h, :]

    def _print_window(self, hwnd: int) -> np.ndarray:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top
        hwnd_dc = win32gui.GetWindowDC(hwnd)
        src_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        mem_dc = src_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(src_dc, width, height)
        mem_dc.SelectObject(bitmap)
        result = ctypes.windll.user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), 2)
        if result != 1:
            result = ctypes.windll.user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), 0)
        if result != 1:
            mem_dc.DeleteDC()
            src_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwnd_dc)
            win32gui.DeleteObject(bitmap.GetHandle())
            raise RuntimeError("窗口截图失败，请确认 scrcpy 可见且未最小化")
        bmp_info = bitmap.GetInfo()
        bmp_data = bitmap.GetBitmapBits(True)
        mem_dc.DeleteDC()
        src_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
        win32gui.DeleteObject(bitmap.GetHandle())
        image = np.frombuffer(bmp_data, dtype=np.uint8)
        image = image.reshape((bmp_info["bmHeight"], bmp_info["bmWidth"], 4))
        return cv2.cvtColor(image, cv2.COLOR_BGRA2BGR)

    def tap(self, x: int, y: int) -> None:
        self.ensure_ready()
        if self.hwnd and not win32gui.IsIconic(self.hwnd):
            try:
                client_x, client_y = self._device_to_client_point(int(x), int(y))
                lparam = ((client_y & 0xFFFF) << 16) | (client_x & 0xFFFF)
                logger.info(
                    "window_tap hwnd=%s device_x=%s device_y=%s client_x=%s client_y=%s",
                    self.hwnd,
                    int(x),
                    int(y),
                    client_x,
                    client_y,
                )
                win32gui.PostMessage(self.hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
                win32gui.PostMessage(self.hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                win32gui.PostMessage(self.hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                return
            except Exception:
                logger.exception("window_tap_failed hwnd=%s device_x=%s device_y=%s", self.hwnd, x, y)
        if not self.adb_serial:
            raise RuntimeError("未拿到 adb 设备序列号")
        logger.info("adb_tap_fallback serial=%s x=%s y=%s", self.adb_serial, x, y)
        cmd = ["adb", "-s", self.adb_serial, "shell", "input", "tap", str(int(x)), str(int(y))]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def _device_to_client_point(self, x: int, y: int) -> Tuple[int, int]:
        if not self.hwnd or not self.device_size:
            raise RuntimeError("scrcpy 尚未准备好")
        client_rect = win32gui.GetClientRect(self.hwnd)
        client_w = max(1, client_rect[2] - client_rect[0])
        client_h = max(1, client_rect[3] - client_rect[1])
        device_w, device_h = self.device_size
        target_ratio = device_w / device_h
        current_ratio = client_w / client_h
        if current_ratio > target_ratio:
            content_h = client_h
            content_w = int(round(content_h * target_ratio))
            offset_x = max(0, (client_w - content_w) // 2)
            offset_y = 0
        else:
            content_w = client_w
            content_h = int(round(content_w / target_ratio))
            offset_x = 0
            offset_y = max(0, (client_h - content_h) // 2)
        client_x = offset_x + int(round(x * content_w / device_w))
        client_y = offset_y + int(round(y * content_h / device_h))
        client_x = min(max(0, client_x), client_w - 1)
        client_y = min(max(0, client_y), client_h - 1)
        return client_x, client_y


class SudokuVision:
    GIVEN_RGB = np.array([156, 97, 47], dtype=np.float32)
    FILLED_RGB = np.array([203, 113, 24], dtype=np.float32)

    def __init__(self) -> None:
        self.template_cache: Dict[int, np.ndarray] = {}
        self.board_cache: Optional[List[List[int]]] = None

    def read(self, image_bgr: np.ndarray) -> ReadResult:
        previous_board = [row[:] for row in self.board_cache] if self.board_cache is not None else None
        geometry = self._locate_grid(image_bgr)
        keypad_rect = self._locate_keypad(image_bgr)
        templates, missing_templates = self._build_templates(image_bgr, keypad_rect)
        board, filled_cells, unknown_digit_cells = self._read_board(
            image_bgr,
            geometry,
            templates,
            previous_board=previous_board,
            allow_unknown=bool(missing_templates),
        )
        self.template_cache.update(templates)
        self.board_cache = [row[:] for row in board]
        if missing_templates:
            logger.info(
                "template_slots_missing values=%s unknown_cells=%s",
                list(missing_templates),
                sorted((row + 1, col + 1) for row, col in unknown_digit_cells),
            )
        return ReadResult(
            board=board,
            filled_cells=filled_cells,
            templates=templates,
            geometry=geometry,
            keypad_rect=keypad_rect,
            content_image=image_bgr,
            missing_templates=missing_templates,
            unknown_digit_cells=unknown_digit_cells,
        )

    def read_single_cell(
        self,
        image_bgr: np.ndarray,
        geometry: GridGeometry,
        templates: Dict[int, np.ndarray],
        row: int,
        col: int,
    ) -> Tuple[int, float, float]:
        x1, y1, x2, y2 = self._cell_rect(geometry, row, col)
        crop = image_bgr[y1:y2, x1:x2]
        digit, score, gap, _ = self._recognize_digit(crop, templates)
        return digit, score, gap

    def _locate_grid(self, image_bgr: np.ndarray) -> GridGeometry:
        hsv = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2HSV)
        mask = cv2.inRange(hsv, (8, 90, 100), (28, 255, 255))
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8), iterations=2)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            raise RuntimeError("未找到数独棋盘")
        h, w = image_bgr.shape[:2]
        candidates: List[Tuple[float, Rect]] = []
        for contour in contours:
            x, y, cw, ch = cv2.boundingRect(contour)
            if cw < w * 0.25 or ch < h * 0.15:
                continue
            ratio = cw / max(ch, 1)
            if not 0.85 <= ratio <= 1.15:
                continue
            candidates.append((cw * ch, Rect(x, y, cw, ch)))
        if not candidates:
            raise RuntimeError("未找到稳定的棋盘区域")
        rough = max(candidates, key=lambda item: item[0])[1]
        crop = image_bgr[rough.y : rough.y2, rough.x : rough.x2]
        hsv_crop = cv2.cvtColor(crop, cv2.COLOR_BGR2HSV)
        grid_mask = cv2.inRange(hsv_crop, (8, 90, 100), (28, 255, 255))
        x_lines = self._extract_lines(grid_mask, axis=0)
        y_lines = self._extract_lines(grid_mask, axis=1)
        if len(x_lines) < 10 or len(y_lines) < 10:
            x_lines = self._fallback_lines(rough.x, rough.w)
            y_lines = self._fallback_lines(rough.y, rough.h)
            return GridGeometry(rect=rough, x_lines=x_lines, y_lines=y_lines)
        x_lines = [rough.x + value for value in x_lines[:10]]
        y_lines = [rough.y + value for value in y_lines[:10]]
        rect = Rect(x_lines[0], y_lines[0], x_lines[-1] - x_lines[0], y_lines[-1] - y_lines[0])
        return GridGeometry(rect=rect, x_lines=x_lines, y_lines=y_lines)

    def _extract_lines(self, mask: np.ndarray, axis: int) -> List[int]:
        proj = mask.sum(axis=1 - axis).astype(np.float32)
        if proj.max() <= 0:
            return []
        threshold = max(255.0 * 2, proj.max() * 0.35)
        positions = np.where(proj >= threshold)[0]
        if len(positions) == 0:
            return []
        groups: List[List[int]] = [[int(positions[0])]]
        for pos in positions[1:]:
            if int(pos) - groups[-1][-1] <= 4:
                groups[-1].append(int(pos))
            else:
                groups.append([int(pos)])
        centers = [int(round((group[0] + group[-1]) / 2)) for group in groups]
        if len(centers) > 10:
            scores = [proj[group].max() for group in groups]
            chosen = sorted(
                sorted(range(len(centers)), key=lambda i: scores[i], reverse=True)[:10]
            )
            centers = [centers[i] for i in chosen]
        centers = sorted(centers)
        if len(centers) == 10:
            return centers
        if len(centers) > 1:
            first = centers[0]
            last = centers[-1]
            step = (last - first) / 9
            return [int(round(first + step * i)) for i in range(10)]
        return centers

    def _fallback_lines(self, start: int, length: int) -> List[int]:
        step = length / 9
        return [int(round(start + i * step)) for i in range(10)]

    def _locate_keypad(self, image_bgr: np.ndarray) -> Rect:
        h, w = image_bgr.shape[:2]
        lower = image_bgr[int(h * 0.52) :, :]
        hsv = cv2.cvtColor(lower, cv2.COLOR_BGR2HSV)
        active = cv2.inRange(hsv, (5, 80, 60), (25, 255, 190))
        disabled = cv2.inRange(hsv, (10, 10, 110), (30, 90, 220))
        mask = cv2.bitwise_or(active, disabled)
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((17, 17), np.uint8), iterations=2)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        candidates: List[Tuple[int, Rect]] = []
        for contour in contours:
            x, y, cw, ch = cv2.boundingRect(contour)
            if cw < w * 0.45 or ch < h * 0.06:
                continue
            candidates.append((cw * ch, Rect(x, y + int(h * 0.52), cw, ch)))
        if candidates:
            rect = max(candidates, key=lambda item: item[0])[1]
            pad_x = int(rect.w * 0.02)
            pad_y = int(rect.h * 0.12)
            left = max(0, rect.x - pad_x)
            top = max(0, rect.y - pad_y)
            return Rect(
                left,
                top,
                min(w - left, rect.w + pad_x * 2),
                min(h - top, rect.h + pad_y * 2),
            )
        return Rect(int(w * 0.05), int(h * 0.79), int(w * 0.90), int(h * 0.14))

    def _build_templates(self, image_bgr: np.ndarray, keypad_rect: Rect) -> Tuple[Dict[int, np.ndarray], Tuple[int, ...]]:
        templates: Dict[int, np.ndarray] = {}
        missing_templates: List[int] = []
        slot_w = keypad_rect.w / 9
        for value in range(1, 10):
            x1 = int(round(keypad_rect.x + (value - 1) * slot_w))
            x2 = int(round(keypad_rect.x + value * slot_w))
            crop = image_bgr[keypad_rect.y : keypad_rect.y2, x1:x2]
            candidates = self._extract_digit_candidates(crop)
            if not candidates:
                cached = self.template_cache.get(value)
                if cached is not None:
                    templates[value] = cached
                    continue
                missing_templates.append(value)
                continue
            mask = max(candidates, key=lambda item: int((item > 0).sum()))
            templates[value] = self._normalize_mask(mask)
        if not templates:
            raise RuntimeError("未能提取任何底部数字模板")
        return templates, tuple(missing_templates)

    def _read_board(
        self,
        image_bgr: np.ndarray,
        geometry: GridGeometry,
        templates: Dict[int, np.ndarray],
        previous_board: Optional[List[List[int]]] = None,
        allow_unknown: bool = False,
    ) -> Tuple[List[List[int]], set, set]:
        board = [[0 for _ in range(9)] for _ in range(9)]
        filled_cells = set()
        unknown_digit_cells = set()
        for row in range(9):
            for col in range(9):
                x1, y1, x2, y2 = self._cell_rect(geometry, row, col)
                crop = image_bgr[y1:y2, x1:x2]
                digit, score, gap, mask = self._recognize_digit(crop, templates)
                if mask is None:
                    continue
                if digit == 0:
                    if previous_board is not None and previous_board[row][col] != 0:
                        board[row][col] = previous_board[row][col]
                        if self._is_filled_digit(crop, mask):
                            filled_cells.add((row, col))
                        continue
                    if allow_unknown:
                        unknown_digit_cells.add((row, col))
                        continue
                    raise RuntimeError(
                        f"第 {row + 1} 行第 {col + 1} 列识别不稳定，最佳分数 {score:.3f}，次优差值 {gap:.3f}"
                    )
                board[row][col] = digit
                if self._is_filled_digit(crop, mask):
                    filled_cells.add((row, col))
        return board, filled_cells, unknown_digit_cells

    def _is_filled_digit(self, crop_bgr: np.ndarray, mask: np.ndarray) -> bool:
        pixels = crop_bgr[mask > 0]
        if len(pixels) == 0:
            return False
        mean_rgb = pixels[:, ::-1].mean(axis=0).astype(np.float32)
        if mean_rgb.mean() > 220:
            return False
        orange_distance = np.linalg.norm(mean_rgb - self.FILLED_RGB)
        brown_distance = np.linalg.norm(mean_rgb - self.GIVEN_RGB)
        return orange_distance + 10 < brown_distance

    def _recognize_digit(
        self, crop_bgr: np.ndarray, templates: Dict[int, np.ndarray]
    ) -> Tuple[int, float, float, Optional[np.ndarray]]:
        best_digit = 0
        best_score = 0.0
        best_gap = 0.0
        best_mask: Optional[np.ndarray] = None
        for mask in self._extract_digit_candidates(crop_bgr):
            digit, score, gap = self._match_digit(mask, templates)
            if score > best_score or (score == best_score and gap > best_gap):
                best_digit = digit
                best_score = score
                best_gap = gap
                best_mask = mask
        return best_digit, best_score, best_gap, best_mask

    def _cell_rect(self, geometry: GridGeometry, row: int, col: int) -> Tuple[int, int, int, int]:
        left = geometry.x_lines[col]
        right = geometry.x_lines[col + 1]
        top = geometry.y_lines[row]
        bottom = geometry.y_lines[row + 1]
        pad_x = max(2, int((right - left) * 0.12))
        pad_y = max(2, int((bottom - top) * 0.12))
        return left + pad_x, top + pad_y, right - pad_x, bottom - pad_y

    def _extract_digit_candidates(self, crop_bgr: np.ndarray) -> List[np.ndarray]:
        if crop_bgr.size == 0:
            return []
        h, w = crop_bgr.shape[:2]
        border = np.concatenate(
            [
                crop_bgr[: max(1, h // 8), :, :].reshape(-1, 3),
                crop_bgr[-max(1, h // 8) :, :, :].reshape(-1, 3),
                crop_bgr[:, : max(1, w // 8), :].reshape(-1, 3),
                crop_bgr[:, -max(1, w // 8) :, :].reshape(-1, 3),
            ],
            axis=0,
        )
        bg = np.median(border, axis=0)
        bg_gray = float(0.114 * bg[0] + 0.587 * bg[1] + 0.299 * bg[2])
        gray = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2GRAY)
        hsv = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2HSV)
        diff = np.linalg.norm(crop_bgr.astype(np.float32) - bg.astype(np.float32), axis=2)
        threshold = max(24.0, float(np.percentile(diff, 84)))
        diff_mask = (diff >= threshold).astype(np.uint8) * 255
        dark_mask = (gray <= max(0.0, bg_gray - 24.0)).astype(np.uint8) * 255
        light_mask = (gray >= min(255.0, bg_gray + 28.0)).astype(np.uint8) * 255
        white_mask = (
            (hsv[:, :, 2] >= max(170.0, bg_gray + 24.0)) & (hsv[:, :, 1] <= 130)
        ).astype(np.uint8) * 255
        blur = cv2.GaussianBlur(gray, (3, 3), 0)
        adaptive_dark = cv2.adaptiveThreshold(
            blur,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            11,
            2,
        )
        adaptive_light = cv2.adaptiveThreshold(
            blur,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            11,
            2,
        )
        raw_masks: List[np.ndarray] = [diff_mask]
        if bg_gray >= 170:
            raw_masks.extend([dark_mask, adaptive_dark])
        else:
            raw_masks.extend([light_mask, white_mask, adaptive_light])
        masks: List[np.ndarray] = []
        for raw_mask in raw_masks:
            mask = cv2.morphologyEx(raw_mask, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8), iterations=1)
            mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8), iterations=1)
            masks.extend(self._component_masks(mask))
        return self._dedupe_masks(masks)

    def _component_masks(self, mask: np.ndarray) -> List[np.ndarray]:
        h, w = mask.shape[:2]
        num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(mask, 8)
        if num_labels <= 1:
            return []
        center = np.array([w / 2, h / 2], dtype=np.float32)
        components: List[Tuple[float, np.ndarray]] = []
        for label in range(1, num_labels):
            x, y, cw, ch, area = stats[label]
            if area < (w * h) * 0.015 or area > (w * h) * 0.55:
                continue
            if cw < w * 0.06 or ch < h * 0.18:
                continue
            cx = x + cw / 2
            cy = y + ch / 2
            dist = np.linalg.norm(np.array([cx, cy], dtype=np.float32) - center)
            ratio_penalty = abs((cw / max(ch, 1)) - 0.55) * 18.0
            score = area - dist * 0.7 - ratio_penalty
            components.append((score, np.where(labels == label, 255, 0).astype(np.uint8)))
        components.sort(key=lambda item: item[0], reverse=True)
        return [component for _, component in components[:4]]

    def _dedupe_masks(self, masks: List[np.ndarray]) -> List[np.ndarray]:
        unique: List[np.ndarray] = []
        for mask in masks:
            mask_bool = mask > 0
            duplicate = False
            for existing in unique:
                existing_bool = existing > 0
                union = np.logical_or(mask_bool, existing_bool).sum()
                if union == 0:
                    continue
                overlap = np.logical_and(mask_bool, existing_bool).sum() / union
                if overlap > 0.88:
                    duplicate = True
                    break
            if not duplicate:
                unique.append(mask)
        return unique

    def _normalize_mask(self, mask: np.ndarray, size: Tuple[int, int] = (52, 52)) -> np.ndarray:
        ys, xs = np.where(mask > 0)
        if len(xs) == 0:
            return np.zeros(size, dtype=np.uint8)
        crop = mask[ys.min() : ys.max() + 1, xs.min() : xs.max() + 1]
        target_h, target_w = size
        scale = min((target_w - 8) / max(1, crop.shape[1]), (target_h - 8) / max(1, crop.shape[0]))
        resized = cv2.resize(
            crop,
            (max(1, int(round(crop.shape[1] * scale))), max(1, int(round(crop.shape[0] * scale)))),
            interpolation=cv2.INTER_NEAREST,
        )
        canvas = np.zeros((target_h, target_w), dtype=np.uint8)
        y = (target_h - resized.shape[0]) // 2
        x = (target_w - resized.shape[1]) // 2
        canvas[y : y + resized.shape[0], x : x + resized.shape[1]] = resized
        return canvas

    def _match_digit(self, mask: np.ndarray, templates: Dict[int, np.ndarray]) -> Tuple[int, float, float]:
        normalized = self._normalize_mask(mask)
        scored: List[Tuple[float, int]] = []
        for value, template in templates.items():
            intersection = np.logical_and(normalized > 0, template > 0).sum()
            union = np.logical_or(normalized > 0, template > 0).sum()
            if union == 0:
                continue
            area_diff = abs((normalized > 0).sum() - (template > 0).sum()) / max(
                1, (template > 0).sum()
            )
            score = (intersection / union) - area_diff * 0.12
            scored.append((float(score), value))
        scored.sort(reverse=True)
        if not scored:
            return 0, 0.0, 0.0
        best_score, best_value = scored[0]
        second_score = scored[1][0] if len(scored) > 1 else -1.0
        gap = best_score - second_score
        if best_score < 0.34 or gap < 0.03:
            return 0, best_score, gap
        return best_value, best_score, gap


class SudokuSolver:
    def validate(self, board: List[List[int]]) -> None:
        for index in range(9):
            self._validate_unit(board[index], f"第 {index + 1} 行")
            self._validate_unit([board[row][index] for row in range(9)], f"第 {index + 1} 列")
        for box_row in range(3):
            for box_col in range(3):
                values = []
                for row in range(box_row * 3, box_row * 3 + 3):
                    for col in range(box_col * 3, box_col * 3 + 3):
                        values.append(board[row][col])
                self._validate_unit(values, f"宫 {box_row + 1}-{box_col + 1}")

    def _validate_unit(self, values: List[int], name: str) -> None:
        existing = [value for value in values if value != 0]
        if len(existing) != len(set(existing)):
            raise RuntimeError(f"{name} 出现重复数字，当前画面可能有错误输入")

    def next_step(self, board: List[List[int]]) -> Optional[StepMove]:
        candidates = self._build_candidates(board)
        if not candidates:
            return None
        move = self._direct_move(board, candidates)
        if move:
            return move
        working = {cell: values.copy() for cell, values in candidates.items()}
        reasons: List[str] = []
        reducers = (
            self._reduce_naked_pairs,
            self._reduce_naked_triples,
            self._reduce_hidden_pairs,
            self._reduce_pointing_locked_candidates,
            self._reduce_claiming_locked_candidates,
            self._reduce_x_wing,
            self._reduce_xy_wing,
        )
        for _ in range(64):
            progress = False
            for reducer in reducers:
                reason = reducer(working)
                if not reason:
                    continue
                reasons.append(reason)
                self._validate_candidates(working)
                move = self._direct_move(board, working)
                if move:
                    move.reason = " -> ".join(reasons + [move.reason])
                    return move
                progress = True
                break
            if not progress:
                break
        return None

    def _build_candidates(self, board: List[List[int]]) -> Dict[Tuple[int, int], set]:
        return {
            (row, col): self._candidates(board, row, col)
            for row in range(9)
            for col in range(9)
            if board[row][col] == 0
        }

    def _direct_move(
        self,
        board: List[List[int]],
        candidates: Dict[Tuple[int, int], set],
    ) -> Optional[StepMove]:
        for row in range(9):
            for col in range(9):
                values = candidates.get((row, col))
                if values and len(values) == 1:
                    return StepMove(row=row, col=col, value=next(iter(values)), reason="唯一候选")
        for row in range(9):
            move = self._hidden_single(
                [(row, col) for col in range(9) if board[row][col] == 0], candidates, "行唯一"
            )
            if move:
                return move
        for col in range(9):
            move = self._hidden_single(
                [(row, col) for row in range(9) if board[row][col] == 0], candidates, "列唯一"
            )
            if move:
                return move
        for box_row in range(3):
            for box_col in range(3):
                cells = [
                    (row, col)
                    for row in range(box_row * 3, box_row * 3 + 3)
                    for col in range(box_col * 3, box_col * 3 + 3)
                    if board[row][col] == 0
                ]
                move = self._hidden_single(cells, candidates, "宫唯一")
                if move:
                    return move
        return None

    def _reduce_naked_pairs(self, candidates: Dict[Tuple[int, int], set]) -> Optional[str]:
        for label, cells in self._unit_groups(candidates):
            pairs: Dict[Tuple[int, int], List[Tuple[int, int]]] = {}
            for cell in cells:
                values = candidates[cell]
                if len(values) == 2:
                    pairs.setdefault(tuple(sorted(values)), []).append(cell)
            for pair, pair_cells in pairs.items():
                if len(pair_cells) != 2:
                    continue
                changed_cells = []
                pair_values = set(pair)
                for cell in cells:
                    if cell in pair_cells:
                        continue
                    before = len(candidates[cell])
                    candidates[cell] -= pair_values
                    if len(candidates[cell]) != before:
                        changed_cells.append(cell)
                if changed_cells:
                    return f"裸对({label} 的 {pair[0]}/{pair[1]})"
        return None

    def _reduce_naked_triples(self, candidates: Dict[Tuple[int, int], set]) -> Optional[str]:
        for label, cells in self._unit_groups(candidates):
            triple_cells = [cell for cell in cells if 2 <= len(candidates[cell]) <= 3]
            if len(triple_cells) < 3:
                continue
            for combo in combinations(triple_cells, 3):
                union = set().union(*(candidates[cell] for cell in combo))
                if len(union) != 3:
                    continue
                covered_cells = [cell for cell in cells if candidates[cell].issubset(union)]
                if len(covered_cells) != 3:
                    continue
                changed = False
                for cell in cells:
                    if cell in covered_cells:
                        continue
                    before = len(candidates[cell])
                    candidates[cell] -= union
                    if len(candidates[cell]) != before:
                        changed = True
                if changed:
                    digits = "/".join(str(value) for value in sorted(union))
                    return f"裸三元组({label} 的 {digits})"
        return None

    def _reduce_hidden_pairs(self, candidates: Dict[Tuple[int, int], set]) -> Optional[str]:
        for label, cells in self._unit_groups(candidates):
            positions: Dict[int, set] = {}
            for value in range(1, 10):
                value_cells = {cell for cell in cells if value in candidates[cell]}
                if len(value_cells) == 2:
                    positions[value] = value_cells
            for first, second in combinations(sorted(positions), 2):
                target_cells = positions[first] | positions[second]
                if len(target_cells) != 2:
                    continue
                pair_values = {first, second}
                changed = False
                for cell in target_cells:
                    if candidates[cell] != pair_values:
                        candidates[cell] &= pair_values
                        changed = True
                if changed:
                    return f"隐藏对({label} 的 {first}/{second})"
        return None

    def _reduce_pointing_locked_candidates(self, candidates: Dict[Tuple[int, int], set]) -> Optional[str]:
        for box_row in range(3):
            for box_col in range(3):
                cells = [
                    (row, col)
                    for row in range(box_row * 3, box_row * 3 + 3)
                    for col in range(box_col * 3, box_col * 3 + 3)
                    if (row, col) in candidates
                ]
                for value in range(1, 10):
                    value_cells = [cell for cell in cells if value in candidates[cell]]
                    if len(value_cells) < 2:
                        continue
                    rows = {row for row, _ in value_cells}
                    if len(rows) == 1:
                        row = next(iter(rows))
                        changed = False
                        for col in range(9):
                            cell = (row, col)
                            if cell in candidates and cell not in value_cells and value in candidates[cell]:
                                candidates[cell].discard(value)
                                changed = True
                        if changed:
                            return f"宫区块摒除(宫 {box_row + 1}-{box_col + 1} 的 {value} 锁定在第 {row + 1} 行)"
                    cols = {col for _, col in value_cells}
                    if len(cols) == 1:
                        col = next(iter(cols))
                        changed = False
                        for row in range(9):
                            cell = (row, col)
                            if cell in candidates and cell not in value_cells and value in candidates[cell]:
                                candidates[cell].discard(value)
                                changed = True
                        if changed:
                            return f"宫区块摒除(宫 {box_row + 1}-{box_col + 1} 的 {value} 锁定在第 {col + 1} 列)"
        return None

    def _reduce_claiming_locked_candidates(self, candidates: Dict[Tuple[int, int], set]) -> Optional[str]:
        for row in range(9):
            cells = [(row, col) for col in range(9) if (row, col) in candidates]
            for value in range(1, 10):
                value_cells = [cell for cell in cells if value in candidates[cell]]
                boxes = {(cell_row // 3, cell_col // 3) for cell_row, cell_col in value_cells}
                if len(value_cells) < 2 or len(boxes) != 1:
                    continue
                box_row, box_col = next(iter(boxes))
                changed = False
                for cell_row in range(box_row * 3, box_row * 3 + 3):
                    for cell_col in range(box_col * 3, box_col * 3 + 3):
                        cell = (cell_row, cell_col)
                        if cell in candidates and cell not in value_cells and value in candidates[cell]:
                            candidates[cell].discard(value)
                            changed = True
                if changed:
                    return f"区块摒除(第 {row + 1} 行的 {value} 锁定在宫 {box_row + 1}-{box_col + 1})"
        for col in range(9):
            cells = [(row, col) for row in range(9) if (row, col) in candidates]
            for value in range(1, 10):
                value_cells = [cell for cell in cells if value in candidates[cell]]
                boxes = {(cell_row // 3, cell_col // 3) for cell_row, cell_col in value_cells}
                if len(value_cells) < 2 or len(boxes) != 1:
                    continue
                box_row, box_col = next(iter(boxes))
                changed = False
                for cell_row in range(box_row * 3, box_row * 3 + 3):
                    for cell_col in range(box_col * 3, box_col * 3 + 3):
                        cell = (cell_row, cell_col)
                        if cell in candidates and cell not in value_cells and value in candidates[cell]:
                            candidates[cell].discard(value)
                            changed = True
                if changed:
                    return f"区块摒除(第 {col + 1} 列的 {value} 锁定在宫 {box_row + 1}-{box_col + 1})"
        return None

    def _reduce_x_wing(self, candidates: Dict[Tuple[int, int], set]) -> Optional[str]:
        for value in range(1, 10):
            row_patterns: Dict[Tuple[int, int], List[int]] = {}
            for row in range(9):
                cols = tuple(col for col in range(9) if (row, col) in candidates and value in candidates[(row, col)])
                if len(cols) == 2:
                    row_patterns.setdefault(cols, []).append(row)
            for cols, rows in row_patterns.items():
                if len(rows) < 2:
                    continue
                for row_a, row_b in combinations(rows, 2):
                    changed = False
                    for row in range(9):
                        if row in (row_a, row_b):
                            continue
                        for col in cols:
                            cell = (row, col)
                            if cell in candidates and value in candidates[cell]:
                                candidates[cell].discard(value)
                                changed = True
                    if changed:
                        return f"X-Wing(数字 {value}，第 {row_a + 1}/{row_b + 1} 行锁定第 {cols[0] + 1}/{cols[1] + 1} 列)"
            col_patterns: Dict[Tuple[int, int], List[int]] = {}
            for col in range(9):
                rows = tuple(row for row in range(9) if (row, col) in candidates and value in candidates[(row, col)])
                if len(rows) == 2:
                    col_patterns.setdefault(rows, []).append(col)
            for rows, cols in col_patterns.items():
                if len(cols) < 2:
                    continue
                for col_a, col_b in combinations(cols, 2):
                    changed = False
                    for col in range(9):
                        if col in (col_a, col_b):
                            continue
                        for row in rows:
                            cell = (row, col)
                            if cell in candidates and value in candidates[cell]:
                                candidates[cell].discard(value)
                                changed = True
                    if changed:
                        return f"X-Wing(数字 {value}，第 {col_a + 1}/{col_b + 1} 列锁定第 {rows[0] + 1}/{rows[1] + 1} 行)"
        return None

    def _reduce_xy_wing(self, candidates: Dict[Tuple[int, int], set]) -> Optional[str]:
        bivalue_cells = [(cell, values) for cell, values in candidates.items() if len(values) == 2]
        for pivot, pivot_values in bivalue_cells:
            for wing_a, wing_a_values in bivalue_cells:
                if wing_a == pivot or not self._cells_share_unit(pivot, wing_a):
                    continue
                shared_a = pivot_values & wing_a_values
                if len(shared_a) != 1:
                    continue
                shared_a_value = next(iter(shared_a))
                pivot_other = pivot_values - {shared_a_value}
                wing_a_other = wing_a_values - {shared_a_value}
                if len(pivot_other) != 1 or len(wing_a_other) != 1:
                    continue
                pivot_other_value = next(iter(pivot_other))
                target_value = next(iter(wing_a_other))
                if pivot_other_value == target_value:
                    continue
                for wing_b, wing_b_values in bivalue_cells:
                    if wing_b in (pivot, wing_a) or not self._cells_share_unit(pivot, wing_b):
                        continue
                    if (pivot_values & wing_b_values) != {pivot_other_value}:
                        continue
                    if (wing_b_values - {pivot_other_value}) != {target_value}:
                        continue
                    changed = False
                    for cell, cell_values in candidates.items():
                        if cell in (pivot, wing_a, wing_b):
                            continue
                        if (
                            self._cells_share_unit(cell, wing_a)
                            and self._cells_share_unit(cell, wing_b)
                            and target_value in cell_values
                        ):
                            cell_values.discard(target_value)
                            changed = True
                    if changed:
                        return (
                            "XY-Wing("
                            f"枢纽 r{pivot[0] + 1}c{pivot[1] + 1}，"
                            f"翼 r{wing_a[0] + 1}c{wing_a[1] + 1} / r{wing_b[0] + 1}c{wing_b[1] + 1}，"
                            f"删除 {target_value})"
                        )
        return None

    def _cells_share_unit(self, first: Tuple[int, int], second: Tuple[int, int]) -> bool:
        return (
            first[0] == second[0]
            or first[1] == second[1]
            or (first[0] // 3, first[1] // 3) == (second[0] // 3, second[1] // 3)
        )

    def _unit_groups(
        self,
        candidates: Dict[Tuple[int, int], set],
    ) -> List[Tuple[str, List[Tuple[int, int]]]]:
        groups: List[Tuple[str, List[Tuple[int, int]]]] = []
        for row in range(9):
            groups.append((f"第 {row + 1} 行", [(row, col) for col in range(9) if (row, col) in candidates]))
        for col in range(9):
            groups.append((f"第 {col + 1} 列", [(row, col) for row in range(9) if (row, col) in candidates]))
        for box_row in range(3):
            for box_col in range(3):
                groups.append(
                    (
                        f"宫 {box_row + 1}-{box_col + 1}",
                        [
                            (row, col)
                            for row in range(box_row * 3, box_row * 3 + 3)
                            for col in range(box_col * 3, box_col * 3 + 3)
                            if (row, col) in candidates
                        ],
                    )
                )
        return groups

    def _validate_candidates(self, candidates: Dict[Tuple[int, int], set]) -> None:
        for (row, col), values in candidates.items():
            if not values:
                raise RuntimeError(f"第 {row + 1} 行第 {col + 1} 列没有候选数，当前画面可能有错误输入")

    def _hidden_single(
        self,
        cells: List[Tuple[int, int]],
        candidates: Dict[Tuple[int, int], set],
        reason: str,
    ) -> Optional[StepMove]:
        positions: Dict[int, List[Tuple[int, int]]] = {}
        for cell in cells:
            for value in candidates.get(cell, set()):
                positions.setdefault(value, []).append(cell)
        for value, value_cells in positions.items():
            if len(value_cells) == 1:
                row, col = value_cells[0]
                return StepMove(row=row, col=col, value=value, reason=reason)
        return None

    def _candidates(self, board: List[List[int]], row: int, col: int) -> set:
        used = set(board[row])
        used.update(board[r][col] for r in range(9))
        box_row = (row // 3) * 3
        box_col = (col // 3) * 3
        for r in range(box_row, box_row + 3):
            for c in range(box_col, box_col + 3):
                used.add(board[r][c])
        return {value for value in range(1, 10) if value not in used}


class SudokuApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("数独助手")
        self.root.geometry("860x620")
        self.root.configure(bg="#F7F1E6")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.scrcpy = ScrcpyManager(self.push_status)
        self.vision = SudokuVision()
        self.solver = SudokuSolver()
        self.status_queue: "queue.Queue[str]" = queue.Queue()
        self.board = [[0 for _ in range(9)] for _ in range(9)]
        self.given_cells = set()
        self.filled_cells = set()
        self.last_read: Optional[ReadResult] = None
        self.current_step: Optional[StepMove] = None
        self.auto_running = False
        self.reading = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="准备启动 scrcpy")
        self.display_info_var = tk.StringVar(value="等待读取画面")
        self._build_ui()
        self.scrcpy.start()
        self.root.after(150, self._drain_status)
        self.root.after(500, self._tick_read)
        self._log_event("app_initialized")
        self._log_event("app_initialized")

    def _build_ui(self) -> None:
        container = tk.Frame(self.root, bg="#F7F1E6")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        left = tk.Frame(container, width=220, bg="#EFE2C7", relief="ridge", bd=2)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        right = tk.Frame(container, bg="#F7F1E6")
        right.pack(side="right", fill="both", expand=True, padx=(16, 0))

        title = tk.Label(
            left,
            text="操作区",
            font=("Microsoft YaHei UI", 15, "bold"),
            bg="#EFE2C7",
            fg="#7A4A22",
        )
        title.pack(anchor="w", padx=14, pady=(14, 10))

        self.read_check = tk.Checkbutton(
            left,
            text="读取",
            variable=self.reading,
            command=self._toggle_reading,
            font=("Microsoft YaHei UI", 12),
            bg="#EFE2C7",
            activebackground="#EFE2C7",
            fg="#7A4A22",
            selectcolor="#FFF8EC",
        )
        self.read_check.pack(anchor="w", padx=14, pady=8)

        self.step_button = tk.Button(
            left,
            text="下一步",
            command=self.on_next_step,
            font=("Microsoft YaHei UI", 12),
            bg="#D78826",
            fg="#FFFFFF",
            activebackground="#C97815",
            relief="flat",
            width=14,
            pady=8,
        )
        self.step_button.pack(anchor="w", padx=14, pady=(10, 8))

        self.apply_button = tk.Button(
            left,
            text="操作",
            command=self.on_apply_step,
            font=("Microsoft YaHei UI", 12),
            bg="#A1602B",
            fg="#FFFFFF",
            activebackground="#8F5322",
            relief="flat",
            width=14,
            pady=8,
        )
        self.apply_button.pack(anchor="w", padx=14, pady=8)

        self.auto_button = tk.Button(
            left,
            text="自动操作",
            command=self.on_auto_apply,
            font=("Microsoft YaHei UI", 12),
            bg="#70A210",
            fg="#FFFFFF",
            activebackground="#5E8B0E",
            relief="flat",
            width=14,
            pady=8,
        )
        self.auto_button.pack(anchor="w", padx=14, pady=8)

        status_title = tk.Label(
            left,
            text="状态",
            font=("Microsoft YaHei UI", 12, "bold"),
            bg="#EFE2C7",
            fg="#7A4A22",
        )
        status_title.pack(anchor="w", padx=14, pady=(18, 6))

        status_label = tk.Label(
            left,
            textvariable=self.status_var,
            wraplength=180,
            justify="left",
            font=("Microsoft YaHei UI", 10),
            bg="#EFE2C7",
            fg="#6B512F",
        )
        status_label.pack(anchor="w", padx=14)

        self.canvas = tk.Canvas(
            right,
            width=520,
            height=520,
            bg="#FFFBE1",
            highlightthickness=0,
        )
        self.canvas.pack(side="top", anchor="center")

        self.info_label = tk.Label(
            right,
            textvariable=self.display_info_var,
            wraplength=520,
            justify="left",
            anchor="w",
            font=("Microsoft YaHei UI", 10),
            bg="#F7F1E6",
            fg="#6B512F",
        )
        self.info_label.pack(side="top", fill="x", pady=(10, 0))

        self._draw_board()

    def push_status(self, text: str) -> None:
        self.status_queue.put(text)
        logger.info("status text=%s", text)

    def _board_to_log_text(self, board: Optional[List[List[int]]] = None) -> str:
        active_board = board if board is not None else self.board
        return "/".join("".join(str(value) for value in row) for row in active_board)

    def _log_event(self, event: str, **kwargs: object) -> None:
        parts = [event]
        for key, value in kwargs.items():
            parts.append(f"{key}={value}")
        parts.append(f"board={self._board_to_log_text()}")
        parts.append(f"filled={sorted(self.filled_cells)}")
        if self.current_step:
            parts.append(
                f"step=r{self.current_step.row + 1}c{self.current_step.col + 1}={self.current_step.value}:{self.current_step.reason}"
            )
        logger.info(" ".join(parts))

    def _drain_status(self) -> None:
        latest = None
        while not self.status_queue.empty():
            latest = self.status_queue.get_nowait()
        if latest:
            self.status_var.set(latest)
        self.root.after(150, self._drain_status)

    def _toggle_reading(self) -> None:
        self._log_event("user_toggle_reading", enabled=self.reading.get())
        if self.reading.get():
            self.read_screen()
        else:
            self.push_status("已停止自动读取")

    def _tick_read(self) -> None:
        if self.reading.get() and not self.auto_running:
            self.read_screen(auto=True)
        self.root.after(900, self._tick_read)

    def read_screen(self, auto: bool = False) -> None:
        self._log_event("read_screen_started", auto=auto)
        try:
            previous_board = [row[:] for row in self.board]
            previous_step = self.current_step
            result = self._capture_read_result()
            board_changed = result.board != previous_board
            self._sync_cell_tracking(result.board, previous_board, result.filled_cells)
            self.board = [row[:] for row in result.board]
            self.last_read = result
            keep_step = (
                previous_step is not None
                and not board_changed
                and self.board[previous_step.row][previous_step.col] == 0
            )
            self.current_step = previous_step if keep_step else None
            self._draw_board()
            info_message: Optional[str] = None
            if result.unknown_digit_cells:
                missing_text = "、".join(str(value) for value in result.missing_templates) or "部分"
                info_message = (
                    f"底部数字 {missing_text} 已消失，本次先完成同步；仍有 "
                    f"{len(result.unknown_digit_cells)} 个已填格未能可靠识别，暂不继续求解。"
                )
            elif result.missing_templates:
                missing_text = "、".join(str(value) for value in result.missing_templates)
                info_message = f"底部数字 {missing_text} 已消失，已复用历史模板继续识别。"
            self._update_display_info(info_message)
            self._log_event(
                "read_screen_completed",
                auto=auto,
                board_changed=board_changed,
                result_board=self._board_to_log_text(result.board),
                missing_templates=list(result.missing_templates),
                unknown_digit_cells=sorted((row + 1, col + 1) for row, col in result.unknown_digit_cells),
            )
            if auto:
                if result.unknown_digit_cells:
                    self.push_status("已同步画面；有消失数字且仍缺少可靠模板，暂不继续求解")
                elif board_changed:
                    self.push_status("已自动同步最新画面")
            else:
                if result.unknown_digit_cells:
                    self.push_status("读取完成，但当前有消失数字未完全识别，程序不会报错也不会贸然求解")
                else:
                    self.push_status("读取完成")
        except Exception as exc:
            logger.exception("read_screen_failed auto=%s board=%s", auto, self._board_to_log_text())
            if not auto:
                self.push_status(str(exc))

    def on_next_step(self) -> None:
        self._log_event("user_next_step_clicked")
        self._compute_next_step(push_status=True)

    def _compute_next_step(self, push_status: bool) -> bool:
        if not self.last_read:
            self.read_screen()
        if not self.last_read:
            return False
        if self.last_read.unknown_digit_cells:
            self.current_step = None
            self._draw_board()
            self._update_display_info(
                f"当前仍有 {len(self.last_read.unknown_digit_cells)} 个已填格因底部数字消失而未完全确认，先不同步下一步。"
            )
            self._log_event(
                "compute_next_step_blocked_missing_templates",
                missing_templates=list(self.last_read.missing_templates),
                unknown_digit_cells=sorted((row + 1, col + 1) for row, col in self.last_read.unknown_digit_cells),
                push_status=push_status,
            )
            if push_status:
                self.push_status("当前有消失数字尚未可靠识别，暂不计算下一步")
            return False
        try:
            self.solver.validate(self.board)
            move = self.solver.next_step(self.board)
            if not move:
                self.current_step = None
                self._draw_board()
                self._update_display_info()
                self._log_event("compute_next_step_none", push_status=push_status)
                if push_status:
                    self.push_status("当前局面没有找到可直接确定的一步")
                return False
            self.current_step = move
            self._draw_board()
            self._update_display_info()
            self._log_event(
                "compute_next_step_success",
                push_status=push_status,
                row=move.row + 1,
                col=move.col + 1,
                value=move.value,
                reason=move.reason,
            )
            if push_status:
                self.push_status(
                    f"下一步：第 {move.row + 1} 行第 {move.col + 1} 列填 {move.value}（{move.reason}）"
                )
            return True
        except Exception as exc:
            logger.exception("compute_next_step_failed board=%s", self._board_to_log_text())
            self.push_status(str(exc))
            return False

    def on_apply_step(self) -> None:
        self._log_event("user_apply_clicked")
        if not self.current_step and not self._compute_next_step(push_status=False):
            self.push_status("当前局面没有找到可直接确定的一步")
            return
        if not self.last_read or not self.scrcpy.device_size:
            self.push_status("请先读取并计算下一步")
            return
        self._apply_current_step()

    def _apply_current_step(self) -> bool:
        if not self.current_step or not self.last_read or not self.scrcpy.device_size:
            self._log_event("apply_current_step_skipped")
            return False
        try:
            current_image = self.scrcpy.capture_content()
            current_geometry = self._scaled_geometry(
                self.last_read.geometry,
                self.last_read.content_image.shape[:2],
                current_image.shape[:2],
            )
            move = self.current_step
            current_value, _, _ = self.vision.read_single_cell(
                current_image,
                current_geometry,
                self.last_read.templates,
                move.row,
                move.col,
            )
            self._log_event(
                "apply_current_step_precheck",
                row=move.row + 1,
                col=move.col + 1,
                expected=move.value,
                current_value=current_value,
            )
            if current_value != 0:
                self.read_screen()
                self.push_status("目标格状态已变化，已重新同步，请再点一次“下一步”")
                return
            content_shape = self.last_read.content_image.shape[:2]
            board_center = self._cell_center_in_device(self.last_read.geometry, move.row, move.col, content_shape)
            keypad_center = self._keypad_center_in_device(self.last_read.keypad_rect, move.value, content_shape)
            self.scrcpy.tap(*board_center)
            time.sleep(0.05)
            self.scrcpy.tap(*keypad_center)
            time.sleep(0.14)
            value = 0
            score = 0.0
            confirm_image = None
            for _ in range(2):
                confirm_image = self.scrcpy.capture_content()
                confirm_geometry = self._scaled_geometry(
                    self.last_read.geometry,
                    self.last_read.content_image.shape[:2],
                    confirm_image.shape[:2],
                )
                value, score, gap = self.vision.read_single_cell(
                    confirm_image,
                    confirm_geometry,
                    self.last_read.templates,
                    move.row,
                    move.col,
                )
                self._log_event(
                    "apply_current_step_confirm_attempt",
                    row=move.row + 1,
                    col=move.col + 1,
                    expected=move.value,
                    detected=value,
                    score=f"{score:.3f}",
                    gap=f"{gap:.3f}",
                )
                if value == move.value:
                    break
                time.sleep(0.08)
            if value != move.value:
                self.read_screen()
                raise RuntimeError("回填后未确认到目标数字，为避免误输，已停止本次操作")
            self.board[move.row][move.col] = move.value
            self.filled_cells.add((move.row, move.col))
            self.last_read = ReadResult(
                board=[row[:] for row in self.board],
                filled_cells=set(self.filled_cells),
                templates=self.last_read.templates,
                geometry=self.last_read.geometry,
                keypad_rect=self.last_read.keypad_rect,
                content_image=confirm_image,
            )
            self.vision.template_cache.update(self.last_read.templates)
            self.vision.board_cache = [row[:] for row in self.board]
            self.current_step = None
            self._draw_board()
            self._update_display_info(
                f"已填入第 {move.row + 1} 行第 {move.col + 1} 列 = {move.value}，确认分数 {score:.3f}"
            )
            self._log_event(
                "apply_current_step_success",
                row=move.row + 1,
                col=move.col + 1,
                value=move.value,
                score=f"{score:.3f}",
            )
            self.push_status(f"已填入 {move.value}")
            return True
        except Exception as exc:
            logger.exception("apply_current_step_failed board=%s", self._board_to_log_text())
            self.push_status(str(exc))
            return False

    def on_auto_apply(self) -> None:
        self._log_event("user_auto_apply_clicked")
        if self.auto_running:
            return
        if not self.last_read:
            self.read_screen()
        if not self.last_read:
            return
        if self._board_full():
            self.push_status("当前数独已经填满")
            return
        self.auto_running = True
        self.step_button.configure(state="disabled")
        self.apply_button.configure(state="disabled")
        self.auto_button.configure(state="disabled")
        self.read_check.configure(state="disabled")
        self._log_event("auto_apply_started")
        self.push_status("开始自动操作")
        self.root.after(50, self._auto_apply_loop)

    def _auto_apply_loop(self) -> None:
        if not self.auto_running:
            return
        self._log_event("auto_apply_loop_tick")
        if self._board_full():
            self._finish_auto_apply("数独已填满")
            return
        if not self.current_step and not self._compute_next_step(push_status=False):
            self._finish_auto_apply("当前局面没有更多可直接确定的步骤")
            return
        if not self._apply_current_step():
            self._finish_auto_apply()
            return
        if self._board_full():
            self._finish_auto_apply("数独已填满")
            return
        self.root.after(120, self._auto_apply_loop)

    def _finish_auto_apply(self, message: Optional[str] = None) -> None:
        self.auto_running = False
        self.step_button.configure(state="normal")
        self.apply_button.configure(state="normal")
        self.auto_button.configure(state="normal")
        self.read_check.configure(state="normal")
        self._log_event("auto_apply_finished", message=message or "")
        if message:
            self.push_status(message)

    def _board_full(self) -> bool:
        return all(value != 0 for row in self.board for value in row)

    def _cell_center_in_device(
        self, geometry: GridGeometry, row: int, col: int, content_shape: Tuple[int, int]
    ) -> Tuple[int, int]:
        if not self.scrcpy.device_size:
            raise RuntimeError("设备尺寸未知")
        content_h, content_w = content_shape
        left = geometry.x_lines[col]
        right = geometry.x_lines[col + 1]
        top = geometry.y_lines[row]
        bottom = geometry.y_lines[row + 1]
        x = (left + right) / 2
        y = (top + bottom) / 2
        device_w, device_h = self.scrcpy.device_size
        return int(round(x * device_w / content_w)), int(round(y * device_h / content_h))

    def _keypad_center_in_device(
        self, keypad_rect: Rect, value: int, content_shape: Tuple[int, int]
    ) -> Tuple[int, int]:
        if not self.scrcpy.device_size:
            raise RuntimeError("设备尺寸未知")
        content_h, content_w = content_shape
        slot_w = keypad_rect.w / 9
        x = keypad_rect.x + (value - 0.5) * slot_w
        y = keypad_rect.y + keypad_rect.h / 2
        device_w, device_h = self.scrcpy.device_size
        return int(round(x * device_w / content_w)), int(round(y * device_h / content_h))

    def _scaled_geometry(
        self,
        geometry: GridGeometry,
        source_shape: Tuple[int, int],
        target_shape: Tuple[int, int],
    ) -> GridGeometry:
        src_h, src_w = source_shape
        dst_h, dst_w = target_shape
        if src_h == dst_h and src_w == dst_w:
            return geometry
        scale_x = dst_w / src_w
        scale_y = dst_h / src_h
        x_lines = [int(round(value * scale_x)) for value in geometry.x_lines]
        y_lines = [int(round(value * scale_y)) for value in geometry.y_lines]
        rect = Rect(x_lines[0], y_lines[0], x_lines[-1] - x_lines[0], y_lines[-1] - y_lines[0])
        return GridGeometry(rect=rect, x_lines=x_lines, y_lines=y_lines)

    def _capture_read_result(self, retries: int = 2) -> ReadResult:
        last_error: Optional[Exception] = None
        for attempt in range(retries + 1):
            try:
                logger.info("capture_read_result_attempt attempt=%s board=%s", attempt + 1, self._board_to_log_text())
                result = self.vision.read(self.scrcpy.capture_content())
                self.solver.validate(result.board)
                return result
            except Exception as exc:
                last_error = exc
                logger.exception(
                    "capture_read_result_failed attempt=%s board=%s",
                    attempt + 1,
                    self._board_to_log_text(),
                )
                if attempt < retries:
                    time.sleep(0.12)
        raise RuntimeError(str(last_error))

    def _sync_cell_tracking(
        self, new_board: List[List[int]], previous_board: List[List[int]], detected_filled: set
    ) -> None:
        if not self.given_cells:
            current_nonzero = {
                (row, col) for row in range(9) for col in range(9) if new_board[row][col] != 0
            }
            self.filled_cells = set(detected_filled)
            self.given_cells = current_nonzero - self.filled_cells
            return
        reset_needed = False
        for row, col in self.given_cells:
            if new_board[row][col] == 0:
                reset_needed = True
                break
            if previous_board[row][col] != 0 and new_board[row][col] != previous_board[row][col]:
                reset_needed = True
                break
        if reset_needed:
            current_nonzero = {
                (row, col) for row in range(9) for col in range(9) if new_board[row][col] != 0
            }
            self.filled_cells = set(detected_filled)
            self.given_cells = current_nonzero - self.filled_cells
            return
        previous_nonzero = {
            (row, col) for row in range(9) for col in range(9) if previous_board[row][col] != 0
        }
        current_nonzero = {
            (row, col) for row in range(9) for col in range(9) if new_board[row][col] != 0
        }
        appeared = current_nonzero - previous_nonzero - self.given_cells
        self.filled_cells |= appeared | set(detected_filled)
        self.filled_cells = {cell for cell in self.filled_cells if new_board[cell[0]][cell[1]] != 0}
        self.given_cells = current_nonzero - self.filled_cells

    def _update_display_info(self, message: Optional[str] = None) -> None:
        known = sum(1 for row in self.board for value in row if value != 0)
        filled = len(self.filled_cells)
        lines = [f"已识别 {known} 个数字，其中已填并确认 {filled} 个。"]
        if self.current_step:
            move = self.current_step
            lines.append(
                f"当前建议：第 {move.row + 1} 行第 {move.col + 1} 列填 {move.value}（{move.reason}）"
            )
        elif message:
            lines.append(message)
        else:
            lines.append("点击“下一步”后，这里会显示当前可直接确定的一步。")
        self.display_info_var.set("\n".join(lines))

    def _draw_board(self) -> None:
        self.canvas.delete("all")
        size = 500
        offset = 10
        cell = size / 9
        self.canvas.create_rectangle(offset, offset, offset + size, offset + size, fill="#FFFBE1", outline="")
        if self.current_step:
            self._draw_step_highlight(offset, cell)
        for row in range(9):
            for col in range(9):
                x1 = offset + col * cell
                y1 = offset + row * cell
                x2 = x1 + cell
                y2 = y1 + cell
                fill = ""
                if self.current_step and row == self.current_step.row and col == self.current_step.col:
                    fill = "#F7E0B0"
                if fill:
                    self.canvas.create_rectangle(x1, y1, x2, y2, fill=fill, outline="")
                value = self.board[row][col]
                color = "#9C612F"
                if (row, col) in self.filled_cells:
                    color = "#CB7118"
                if self.current_step and row == self.current_step.row and col == self.current_step.col:
                    value = self.current_step.value
                    color = "#1F5AA8"
                if value:
                    self.canvas.create_text(
                        (x1 + x2) / 2,
                        (y1 + y2) / 2,
                        text=str(value),
                        fill=color,
                        font=("Microsoft YaHei UI", 26, "bold"),
                    )
        for index in range(10):
            width = 3 if index % 3 == 0 else 1
            x = offset + index * cell
            y = offset + index * cell
            self.canvas.create_line(offset, y, offset + size, y, fill="#D78826", width=width)
            self.canvas.create_line(x, offset, x, offset + size, fill="#D78826", width=width)

    def _draw_step_highlight(self, offset: int, cell: float) -> None:
        if not self.current_step:
            return
        row = self.current_step.row
        col = self.current_step.col
        box_row = (row // 3) * 3
        box_col = (col // 3) * 3
        for r in range(9):
            for c in range(9):
                if r == row or c == col or (box_row <= r < box_row + 3 and box_col <= c < box_col + 3):
                    x1 = offset + c * cell
                    y1 = offset + r * cell
                    x2 = x1 + cell
                    y2 = y1 + cell
                    self.canvas.create_rectangle(x1, y1, x2, y2, fill="#FEECCC", outline="")

    def run(self) -> None:
        self.root.mainloop()

    def on_close(self) -> None:
        try:
            self._log_event("app_closing")
            self.scrcpy.stop()
        finally:
            self.root.destroy()


if __name__ == "__main__":
    SudokuApp().run()
