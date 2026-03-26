from __future__ import annotations

from pathlib import Path
import os
import time

import cv2
import numpy as np
import pyautogui

# =========================
# Configuration
# =========================
INPUT_FOLDER = Path(
    r"C:\Users\Hatna\Documents\Python\Automation_Step7ProjectHwExtractor\collected_projects_test"
)

MSP_WINDOW = "msp_loaded.png"
MSP_BTN = "msp_close.png"
PROJECT_MAX = "project_max.png"
PROJECT_CLOSE = "project_close.png"
PLUS_BTN = "plus_btn.png"
MACHINE_IMG = "machine_btn.png"

HW_BTN = "hw_btn.png"
STATION_BTN = "station_btn.png"
EXPORT_BTN = "export_btn.png"
EXPORT_WINDOW = "export_loaded.png"
BROWSE_BTN = "browse_btn.png"
SAVE_BTN = "save_btn.png"
HW_CLOSE = "hw_close.png"

THR = 0.75
POLL_S = 0.5
TIMEOUT_S = 60

pyautogui.FAILSAFE = True

# =========================
# Operator decisions
# =========================
class Decision:
    CONTINUE = "continue"      # continue current flow
    SKIP_PROJECT = "skip"      # skip current project
    EXIT = "exit"              # stop script


def operator_decision(reason: str) -> str:
    """
    Pause execution and let operator decide what to do.
    E = exit, C = continue, S = skip project
    """
    print("\n" + "=" * 80)
    print("OPERATOR INTERVENTION REQUIRED")
    print(reason)
    print("-" * 80)
    print("[C] Continue  |  [S] Skip project  |  [E] Exit")
    print("=" * 80)

    while True:
        choice = input("Your choice (C/S/E): ").strip().lower()
        if choice == "c":
            return Decision.CONTINUE
        if choice == "s":
            return Decision.SKIP_PROJECT
        if choice == "e":
            return Decision.EXIT
        print("Invalid choice. Use C, S or E.")


def handle_failure(reason: str) -> str:
    """
    Standard failure handler. Returns Decision.*
    """
    return operator_decision(reason)


# =========================
# Vision / clicking helpers
# =========================
def _match(tpl_path: str) -> tuple[float, tuple[int, int], tuple[int, int]]:
    scr = cv2.cvtColor(np.array(pyautogui.screenshot()), cv2.COLOR_RGB2GRAY)
    tpl = cv2.imread(tpl_path, 0)
    if tpl is None:
        # Template missing or unreadable -> treat as hard fail
        return 0.0, (0, 0), (0, 0)

    res = cv2.matchTemplate(scr, tpl, cv2.TM_CCOEFF_NORMED)
    _, m, _, p = cv2.minMaxLoc(res)
    h, w = tpl.shape[:2]
    return m, p, (h, w)


def _click_center(p: tuple[int, int], size: tuple[int, int], double: bool = False) -> None:
    h, w = size
    x = p[0] + w // 2
    y = p[1] + h // 2
    if double:
        pyautogui.doubleClick(x, y)
    else:
        pyautogui.click(x, y)


def click_icon(tpl_path: str, thr: float = THR) -> bool:
    m, p, (h, w) = _match(tpl_path)
    if m < thr or (h, w) == (0, 0):
        return False
    _click_center(p, (h, w), double=False)
    return True


def dblclick_icon(tpl_path: str, thr: float = THR) -> bool:
    m, p, (h, w) = _match(tpl_path)
    if m < thr or (h, w) == (0, 0):
        return False
    _click_center(p, (h, w), double=True)
    return True


def wait_for(tpl_path: str, thr: float = THR, timeout: float = TIMEOUT_S) -> bool:
    t0 = time.time()
    while (time.time() - t0) < timeout:
        m, _, _ = _match(tpl_path)
        if m >= thr:
            return True
        time.sleep(POLL_S)
    return False


def find_machine_icons_left(
    img_path: str,
    x_max: int,
    thr: float,
    min_dist: int,
) -> list[tuple[int, int]]:
    scr = cv2.cvtColor(np.array(pyautogui.screenshot()), cv2.COLOR_RGB2GRAY)
    tpl = cv2.imread(img_path, 0)
    if tpl is None:
        return []

    h, w = tpl.shape[:2]
    res = cv2.matchTemplate(scr, tpl, cv2.TM_CCOEFF_NORMED)
    ys, xs = np.where(res >= thr)

    pts: list[tuple[int, int]] = []
    for x, y in zip(xs, ys):
        cx, cy = x + w // 2, y + h // 2
        if cx >= x_max:
            continue
        if all((cx - px) ** 2 + (cy - py) ** 2 > (min_dist**2) for px, py in pts):
            pts.append((cx, cy))

    pts.sort(key=lambda p: p[1])
    return pts


def wait_hw_ready_by_station_export(
    station_tpl: str = STATION_BTN,
    export_tpl: str = EXPORT_BTN,
    thr: float = THR,
    timeout: float = 120,
    click_interval_s: float = 10.0,
) -> bool:
    # keep your original initial delay (HW config load time)
    time.sleep(30)

    t0 = time.time()
    last_click = 0.0

    while (time.time() - t0) < timeout:
        m_exp, _, _ = _match(export_tpl)
        if m_exp >= thr:
            return True

        now = time.time()
        if (now - last_click) >= click_interval_s:
            click_icon(station_tpl, thr=thr)
            last_click = now

        time.sleep(POLL_S)

    return False


# =========================
# Safe step wrapper
# =========================
def require(
    condition: bool,
    reason: str,
    proj_name: str,
) -> str:
    """
    If condition is False -> pause and ask operator. Returns Decision.*
    """
    if condition:
        return Decision.CONTINUE

    decision = handle_failure(f"[{proj_name}] {reason}")
    return decision


# =========================
# Main
# =========================
def main() -> int:
    projects = [p for p in INPUT_FOLDER.iterdir() if p.is_dir()]
    n = len(projects)

    for i, proj in enumerate(projects, 1):
        # Find .s7p (case-insensitive)
        s7p = next(proj.rglob("*.S7P"), None) or next(proj.rglob("*.s7p"), None)
        print(f"{i}/{n}  {proj.name}  ->  {(s7p.name if s7p else 'NO_S7P_FOUND')}")

        if not s7p:
            continue

        # --- Open project ---
        os.startfile(str(s7p))

        # Close MSP popup if it appears
        if wait_for(MSP_WINDOW, timeout=TIMEOUT_S):
            click_icon(MSP_BTN)

        # Maximize project window
        decision = require(
            click_icon(PROJECT_MAX),
            f"PROJECT_MAX not found/click failed ({PROJECT_MAX})",
            proj.name,
        )
        if decision == Decision.EXIT:
            return 0
        if decision == Decision.SKIP_PROJECT:
            continue
        time.sleep(2)

        # Expand tree
        decision = require(
            click_icon(PLUS_BTN),
            f"PLUS_BTN not found/click failed ({PLUS_BTN})",
            proj.name,
        )
        if decision == Decision.EXIT:
            return 0
        if decision == Decision.SKIP_PROJECT:
            continue
        time.sleep(2)

        # Find machine icons
        pts = find_machine_icons_left(MACHINE_IMG, x_max=150, thr=THR, min_dist=0)
        if not pts:
            decision = handle_failure(f"[{proj.name}] No machine icons found ({MACHINE_IMG}).")
            if decision == Decision.EXIT:
                return 0
            if decision == Decision.SKIP_PROJECT:
                continue

        for j, (x, y) in enumerate(pts):
            total = len(pts)
            print(f"   {j+1}/{total} Machine - Opening Hardware")

            pyautogui.click(x, y)
            time.sleep(2)

            # Wait HW button
            decision = require(
                wait_for(HW_BTN, timeout=TIMEOUT_S),
                f"HW_BTN not found (timeout).",
                proj.name,
            )
            if decision == Decision.EXIT:
                return 0
            if decision == Decision.SKIP_PROJECT:
                break
            if decision == Decision.CONTINUE:
                pass

            # Double click HW
            decision = require(
                dblclick_icon(HW_BTN),
                "HW_BTN dblclick failed.",
                proj.name,
            )
            if decision == Decision.EXIT:
                return 0
            if decision == Decision.SKIP_PROJECT:
                break
            time.sleep(1)

            # Wait export ready
            decision = require(
                wait_hw_ready_by_station_export(timeout=180),
                "Export not found (HW not ready / export button missing).",
                proj.name,
            )
            if decision == Decision.EXIT:
                return 0
            if decision == Decision.SKIP_PROJECT:
                break

            # Click export
            decision = require(
                click_icon(EXPORT_BTN),
                "EXPORT_BTN not found/click failed.",
                proj.name,
            )
            if decision == Decision.EXIT:
                return 0
            if decision == Decision.SKIP_PROJECT:
                break
            time.sleep(2)

            # Wait export window
            decision = require(
                wait_for(EXPORT_WINDOW, timeout=TIMEOUT_S),
                "EXPORT_WINDOW not found (timeout).",
                proj.name,
            )
            if decision == Decision.EXIT:
                return 0
            if decision == Decision.SKIP_PROJECT:
                break

            # Browse
            decision = require(
                click_icon(BROWSE_BTN),
                "BROWSE_BTN not found/click failed.",
                proj.name,
            )
            if decision == Decision.EXIT:
                return 0
            if decision == Decision.SKIP_PROJECT:
                break
            time.sleep(2)

            # Filename typing (keep your original approach)
            pyautogui.press("left", presses=20)
            time.sleep(0.2)
            pyautogui.write(f"{proj.name}-{j+1}-", interval=0.05)
            time.sleep(0.2)
            pyautogui.press("enter")
            time.sleep(2)

            # Save
            if click_icon(SAVE_BTN):
                print(f"\t\tHardware exported for {proj.name} - {j+1}")
            else:
                decision = handle_failure(f"[{proj.name}] SAVE_BTN not found/click failed.")
                if decision == Decision.EXIT:
                    return 0
                if decision == Decision.SKIP_PROJECT:
                    break

            time.sleep(8)

            # Close HW
            if not click_icon(HW_CLOSE):
                decision = handle_failure(f"[{proj.name}] HW_CLOSE not found/click failed.")
                if decision == Decision.EXIT:
                    return 0
                if decision == Decision.SKIP_PROJECT:
                    break

            time.sleep(2)
            pyautogui.moveTo(500, 500)

        # Close project (best-effort)
        if not click_icon(PROJECT_CLOSE):
            decision = handle_failure(f"[{proj.name}] PROJECT_CLOSE not found/click failed.")
            if decision == Decision.EXIT:
                return 0
            # continue or skip just means go next anyway
        time.sleep(2)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
