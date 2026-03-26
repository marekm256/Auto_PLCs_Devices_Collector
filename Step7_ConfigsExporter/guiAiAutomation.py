from pathlib import Path
import os, time
import cv2, numpy as np, pyautogui
import shutil
 
INPUT_FOLDER = Path(r"C:\Users\Hatna\Documents\Python\Automatic PLCs Device Exporter\collected_projects\step7")
IMPORTED_FOLDER = INPUT_FOLDER / "imported"
IMPORTED_FOLDER.mkdir(exist_ok=True)

 
PROJECT_MAX   = "project_max.png"
PROJECT_CLOSE = "project_close.png"
PLUS_BTN      = "plus_btn.png"
MACHINE_IMG   = "machine_btn.png"
HW_BTN     = "hw_btn.png"
STATION_BTN = "station_btn.png"
EXPORT_BTN = "export_btn.png"
EXPORT_WINDOW = "export_loaded.png"
BROWSE_BTN = "browse_btn.png"
SAVE_BTN = "save_btn.png"
HW_CLOSE   = "hw_close.png"
OK_BTN = "ok_button.png"
CLOSE_BTN = "close_btn.png"
 
THR = 0.75
POLL_S = 0.5
TIMEOUT_S = 10
 
pyautogui.FAILSAFE = True
 
def _match(tpl_path):
    scr = cv2.cvtColor(np.array(pyautogui.screenshot()), cv2.COLOR_RGB2GRAY)
    tpl = cv2.imread(tpl_path, 0)
    res = cv2.matchTemplate(scr, tpl, cv2.TM_CCOEFF_NORMED)
    _, m, _, p = cv2.minMaxLoc(res)
    return m, p, tpl.shape[:2]
 
def click_icon(tpl_path, thr=THR):
    m, p, (h, w) = _match(tpl_path)
    if m < thr:
        return False
    pyautogui.click(p[0] + w//2, p[1] + h//2)
    return True
 
def dblclick_icon(tpl_path, thr=THR):
    m, p, (h, w) = _match(tpl_path)
    if m < thr:
        return False
    pyautogui.doubleClick(p[0] + w//2, p[1] + h//2)
    return True
 
def wait_for(tpl_path, thr=THR, timeout=TIMEOUT_S):
    t0 = time.time()
    while time.time() - t0 < timeout:
        m, _, _ = _match(tpl_path)
        if m >= thr:
            return True
        time.sleep(POLL_S)
    return False
 
def find_machine_icons_left(img_path, x_max, thr, min_dist):
    scr = cv2.cvtColor(np.array(pyautogui.screenshot()), cv2.COLOR_RGB2GRAY)
    tpl = cv2.imread(img_path, 0)
    h, w = tpl.shape[:2]
 
    res = cv2.matchTemplate(scr, tpl, cv2.TM_CCOEFF_NORMED)
    ys, xs = np.where(res >= thr)
 
    pts = []
    for x, y in zip(xs, ys):
        cx, cy = x + w // 2, y + h // 2
        if cx >= x_max:
            continue
        if all((cx - px) ** 2 + (cy - py) ** 2 > min_dist ** 2 for px, py in pts):
            pts.append((cx, cy))
 
    pts.sort(key=lambda p: p[1])
    return pts

def wait_hw_ready_by_station_export(
    station_tpl=STATION_BTN,
    export_tpl=EXPORT_BTN,
    thr=THR,
    timeout=120,
    click_interval_s=10.0
):
    time.sleep(30)
    t0 = time.time()
    last_click = 0.0
 
    while time.time() - t0 < timeout:
        m_ok, _, _ = _match(OK_BTN)
        if m_ok >= thr:
            click_icon(OK_BTN)
            time.sleep(2)

        m_close, _, _ = _match(CLOSE_BTN)
        if m_close >= thr:
            click_icon(CLOSE_BTN)
            time.sleep(2)

        # 1) export button = ready
        m_exp, _, _ = _match(export_tpl)
        if m_exp >= thr:
            return True

        now = time.time()
        if (now - last_click) >= click_interval_s:
            click_icon(station_tpl, thr=thr)
            last_click = now

        time.sleep(POLL_S)
    return False

# find all projects in INPUT_FOLDER but ignore imported folder
projects = [p for p in INPUT_FOLDER.iterdir() if p.is_dir()
            if p.is_dir() and p.name != "imported"]
n = len(projects)
 
for i, proj in enumerate(projects, 1):
    s7p = next(proj.rglob("*.S7P"), None) or next(proj.rglob("*.s7p"), None)
    print(f"{i}/{n}  {proj.name}  ->  {(s7p.name if s7p else 'NO_S7P_FOUND')}")
    if not s7p:
        continue

    g = proj / "Global"
    for f in g.glob("Language"):
        try: f.unlink()
        except: pass
 
    os.startfile(str(s7p))
    if wait_for(OK_BTN):
        click_icon(OK_BTN)
    time.sleep(2)

    if wait_for(OK_BTN, timeout=TIMEOUT_S):
        click_icon(OK_BTN)
    time.sleep(2)

    click_icon(PROJECT_MAX)                 # maximize project
    time.sleep(2)
 
    click_icon(PLUS_BTN)                    # show tree with machines
    time.sleep(2)
 
    pts = find_machine_icons_left(MACHINE_IMG, 150, THR, 0)
    if pts == 0:
        print("\tNo machines has been found.")

    for j, (x, y) in enumerate(pts):
        pyautogui.click(x, y)               # click machine
        time.sleep(2)

        total = len(pts)
        print(f"\t{j+1}/{total} Machine - Opening Hardware")
 
        if not wait_for(HW_BTN, timeout=TIMEOUT_S):
            print("\tHW_BTN not found ")
            continue
        time.sleep(2)

        click_icon(HW_BTN)
        if not dblclick_icon(HW_BTN):
            print("\tHW_BTN dblclick failed")
            continue
        time.sleep(2)

        if not wait_hw_ready_by_station_export(timeout=300):
            print("\tExport not found")
            continue
        time.sleep(2)

        if not click_icon(EXPORT_BTN):
            print("\tExport button not found")
            continue
        time.sleep(2)

        if not wait_for(EXPORT_WINDOW, timeout=TIMEOUT_S):
            print("\tExport windows not found")
            continue

        if not click_icon(BROWSE_BTN):
            print("\tBrowse button not found")
            continue
        time.sleep(2)

        pyautogui.press('left', presses=20)
        time.sleep(0.2)
        pyautogui.write(f"{proj.name}-{j+1}-", interval=0.05)
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(2)

        if click_icon(SAVE_BTN):
            print(f"\t\t Hardware exported for {proj.name} - {j+1}")
        else:
            print("\tSave button not found")

        if wait_for(OK_BTN, timeout=TIMEOUT_S):
            click_icon(OK_BTN)
            time.sleep(2)
 
        if not click_icon(HW_CLOSE):
            print("  HW_CLOSE not found/click failed")
            continue
        time.sleep(2)
        pyautogui.moveTo(500,500)

    click_icon(PROJECT_CLOSE)
    time.sleep(2)

    target = IMPORTED_FOLDER / proj.name
    if target.exists():
        i = 1
        while True:
            candidate = IMPORTED_FOLDER / f"{proj.name}-{i}"
            if not candidate.exists():
                target = candidate
                break
            i += 1
    shutil.move(str(proj), str(target))
    print(f"[MOVED] {proj}  ->  {target}")