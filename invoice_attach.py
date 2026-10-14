# Copyright (c) 2026 Petros Vachaviolos. All rights reserved.
# Written by Petros Vachaviolos (sole author), for personal use only. No use or exploitation without
# the author's written permission, also after any employment ends. See LICENSE.txt.
# ==============================================================================
# vf02_attach.py — Attach τιμολογίων (PDF) σε παραστατικά τιμολόγησης (VF02)
# ==============================================================================
# Ο φάκελος Documents\SAP\SAP GUI = ό,τι PDF εκκρεμεί. Παίρνει από εκεί μόνο τα
# "CC-<αριθμός>.pdf" (ίδια λογική με το Power Query "ΠΕΤΡΟΣ" του VF02.xlsx —
# Text.BetweenDelimiters "-"/"." για τον αριθμό παραστατικού).
#
# Βήμα 0: τα αδιάβαστα emails του Outlook φακέλου ΤΙΜΟΛΟΓΙΑ (τα βάζει εκεί το flow
# ΕΚΤΕΛΩΝΙΣΤΕΣ) -> CC-*.pdf στο SAP GUI -> email read.
#
# Ανά αρχείο: VBELN=αριθμός -> Enter -> GOS attach (%GOS_TOOLBOX -> CREATE_ATTA -> PCATTA_CREA) -> DY_FILENAME ->
# btn[0] -> κλείσιμο GOS panel -> Πίσω (btn[3]) -> σβήσιμο τοπικού PDF (είναι στο SAP + στο email).
#
# Όλα τα SAP element IDs είναι ήδη-validated (από το Power Automate flow που
# έδωσε ο χρήστης) — καμία μαντεψιά.
# ==============================================================================

import sys
import io
import re
import time
import shutil
import logging
import subprocess
from pathlib import Path
from datetime import datetime

import keyring
import win32com.client
import win32event
import tkinter as tk
from tkinter import ttk

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

SCRIPT_DIR = Path(__file__).resolve().parent
_SCRIPT_NAME = Path(__file__).stem
_LOG_DIR = SCRIPT_DIR / "logs"
_LOG_DIR.mkdir(exist_ok=True)
_LOG_FILE = _LOG_DIR / f"{_SCRIPT_NAME}_{datetime.now():%Y%m%d_%H%M%S}.log"
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.FileHandler(_LOG_FILE, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(_SCRIPT_NAME)
log.setLevel(logging.DEBUG)
log.info(f"Log file: {_LOG_FILE}")


def _discard_log_file():
    """Κλείνει το FileHandler και σβήνει το log αρχείο — για runs που δεν
    βρήκαν τίποτα να επεξεργαστούν, ώστε να μη μαζεύονται άδεια logs."""
    root = logging.getLogger()
    for h in list(root.handlers):
        if isinstance(h, logging.FileHandler) and Path(h.baseFilename) == _LOG_FILE:
            h.close()
            root.removeHandler(h)
    try:
        _LOG_FILE.unlink(missing_ok=True)
    except Exception:
        pass

LOCAL_SAP_DIR = Path.home() / "Documents" / "SAP" / "SAP GUI"   # φάκελος εργασίας του SAP GUI
# Εκκρεμή τιμολόγια: το Βήμα 0 τα σώζει από το Outlook κατευθείαν στον φάκελο του SAP GUI
# (πριν: κοινόχρηστος φάκελος δικτύου). Στον ίδιο φάκελο
# ρίχνει και το flow ΕΚΤΕΛΩΝΙΣΤΕΣ τις διασαφήσεις (<MRN>.pdf) -> παίρνουμε ΜΟΝΟ CC-<αριθμός>.pdf.
SOURCE_DIR = LOCAL_SAP_DIR
DONE_DIR = SCRIPT_DIR / "done"   # όσα έχουν ήδη επισυναφθεί

SAP_USER = "YOUR_SAP_USER"
SAP_PASS = keyring.get_password("YOUR_SAP_KEYRING", SAP_USER)
if SAP_PASS is None:
    raise RuntimeError("Δεν βρέθηκε password στο Credential Manager για το SAP.")

# Κοινό mutex σε ΟΛΑ τα scripts που αγγίζουν SAP (ΤΙΜΟΛΟΓΙΑ.py, migo_entry.py,
# IMA-IMC NEW.py, SAP_LOGIN.py) — αν τρέξουν 2 μαζί, το 2ο περιμένει το 1ο να
# τελειώσει αντί να συγκρουστούν πάνω στο ίδιο SAP session.
SAP_MUTEX_NAME = "Global\\SAP_Automation_Lock"


class InvoiceLockedError(Exception):
    """Το τιμολόγιο είναι ήδη ανοιχτό από άλλο χρήστη στο SAP (lock)."""


def fmt_duration(seconds: float) -> str:
    m, s = divmod(int(round(seconds)), 60)
    return f"{m}λ {s}δευτ" if m else f"{s}δευτ"


def close_sap():
    for proc in ["saplogon.exe", "saplgpad.exe", "sapgui.exe", "nwbc.exe"]:
        subprocess.call(["taskkill", "/F", "/T", "/IM", proc],
                        stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)


OUTLOOK_FOLDER = "ΤΙΜΟΛΟΓΙΑ"   # υποφάκελος του Inbox — εκεί βάζει το flow ΕΚΤΕΛΩΝΙΣΤΕΣ τα «ΕΓΓΡΑΦΑ ΕΞΑΓΩΓΩΝ», αδιάβαστα
INVOICE_RE = re.compile(r"^[A-Za-z]+-\d+\.pdf$", re.IGNORECASE)   # π.χ. CC-95342599.pdf


def fetch_from_outlook() -> int:
    """
    Βήμα 0: αδιάβαστα emails του Outlook φακέλου ΤΙΜΟΛΟΓΙΑ -> τα CC-*.pdf στον φάκελο του SAP GUI -> email read.
    Παρακάμπτει όσα υπάρχουν ήδη στο SOURCE_DIR (εκκρεμεί ήδη) ή στο παλιό DONE_DIR (έγινε attach).
    Αν αποτύχει έστω ένα αρχείο, το email μένει αδιάβαστο ώστε να ξαναδοκιμαστεί την επόμενη φορά.
    """
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = ns.GetDefaultFolder(6).Folders[OUTLOOK_FOLDER]   # 6 = Inbox
    unread = [m for m in folder.Items.Restrict("[UnRead] = True") if m.Class == 43]   # 43 = MailItem
    saved = 0
    for mail in unread:
        ok = True
        for i in range(1, mail.Attachments.Count + 1):
            att = mail.Attachments.Item(i)
            name = att.FileName
            if not INVOICE_RE.match(name):
                continue
            if (DONE_DIR / name).exists() or (SOURCE_DIR / name).exists():
                log.info(f"  Υπάρχει ήδη (OK ή εκκρεμή): {name}")
                continue
            try:
                att.SaveAsFile(str(SOURCE_DIR / name))
                saved += 1
                log.info(f"  Αποθηκεύτηκε: {name}")
            except Exception as e:
                ok = False
                log.warning(f"  Αποτυχία αποθήκευσης {name}: {e}")
        if ok:
            mail.UnRead = False
            mail.Save()
        log.info(f"  Email «{mail.Subject}» ({mail.ReceivedTime:%d/%m %H:%M}) -> "
                 f"{'read' if ok else 'μένει αδιάβαστο (retry την επόμενη φορά)'}")
    log.info(f"Βήμα 0 (Outlook): {len(unread)} αδιάβαστα emails, {saved} νέα PDF στο SAP GUI.")
    return saved


def find_invoice_files() -> list[dict]:
    """Ίδια λογική με το Power Query: filename 'CC-<A>.pdf' -> A=αριθμός
    (κείμενο ανάμεσα σε '-' και '.'), B=πλήρες filename."""
    items = []
    for f in SOURCE_DIR.glob("*.pdf"):
        if not INVOICE_RE.match(f.name):
            continue   # π.χ. διασαφήσεις <MRN>.pdf του flow ΕΚΤΕΛΩΝΙΣΤΕΣ — δεν είναι τιμολόγια
        m = re.search(r"-(.+)\.", f.name)
        if not m:
            log.warning(f"  Παράλειψη (δεν βρέθηκε αριθμός στο filename): {f.name}")
            continue
        items.append({"A": m.group(1), "B": f.name})
    return items


def rename_for_new_number(old_name: str, new_a: str) -> str:
    """Αντικαθιστά τον αριθμό τιμολογίου μέσα στο filename ('CC-<A>.pdf'),
    κρατώντας ό,τι είναι πριν το '-' και την κατάληξη μετά την τελεία."""
    m = re.search(r"-(.+)\.", old_name)
    if not m:
        raise ValueError(f"Δεν βρέθηκε pattern αριθμού στο filename: {old_name}")
    return old_name[:m.start(1)] + new_a + old_name[m.end(1):]


def show_preview_banner(items: list[dict]) -> list[dict] | None:
    """GUI banner προεπισκόπησης με checkboxes (όλα επιλεγμένα by default) —
    επιστρέφει τα επιλεγμένα items, ή None αν ακυρώθηκε."""
    cols = ["A", "B"]

    DARK_GREEN = "#1B4332"
    OFF_WHITE = "#F7F5EE"
    GREEN_TINT = "#DCE7DD"

    HEADER_LABELS = {"A": "ΤΙΜΟΛΟΓΙΟ", "B": "PDF"}

    root = tk.Tk()
    root.title("ΤΙΜΟΛΟΓΙΑ — Προεπισκόπηση προς επισύναψη")
    root.configure(bg=OFF_WHITE)
    root.attributes("-topmost", True)
    root.lift(); root.focus_force()

    header = tk.Frame(root, bg=DARK_GREEN)
    header.pack(fill="x")
    tk.Label(
        header, text=f"ΤΙΜΟΛΟΓΙΑ — {len(items)} PDF προς επισύναψη",
        bg=DARK_GREEN, fg="#FFFFFF", font=("Consolas", 14), pady=14
    ).pack()

    toolbar = tk.Frame(root, bg=OFF_WHITE, pady=8)
    toolbar.pack(fill="x", padx=12)

    body = tk.Frame(root, bg=OFF_WHITE)
    body.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview.Heading", background=GREEN_TINT, foreground=DARK_GREEN,
                     font=("Consolas", 11))
    style.configure("Treeview", font=("Consolas", 12), rowheight=30)

    tree_cols = ("chk",) + tuple(cols)
    visible_rows = min(max(len(items), 10), 20)  # min 10 καταχωρήσεις, μετά τις 20 scroll
    tree = ttk.Treeview(body, columns=tree_cols, show="headings", height=visible_rows)
    tree.heading("chk", text="X")
    tree.column("chk", width=36, anchor="center", stretch=False)
    for c in cols:
        max_len = max([len(HEADER_LABELS[c])] + [len(str(r.get(c, ""))) for r in items])
        tree.heading(c, text=HEADER_LABELS[c])
        tree.column(c, width=min(max(max_len * 7 + 16, 70), 190), anchor="center")

    checked = {}
    for r in items:
        iid = tree.insert("", "end", values=("☒",) + tuple(r.get(c, "") for c in cols))
        checked[iid] = True  # όλα επιλεγμένα by default

    vsb = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    count_lbl = tk.Label(toolbar, text=f"{len(items)} / {len(items)}",
                          bg=OFF_WHITE, fg=DARK_GREEN, font=("Consolas", 11))

    def update_count():
        n = sum(1 for v in checked.values() if v)
        count_lbl.config(text=f"{n} / {len(items)}")

    def on_click(event):
        region = tree.identify_region(event.x, event.y)
        col = tree.identify_column(event.x)
        if region != "cell" or col != "#1":  # μόνο η στήλη chk εναλλάσσει επιλογή
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        checked[iid] = not checked[iid]
        vals = list(tree.item(iid, "values"))
        vals[0] = "☒" if checked[iid] else "☐"
        tree.item(iid, values=vals)
        update_count()

    def start_edit_A(event):
        """Διπλό κλικ στη στήλη A -> επεξεργάσιμο πεδίο (π.χ. λάθος/ελλιπής αριθμός)."""
        region = tree.identify_region(event.x, event.y)
        col = tree.identify_column(event.x)
        if region != "cell" or col != "#2":  # #2 = στήλη "A"
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        x, y, w, h = tree.bbox(iid, col)
        current = tree.set(iid, "A")

        entry = tk.Entry(tree, font=("Consolas", 12))
        entry.insert(0, current)
        entry.select_range(0, "end")
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()

        def commit(_event=None):
            new_val = entry.get().strip()
            entry.destroy()
            if new_val and new_val != current:
                tree.set(iid, "A", new_val)

        entry.bind("<Return>", commit)
        entry.bind("<Escape>", lambda e: entry.destroy())
        entry.bind("<FocusOut>", commit)

    def select_all():
        for iid in checked:
            checked[iid] = True
            vals = list(tree.item(iid, "values")); vals[0] = "☒"
            tree.item(iid, values=vals)
        update_count()

    def deselect_all():
        for iid in checked:
            checked[iid] = False
            vals = list(tree.item(iid, "values")); vals[0] = "☐"
            tree.item(iid, values=vals)
        update_count()

    tree.bind("<Button-1>", on_click)
    tree.bind("<Double-Button-1>", start_edit_A)

    tk.Button(toolbar, text="✔ Επιλογή Όλων", font=("Consolas", 11),
              bg=GREEN_TINT, fg=DARK_GREEN, relief="flat", padx=12, pady=6,
              cursor="hand2", command=select_all).pack(side="left", padx=(0, 8))
    tk.Button(toolbar, text="✘ Αποεπιλογή Όλων", font=("Consolas", 11),
              bg=GREEN_TINT, fg=DARK_GREEN, relief="flat", padx=12, pady=6,
              cursor="hand2", command=deselect_all).pack(side="left")
    count_lbl.pack(side="right")

    result = {"items": None}

    def on_ok():
        final = []
        for i, iid in enumerate(tree.get_children()):
            if not checked[iid]:
                continue
            item = items[i]
            new_a = tree.set(iid, "A")
            if new_a != item["A"]:
                old_b = item["B"]
                new_b = rename_for_new_number(old_b, new_a)
                try:
                    (SOURCE_DIR / old_b).rename(SOURCE_DIR / new_b)
                    log.info(f"  Μετονομάστηκε (διόρθωση αριθμού): {old_b} -> {new_b}")
                    item = {"A": new_a, "B": new_b}
                except Exception as e:
                    log.exception(f"  Αποτυχία μετονομασίας {old_b} -> {new_b}: {e}")
                    continue
            final.append(item)
        result["items"] = final
        root.destroy()

    def on_cancel():
        result["items"] = None
        root.destroy()

    btns = tk.Frame(root, bg=OFF_WHITE)
    btns.pack(pady=(0, 16))
    tk.Button(
        btns, text="OK", font=("Consolas", 13),
        bg=DARK_GREEN, fg="#FFFFFF", relief="flat", padx=26, pady=8,
        cursor="hand2", command=on_ok
    ).pack(side="left", padx=10)
    tk.Button(
        btns, text="Ακύρωση", font=("Consolas", 13),
        bg=GREEN_TINT, fg=DARK_GREEN, relief="flat", padx=26, pady=8,
        cursor="hand2", command=on_cancel
    ).pack(side="left", padx=10)

    root.update_idletasks()
    req_w, req_h = 460, root.winfo_reqheight()
    sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
    root.geometry(f"{req_w}x{req_h}+{(sw - req_w) // 2}+{(sh - req_h) // 2}")

    root.mainloop()
    return result["items"]


def sap_login():
    close_sap(); time.sleep(2)
    subprocess.Popen('start "" saplogon.exe', shell=True)
    time.sleep(3)

    sap_gui_auto = None
    for _ in range(30):
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            break
        except Exception:
            time.sleep(1)
    if sap_gui_auto is None:
        raise RuntimeError("SAPGUI scripting engine δεν βρέθηκε (timeout).")

    application = sap_gui_auto.GetScriptingEngine
    connection = application.OpenConnection("YOUR_SAP_CONNECTION", True)
    session = connection.Children(0)
    time.sleep(1)

    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "YOUR_CLIENT"
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USER
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASS
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "EL"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(3)

    session.findById("wnd[0]/tbar[0]/okcd").text = "VF02"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    time.sleep(2)
    session.findById("wnd[0]").maximize()
    time.sleep(1)

    return session


def vf02_attach_one(session, item: dict):
    a, b = item["A"], item["B"]

    session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = a
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)

    sbar_text = session.findById("wnd[0]/sbar/pane[0]").text
    if "είναι ήδη σε επεξεργασία (από το χρήστη" in sbar_text:
        raise InvoiceLockedError(f"Τιμολόγιο {a} locked: {sbar_text}")

    # Move ΜΕΤΑ την είσοδο στο τιμολόγιο (όχι πριν) — αν κάτι πάει στραβά στο
    # άνοιγμα (π.χ. lock από άλλο χρήστη), το PDF μένει στο network folder.
    src = SOURCE_DIR / b
    local = LOCAL_SAP_DIR / b
    if not local.exists():
        if not src.exists():
            raise FileNotFoundError(f"Δεν βρέθηκε ούτε στο network ούτε τοπικά: {b}")
        shutil.move(str(src), str(local))
        log.info(f"  PDF μετακινήθηκε τοπικά: {b}")

    session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
    session.findById("wnd[0]/shellcont/shell").pressContextButton("CREATE_ATTA")
    session.findById("wnd[0]/shellcont/shell").selectContextMenuItem("PCATTA_CREA")

    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = b
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(2)

    session.findById("wnd[0]").maximize()
    try:
        session.findById("wnd[0]/shellcont").close()
    except Exception:
        pass
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    time.sleep(1)

    # Ανέβηκε στο SAP (και το αρχικό υπάρχει στο email) -> το τοπικό αντίγραφο σβήνεται
    local.unlink(missing_ok=True)
    log.info(f"  PDF σβήστηκε (είναι πλέον στο SAP): {b}")


def main():
    try:
        fetch_from_outlook()
    except Exception as e:
        log.exception(f"Βήμα 0 (Outlook) απέτυχε — συνεχίζω με ό,τι υπάρχει ήδη στον φάκελο: {e}")

    items = find_invoice_files()
    log.info(f"Βρέθηκαν {len(items)} τιμολόγια (CC-*.pdf) στο φάκελο SAP GUI.")
    for it in items:
        log.info(f"  A={it['A']}  B={it['B']}")

    if not items:
        log.info("Τίποτα να επεξεργαστεί.")
        _discard_log_file()
        return

    selected = show_preview_banner(items)
    if not selected:
        log.info("Ακυρώθηκε από τον χρήστη (ή καμία επιλογή).")
        return
    items = selected
    t0 = time.perf_counter()  # από το OK στο banner, όχι από το SAP login

    log.info("\nSAP login...")
    session = sap_login()

    locked = []  # items που ήταν locked, προς retry στο τέλος
    total = len(items)
    for idx, item in enumerate(items, 1):
        log.info(f"\n[{idx}/{total}] Τιμολόγιο={item['A']}  ({item['B']})")
        try:
            vf02_attach_one(session, item)
        except InvoiceLockedError as e:
            log.warning(f"  {e} — παραλείπεται προσωρινά, retry στο τέλος.")
            locked.append(item)
        except Exception as e:
            log.exception(f"  ΣΦΑΛΜΑ: {e}")
            close_sap()
            return

    if locked:
        log.info(f"\nRetry pass για {len(locked)} locked τιμολόγια...")
        still_locked = []
        for idx, item in enumerate(locked, 1):
            log.info(f"\n[retry {idx}/{len(locked)}] Τιμολόγιο={item['A']}  ({item['B']})")
            try:
                vf02_attach_one(session, item)
            except InvoiceLockedError as e:
                log.warning(f"  {e} — ΑΚΟΜΑ locked, χρειάζεται χειροκίνητο follow-up.")
                still_locked.append(item)
            except Exception as e:
                log.exception(f"  ΣΦΑΛΜΑ: {e}")
                close_sap()
                return
        if still_locked:
            log.warning(
                f"\n{len(still_locked)} τιμολόγια ΠΑΡΑΜΕΝΟΥΝ locked μετά το retry "
                f"(χρειάζονται χειροκίνητο follow-up): "
                + ", ".join(it["A"] for it in still_locked)
            )
    else:
        still_locked = []

    elapsed = time.perf_counter() - t0
    done = total - len(still_locked)
    avg = elapsed / done if done else 0
    log.info(f"\nΣύνολο χρόνου: {fmt_duration(elapsed)}  |  {done}/{total} επισυνάφθηκαν  |  "
             f"AVG: {avg:.1f} sec/τιμολόγιο")

    log.info("\nΟλοκληρώθηκαν όλα τα attachments.")
    close_sap()


if __name__ == "__main__":
    mutex = win32event.CreateMutex(None, False, SAP_MUTEX_NAME)
    log.info("Αναμονή SAP mutex (αν τρέχει άλλο script στο SAP)...")
    win32event.WaitForSingleObject(mutex, win32event.INFINITE)
    try:
        main()
    except Exception as e:
        log.exception(f"ΣΦΑΛΜΑ: {e}")
    finally:
        win32event.ReleaseMutex(mutex)
