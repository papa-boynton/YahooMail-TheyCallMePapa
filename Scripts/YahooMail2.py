import imaplib
import ssl
import email
from email.header import decode_header
import sys
import os
import re
import traceback


# ---------- Force UTF-8 output for RunCommand ----------
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="ignore")
except:
    pass

sys.stderr = sys.__stderr__


try:

    # ---------- Read arguments ----------
    EMAIL = sys.argv[1].strip().replace("\x00", "")
    APP_PASSWORD = sys.argv[2].strip().replace("\x00", "")

    MAX_PREVIEWS = 5


    # ---------- Determine script folder ----------
    if getattr(sys, "frozen", False):
        BASE = os.path.dirname(sys.executable)
    else:
        BASE = os.path.dirname(os.path.abspath(__file__))

    ROOT = os.path.dirname(BASE)

    VARIABLE_FILE = os.path.join(
        ROOT,
        "Variables",
        "Variables.inc"
    )


    # ---------- Connect to Yahoo IMAP ----------
    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE

    mail = imaplib.IMAP4_SSL(
        "imap.mail.yahoo.com",
        993,
        ssl_context=context
    )

    mail.login(EMAIL, APP_PASSWORD)

    mail.select("INBOX", readonly=True)

    status, data = mail.search(None, "(NOT SEEN)")
    ids = data[0].split() if status == "OK" else []


    messages = []


    for msg_id in ids[-MAX_PREVIEWS:]:

        status, msg_data = mail.fetch(msg_id, "(BODY.PEEK[])")

        if status != "OK":
            continue

        msg = email.message_from_bytes(msg_data[0][1])


        # ---------- SUBJECT ----------
        raw_subject = msg.get("Subject", "(No subject)")

        decoded_sub, enc = decode_header(raw_subject)[0]

        if isinstance(decoded_sub, bytes):
            decoded_sub = decoded_sub.decode(
                enc or "utf-8",
                errors="ignore"
            )

        subject = decoded_sub.strip()


        # ---------- FROM ----------
        raw_from = msg.get("From", "(Unknown sender)")

        decoded_from, enc = decode_header(raw_from)[0]

        if isinstance(decoded_from, bytes):
            decoded_from = decoded_from.decode(
                enc or "utf-8",
                errors="ignore"
            )

        sender = decoded_from.strip()

        if " <" in sender:
            sender = sender.split("<")[0].strip()


        messages.append((sender, subject))


    mail.logout()


    # ---------- Ensure file exists ----------
    if not os.path.exists(VARIABLE_FILE):

        os.makedirs(
            os.path.dirname(VARIABLE_FILE),
            exist_ok=True
        )

        open(
            VARIABLE_FILE,
            "w",
            encoding="utf-8"
        ).close()


    # ---------- Read Variables.inc (UTF-8 only) ----------
    with open(
        VARIABLE_FILE,
        "r",
        encoding="utf-8",
        errors="ignore"
    ) as f:

        old_lines = f.readlines()


    # ---------- Remove old message lines ----------
    kept_lines = [
        line.rstrip("\n")
        for line in old_lines
        if not re.match(r"^\d+\|", line)
    ]


    # ---------- Update UnreadCount ----------
    count_updated = False

    for i, line in enumerate(kept_lines):

        if line.startswith("UnreadCount="):

            kept_lines[i] = f"UnreadCount={len(messages)}"
            count_updated = True
            break


    if not count_updated:
        kept_lines.append(
            f"UnreadCount={len(messages)}"
        )


    # ---------- Add messages ----------
    for i, (sender, subject) in enumerate(
        messages,
        start=1
    ):

        kept_lines.append(
            f"{i}|{sender} - {subject}"
        )


    new_content = "\n".join(kept_lines)
    old_content = "".join(old_lines).rstrip("\n")


    # ---------- Write file if changed ----------
    if new_content != old_content:

        with open(
            VARIABLE_FILE,
            "w",
            encoding="utf-8",
            newline="\n"
        ) as f:

            f.write(new_content)


    print("YahooMail script ran successfully")


except Exception as e:

    print("Error:", str(e))

    traceback.print_exc()