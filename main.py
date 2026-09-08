import pandas as pd
import win32com.client as win32
import time
import traceback
from datetime import datetime
import random

# === CONFIG ===
excel_file = "Emails Data base.xlsx"          # Input file
sheet_name = "SENT"                        # Sheet name
column_name = "Emails"                        # Email column name

output_file = "Email_Status_Final.xlsx"       # Final output file
output_file_2 = "Email_Status_Working.xlsx"   # Intermediate output file
autosave_interval = 600                       # seconds

# Sender rotation: list of dicts with SMTP + Display Name (rotation order)
#
# FILL THIS IN before running.
# Each account must already be signed in inside Outlook on this computer.
# Add as many lines as you need - the script rotates through them in order.
SENDER_IDENTITIES = [
    {"email": "ENTER EMAIL HERE", "name": "ENTER DISPLAY NAME HERE"},
    # {"email": "ENTER EMAIL HERE", "name": "ENTER DISPLAY NAME HERE"},
    # {"email": "ENTER EMAIL HERE", "name": "ENTER DISPLAY NAME HERE"},
]

# How many successful sends before switching to the next sender account
BATCH_SIZE_PER_SENDER = 30

EMAIL_SUBJECT = "ENTER EMAIL SUBJECT HERE"

# === 1️⃣ Read Excel ===
df = pd.read_excel(excel_file, sheet_name=sheet_name)
if "Status" not in df.columns:
    df["Status"] = ""
if "Used Sender" not in df.columns:
    df["Used Sender"] = ""
if "Used Name" not in df.columns:
    df["Used Name"] = ""

# === 2️⃣ Connect to Outlook ===
outlook = win32.Dispatch('outlook.application')
namespace = outlook.GetNamespace("MAPI")

# === 3️⃣ Map available Outlook accounts by SMTP, verify configured senders exist ===
accounts_by_smtp = {}
for account in namespace.Accounts:
    try:
        smtp = account.SmtpAddress.strip().lower()
    except Exception:
        continue
    accounts_by_smtp[smtp] = account

missing = [ident["email"] for ident in SENDER_IDENTITIES if ident["email"].lower() not in accounts_by_smtp]
if missing:
    raise Exception(
        "⚠️ These sender accounts were not found in your Outlook profile: "
        + ", ".join(missing)
    )

# === Helper: build HTML body with the currently active sender email + name injected ===
def build_html_body(active_sender_email: str, active_sender_name: str) -> str:
    return f"""
    <!DOCTYPE html>
    <html>
    <body style="margin:0; padding:0; background:#ffffff;">
        <!-- Wrapper -->
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#ffffff;">
        <tr>
            <td align="center" style="padding:24px 12px;">
            <!-- Card -->
            <table role="presentation" width="680" cellspacing="0" cellpadding="0" border="0" style="width:680px; max-width:680px; background:#ffffff; border-collapse:collapse;">
                <tr>
                <td style="font-family:Segoe UI, Arial, sans-serif; color:#1c1c1c; font-size:15px; line-height:1.7;">

                    <!-- Body copy -->
                    <p style="margin:0 0 16px;">Good day,</p>
                    <p style="margin:0 0 16px;">I hope this email finds you well.</p>

                    <p style="margin:0 0 16px;">
                    As we have surpassed the critical deadlines for MIPS reporting and Improvement Activities, it is more important than ever for every provider and practice to ensure full compliance with CMS requirements under the Quality Payment Program (QPP).
                    </p>

                    <p style="margin:0 0 16px;">
                    Over the past three years (2022, 2023 &amp; 2024), a majority of providers and practices have faced severe MIPS penalties due to non-compliance or incomplete reporting. These penalties are already impacting 2025 reimbursements and will continue to affect every Medicare Part B claim in 2026 and beyond.
                    </p>

                    <!-- Section -->
                    <h3 style="margin:24px 0 8px; font-size:18px; color:#0b66c3; font-weight:700;">Confirming the Penalties:</h3>
                    <p style="margin:0 0 10px;">You can verify the MIPS penalty directly from your Medicare Part B claim EOBs, where you may notice the following codes:</p>
                    <ul style="margin:0 0 16px 20px; padding:0;">
                    <li style="margin:4px 0;">CO 237 &ndash; Legislated/Regulatory Penalty Applied</li>
                    <li style="margin:4px 0;">N 807 &ndash; Payment Adjusted Due to MIPS Non-Compliance</li>
                    </ul>
                    <p style="margin:0 0 16px;">
                    These codes indicate penalties tied to the MIPS program, often resulting in lost reimbursements per provider, per annum.
                    </p>

                    <h3 style="margin:24px 0 8px; font-size:18px; color:#0b66c3; font-weight:700;">Critical Deadline Updates:</h3>
                    <ul style="margin:0 0 16px 20px; padding:0;">
                    <li style="margin:4px 0;">Promoting Interoperability Deadline: <b>July 5th, 2025 (Passed)</b></li>
                    <li style="margin:4px 0;">Improvement Activities Deadline: <b>October 3rd, 2025 (Passed)</b></li>
                    <li style="margin:4px 0;">Reporting Year 2025: <b>We are in the final quarter — time is extremely important.</b></li>
                    </ul>
                    <p style="margin:0 0 16px;">
                    The data to be reported to CMS covers the entire 2025 year, and our team can still manage and complete this reporting for your practice if correct measures are taken at the earliest.
                    </p>

                    <h3 style="margin:24px 0 8px; font-size:18px; color:#0b66c3; font-weight:700;">2024 MIPS Scores Released — Review Your Scores:</h3>
                    <p style="margin:0 0 16px;">
                    CMS has released for PY 2024 MIPS Final Scores. Review your scores on the QPP portal and ensure they are above 75 points to eliminate penalties on your 2026 reimbursements.
                    If your scores falls below 75, it’s critical to start corrective actions and strategic reporting as early as possible.
                    </p>

                    <h3 style="margin:24px 0 8px; font-size:18px; color:#0b66c3; font-weight:700;">Why You Must Start Early:</h3>
                    <ul style="margin:0 0 16px 20px; padding:0;">
                    <li style="margin:4px 0;">Non-reporting or incorrect reporting automatically results in penalties.</li>
                    <li style="margin:4px 0;">Limitation in quality measures are available within most EHR/EMR systems.</li>
                    <li style="margin:4px 0;">Specialty-based practices face even greater risk due to reduced measure availability.</li>
                    </ul>

                    <h3 style="margin:24px 0 8px; font-size:18px; color:#0b66c3; font-weight:700;">Why Choose Prime Well Med Solutions:</h3>
                    <p style="margin:0 0 16px;">
                    As a CMS-Qualified Registry since the early PQRS and Meaningful Use programs, we’ve guided numerous number of practices nationwide through MIPS, MACRA, and also for MVP reporting.
                    </p>
                    <p style="margin:0 0 16px;">
                    In 2024 alone, we partnered with <b>4,500+ providers</b> and not a single one faced a penalty. Our proven workflow ensures full compliance with almost zero workload on your staff.
                    </p>

                    <!-- Feature list box -->
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse:collapse; background:#f6f9fc; border:1px solid #e2e8f0; border-left:4px solid #0b66c3; margin:8px 0 16px;">
                    <tr>
                        <td style="padding:14px 16px; font-family:Segoe UI, Arial, sans-serif; font-size:15px; color:#1c1c1c;">
                        <div style="margin:0 0 8px; font-weight:600;">Our Services Include:</div>
                        <ul style="margin:0 0 8px 18px; padding:0;">
                            <li style="margin:4px 0;">Comprehensive compliance analysis &amp; tailored reporting plans</li>
                            <li style="margin:4px 0;">Proven penalty prevention &amp; incentive attainment</li>
                            <li style="margin:4px 0;">Full CMS submission with accuracy and timeliness</li>
                            <li style="margin:4px 0;">Benchmark reporting &amp; optimization reviews</li>
                            <li style="margin:4px 0;">Audit documentation support (6+ years)</li>
                            <li style="margin:4px 0;">Uncapped training &amp; dedicated account manager</li>
                            <li style="margin:4px 0;">HIPAA-compliant, secure data handling</li>
                        </ul>
                        <div style="margin-top:8px;">
                            <span style="color:#0a9a4a;">&#10004;</span> Penalty Prevention and Incentives attainment&nbsp;&nbsp;
                            <span style="color:#0a9a4a;">&#10004;</span> Minimal Workload for Your Practice&nbsp;&nbsp;
                            <span style="color:#0a9a4a;">&#10004;</span> Proven Track Record Since 2016
                        </div>
                        </td>
                    </tr>
                    </table>

                    <h3 style="margin:24px 0 8px; font-size:18px; color:#0b66c3; font-weight:700;">Short-Time Compliance Promotion:</h3>
                    <p style="margin:0 0 16px;">
                    To assist practices after the missed Improvement Activities deadline, we’re offering a special discounted promotion for full-service MIPS compliance and reporting. We’ll manage your reporting end-to-end from data collection to CMS submission ensuring your practice meets every requirement without disruption to daily operations.
                    </p>

                    <h3 style="margin:24px 0 8px; font-size:18px; color:#0b66c3; font-weight:700;">Response Before It’s Too Late:</h3>
                    <p style="margin:0 0 16px;">
                    We strongly encourage you to start your 2025 MIPS reporting to prevent penalties and safeguard your Medicare reimbursements. Let’s schedule a brief consultation this week to review your current status and ensure compliance before it’s too late.
                    </p>

                    <p style="margin:0 0 18px;">
                    Our team is ready to assist you with reporting, analysis review, and penalty prevention.
                    </p>

                    <!-- Signature -->
                    <p style="margin:0 0 2px; color:#0b66c3; font-weight:700;">Best Regards,</p>
                    <p style="margin:0 2px 2px; color:#0b66c3; font-weight:700;">{active_sender_name}</p>
                    <p style="margin:0 0 16px; color:#576579;">Operations Manager | MID</p>

                    <!-- Divider -->
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
                    <tr><td style="height:1px; line-height:1px; background:#cfe0f5;">&nbsp;</td></tr>
                    </table>

                    <!-- Footer block with logo + contacts -->
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin-top:14px;">
                    <tr>
                        <!-- Logo -->
                        <td valign="top" width="130" style="padding:6px 10px 6px 0;">
                        <img src="https://primewellmedsolutions.com/wp-content/uploads/2024/10/Source-02.png" alt="Prime Well Med Solutions" width="120" style="display:block; border:0; outline:none; text-decoration:none;">
                        </td>
                        <!-- Contacts -->
                        <td valign="top" style="padding:6px 0; font-size:14px;">
                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:100%;">
                            <tr>
                            <td style="padding:2px 0;"><b>Contact:</b> (267) 319-7722</td>
                            </tr>
                            <tr>
                            <td style="padding:2px 0;"><b>Fax:</b> (267) 319-7736</td>
                            </tr>
                            <tr>
                            <td style="padding:2px 0;"><b>Email:</b> <a href="mailto:{active_sender_email}" style="color:#0b66c3; text-decoration:underline;">{active_sender_email}</a></td>
                            </tr>
                            <tr>
                            <td style="padding:2px 0;"><b>Address:</b> 3070 Bristol Pike, 1-114, Bensalem, PA 19020</td>
                            </tr>
                            <tr>
                            <td style="padding-top:8px;">
                                <!-- Social Media Icons -->
                                <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="margin-top:8px;">
                                <tr>
                                    <td style="padding-right:10px;">
                                    <a href="https://web.facebook.com/Primewellmedsolutions/" target="_blank">
                                        <img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook" width="20" style="display:block; border:0;">
                                    </a>
                                    </td>
                                    <td style="padding-right:10px;">
                                    <a href="https://www.linkedin.com/company/prime-well-med-solutions/" target="_blank">
                                        <img src="https://cdn-icons-png.flaticon.com/512/733/733561.png" alt="LinkedIn" width="20" style="display:block; border:0;">
                                    </a>
                                    </td>
                                    <td style="padding-right:10px;">
                                    <a href="https://www.instagram.com/primewellmedsolutions/" target="_blank">
                                        <img src="https://cdn-icons-png.flaticon.com/512/2111/2111463.png" alt="Instagram" width="20" style="display:block; border:0;">
                                    </a>
                                    </td>
                                </tr>
                                </table>
                            </td>
                            </tr>
                        </table>
                        </td>
                    </tr>
                    </table>

                    <!-- Disclaimer -->
                    <p style="margin:18px 0 0; font-size:12px; color:#6b7280;">
                    The content of this email is confidential and intended for the recipient specified in the message. It is strictly forbidden
                    to share any part of this message with any third party without the written consent of the sender. If you received this
                    message by mistake, kindly reply to this message and follow with its deletion, so that we can ensure such a mistake does not occur in the future.
                    </p>

                </td>
                </tr>
            </table>
            <!-- /Card -->
            </td>
        </tr>
        </table>
    </body>
    </html>
    """

# === 4️⃣ Rotation helpers (email + name stay in sync) ===
successful_sends = 0  # count successful sends for rotation

def current_identity():
    idx = (successful_sends // BATCH_SIZE_PER_SENDER) % len(SENDER_IDENTITIES)
    ident = SENDER_IDENTITIES[idx]
    return ident["email"].lower(), ident["name"]

def current_sender_account():
    email, _name = current_identity()
    return accounts_by_smtp[email]

# Log the first active sender
first_email, first_name = current_identity()
print(f"✅ Starting with Outlook account: {current_sender_account().DisplayName} ({first_email}) as '{first_name}'")

# === 5️⃣ Loop through recipients and send emails with rotation ===
start_time = time.time()

for index, row in df.iterrows():
    recipient = row.get(column_name, None)

    if pd.isna(recipient) or str(recipient).strip() == "":
        df.at[index, "Status"] = "Skipped (Empty Email)"
        df.at[index, "Used Sender"] = ""
        df.at[index, "Used Name"] = ""
        continue

    try:
        # Recompute in case rotation boundary crossed
        active_email, active_name = current_identity()
        active_account = current_sender_account()

        # Compose email
        mail = outlook.CreateItem(0)
        mail.To = str(recipient).strip()
        mail.Subject = EMAIL_SUBJECT
        mail.HTMLBody = build_html_body(active_email, active_name)

        mail.ReadReceiptRequested = True
        mail.OriginatorDeliveryReportRequested = True

        # Force account
        # 64209 = dispidSendUsingAccount
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, active_account))

        # Send
        mail.Send()
        delay = random.randint(90, 180)
        time.sleep(delay)  # wait between 1.5 to 3 minutes before next email

        successful_sends += 1
        df.at[index, "Status"] = "Sent"
        df.at[index, "Used Sender"] = active_email
        df.at[index, "Used Name"] = active_name
        print(f"✅ [{successful_sends}] Sent to: {recipient}  | Using: {active_email} ({active_name})")

        # If we just crossed a boundary, log the switch (next loop will pick the next sender)
        if successful_sends % BATCH_SIZE_PER_SENDER == 0:
            next_idx = (successful_sends // BATCH_SIZE_PER_SENDER) % len(SENDER_IDENTITIES)
            next_sender = SENDER_IDENTITIES[next_idx]
            print(f"🔄 Reached {BATCH_SIZE_PER_SENDER} successful sends. Next sender will be: {next_sender['email']} ({next_sender['name']})")
            time.sleep(900)  # brief pause before switching

    except Exception as e:
        error_message = str(e).split('\n')[0]
        # Log whatever identity we attempted for easier debugging
        attempted_email, attempted_name = current_identity()
        df.at[index, "Status"] = f"Failed: {error_message}"
        df.at[index, "Used Sender"] = attempted_email
        df.at[index, "Used Name"] = attempted_name
        print(f"❌ Failed to send to: {recipient} | Using: {attempted_email} ({attempted_name}) | Error: {error_message}")
        traceback.print_exc()

    # === Auto-save progress every autosave_interval seconds (overwrite same file) ===
    elapsed = time.time() - start_time
    if elapsed > autosave_interval:
        df.to_excel(output_file_2, index=False)
        print(f"💾 Auto-saved progress to {output_file_2} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        start_time = time.time()  # reset timer

    time.sleep(2)  # optional delay between emails

# === 6️⃣ Save final Excel file ===
df.to_excel(output_file, index=False)
print(f"\n✅ All emails processed. Final file saved as: {output_file}")
