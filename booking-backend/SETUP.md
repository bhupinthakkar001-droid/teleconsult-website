# Slot booking system — setup guide

The website shows **only slots whose status is `AVAILABLE`**. When a patient submits a request, the booking script checks the slot again, marks it `PENDING`, and saves the request before offering the patient a WhatsApp button. This prevents a second patient from booking that same slot while you review it.

## One-time setup

1. Open [Code.gs](https://raw.githubusercontent.com/bhupinthakkar001-droid/teleconsult-website/main/booking-backend/Code.gs).
2. Copy all the code, then open the private **Teleconsult Bookings** Google Sheet and choose **Extensions → Apps Script**.
3. In Apps Script, replace the existing `Code.gs` content with the new code.
4. In the line `const DOCTOR_WHATSAPP = '91XXXXXXXXXX';`, put your WhatsApp number with country code, digits only. Keep the number only in this private Apps Script editor, not in public GitHub.
5. Press **Ctrl + S**.
6. Click **Deploy → Manage deployments**, click the pencil icon for the current web app, choose **New version**, and click **Deploy**. Keep using the same `/exec` URL.
7. Open the website and refresh with **Ctrl + Shift + R**.

You should not have to change the website URL if you are updating the existing deployment. The Google Sheet must remain private.

## Adding appointments

In the **Slots** tab, add one row per appointment. Fill columns A–D only:

| slotId | date | time | status |
|---|---|---|---|
| S001 | YYYY-MM-DD | 10:00 AM | AVAILABLE |
| S002 | YYYY-MM-DD | 11:00 AM | AVAILABLE |

Use a unique slot ID; date format `YYYY-MM-DD`; time like `10:00 AM`; status exactly `AVAILABLE`. Columns E–H are filled when a patient requests a slot.

## What happens when someone books

1. Patient chooses a date and an available time.
2. On submission, the server rechecks the slot while holding a lock. If another request already took it, the patient sees a message to choose another slot.
3. If it is still available, the slot is immediately marked `PENDING` and the request is saved in the private **Requests** tab.
4. The patient sees a **Continue to WhatsApp** button and must press **Send** to message you. WhatsApp is not sent automatically.
5. After you approve the patient through WhatsApp, open the **Requests** tab, select that request's row, and use **Teleconsult bookings → Approve selected request**. The slot becomes `CONFIRMED` and stays hidden from patients.
6. If you decline, use **Teleconsult bookings → Decline selected request**. The request is marked declined and the slot becomes `AVAILABLE` again.

The website exposes only the public list of available date/time slots. Patient names and phone numbers are stored in the private spreadsheet and are not returned by the public slot-list endpoint.

## Important

- Test the workflow using dummy details before sharing the site with patients.
- Do not enter symptoms or medical reports into the booking form.
- Keep the spreadsheet private and restrict access to you.
- The slot remains pending until you approve/decline it. If a patient doesn't send the WhatsApp message, you can decline that request in the Requests tab to release the slot again.
