# Teleconsult booking setup and video confirmation

The website shows only appointment slots marked `AVAILABLE`. When a patient requests a slot, Apps Script marks it `PENDING` and saves the request. The next addition lets you enter a Google Meet link when approving the request, then prepares a WhatsApp confirmation message for the patient.

## One-time update for video confirmation

1. Open [the latest Code.gs](https://raw.githubusercontent.com/bhupinthakkar001-droid/teleconsult-website/main/booking-backend/Code.gs).
2. Copy the complete code.
3. Open your **Teleconsult Bookings** spreadsheet and choose **Extensions → Apps Script**.
4. Open `Code.gs`, select its contents, and replace them with the copied code.
5. In the line `const DOCTOR_WHATSAPP = '91XXXXXXXXXX';`, enter your WhatsApp number with country code, digits only. Keep your real number only in your private Apps Script editor, not in public GitHub.
6. Press **Ctrl + S**.
7. In the function dropdown, run `setupBookingSystem` once so the `videoLink` column is added to the Requests tab. It should preserve existing rows.
8. Click **Deploy → Manage deployments → pencil icon → New version → Deploy**. Keep the same `/exec` URL.
9. Refresh the website with **Ctrl + Shift + R**.

## Daily workflow

### Add slots

In the **Slots** tab, add one row per appointment. Fill columns A–D only:

| slotId | date | time | status |
|---|---|---|---|
| S001 | YYYY-MM-DD | 10:00 AM | AVAILABLE |
| S002 | YYYY-MM-DD | 11:00 AM | AVAILABLE |

Use a unique slot ID; date format `YYYY-MM-DD`; time like `10:00 AM`; and status exactly `AVAILABLE`. Don't show patients times you aren't actually available to offer.

### Approve a request and send the video link

1. In **Requests**, select the row for the patient's `PENDING` request.
2. Click **Teleconsult bookings → Approve selected request**.
3. Paste a Google Meet link beginning with `https://meet.google.com/` into the prompt. Create the meeting in Google Meet first, and use a link intended for this patient/appointment.
4. The request becomes `CONFIRMED`, the slot stays hidden from the public calendar, and the Meet link is saved in column K (`videoLink`).
5. A dialog gives you a button to open WhatsApp with a prefilled confirmation containing the date, time and meeting link. Check the message and press **Send** yourself. It is not sent automatically.

### Decline a request

Select the pending request row and click **Teleconsult bookings → Decline selected request**. The request is marked declined and the slot is released back to `AVAILABLE`. Notify the patient on WhatsApp.

## Privacy and reliability

- Keep the Google Sheet private; it contains patient names and phone numbers.
- Don't include symptoms or medical reports in the public booking form.
- The patient must be told the appointment is not confirmed until you approve it.
- The free WhatsApp flow prepares a message but does not send it automatically.
- Test with dummy details before using with real patients.
