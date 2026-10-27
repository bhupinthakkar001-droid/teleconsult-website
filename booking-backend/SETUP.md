# Zero-cost teleconsult booking with automatic Google Meet links

The website displays only slots marked `AVAILABLE`. A patient request immediately changes the slot to `PENDING` and saves the request. When the doctor approves it, the script creates a unique Google Meet event/link in the doctor's Google Calendar, marks the booking confirmed, and prepares a WhatsApp message for the patient. No Daily API key or Daily payment method is needed.

Note: the call opens in Google Meet rather than being embedded inside the website. This is the simplest zero-budget route. The Daily embedded-video integration was removed because Daily was blocking calls until a payment method was added.

## One-time setup

1. Open [the latest Code.gs](https://raw.githubusercontent.com/bhupinthakkar001-droid/teleconsult-website/main/booking-backend/Code.gs).
2. Copy all the code.
3. Open the private **Teleconsult Bookings** Google Sheet → **Extensions → Apps Script**.
4. Open `Code.gs`, select all, paste the copied code, and press **Ctrl + S**.
5. In the left sidebar, find **Services** and click the **+**. Select **Google Calendar API**, version **v3**, then click **Add**. This enables the Calendar service used to create unique Meet links.
6. In the Apps Script function dropdown, run `configureTeleconsultation` if you need to set or change your WhatsApp number. Enter country code + number, digits only. If your number was already saved in Script Properties from the previous setup, it remains there.
7. Select `setupBookingSystem` and click **Run** once. It updates the Requests link header while preserving existing request rows and slots.
8. Click **Deploy → Manage deployments → pencil icon → New version → Deploy**. Keep the existing `/exec` URL.
9. If Google asks for additional permission to manage calendar events, review it and approve only if you recognize this script and the action you initiated.
10. Refresh the website using **Ctrl + Shift + R**.

You do not need a Daily account or Daily API key for this version. Do not add a payment method to Daily just for this workflow.

## How the booking works

1. Add appointments to the **Slots** tab with a unique ID, date as `YYYY-MM-DD`, time such as `10:00 AM`, and status `AVAILABLE`.
2. Patients see only currently available slots. A submitted request is stored as `PENDING`, and the slot is hidden from the public calendar.
3. To approve, open **Requests**, select the pending row, and choose **Teleconsult bookings → Approve request & create Google Meet**.
4. The script creates a unique Google Calendar event with a new Meet conference link. The appointment is marked confirmed, and the Meet URL is saved in the Requests tab's `videoLink` column.
5. The confirmation dialog offers **Join Google Meet as doctor**, **Send via WhatsApp**, and **Send via SMS / Messages**. If the patient isn't on WhatsApp, use SMS / Messages. On computers without a messaging app, copy the confirmation message and send it from your phone. Messages are not sent automatically.
6. To decline, select the pending request and choose **Teleconsult bookings → Decline selected request**. The slot becomes `AVAILABLE` again.

The appointment appears on the doctor's primary Google Calendar. The script creates a 30-minute event starting at the selected appointment time. Check your Google Calendar time zone before using it with real patients.

## Privacy and test checklist

- Keep the Google Sheet private; it contains patient names and mobile numbers.
- Send the Meet link only to the patient whose appointment was confirmed. If they aren't on WhatsApp, use normal SMS or call them; SMS charges depend on your phone plan.
- Use a unique Meet conference per appointment; do not reuse meeting links.
- Test with dummy details before taking real consultations.
- This approach does not embed the actual Google Meet UI into your page; it opens the Google Meet call in a new tab.
