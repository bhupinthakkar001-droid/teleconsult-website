# Free slot-based booking setup

This prototype uses a private Google Sheet plus Google Apps Script. The website will only show slots whose status is `AVAILABLE`; submitting a request changes the slot to `PENDING`, preventing another request for the same slot. The doctor confirms or declines the request through WhatsApp, then updates its status using the spreadsheet menu.

## 1. Create the private booking sheet

1. Sign in to the Google account you want to use for bookings.
2. Create a new Google Sheet named **Teleconsult Bookings**.
3. Keep the sheet private. Do not share it publicly; it will contain patients' names and phone numbers.
4. In the sheet, choose **Extensions → Apps Script**.

## 2. Add the backend code

The public repository keeps the doctor's WhatsApp number as a placeholder for privacy. Enter the doctor's number only inside your private Apps Script editor. The patient must tap the WhatsApp button and press Send; messages are not sent automatically by this free setup.

1. Open `booking-backend/Code.gs` in this repository.
2. Copy its complete contents into the Apps Script editor's `Code.gs`, replacing the starter code.
3. In your private Apps Script editor, set `DOCTOR_WHATSAPP` to your WhatsApp number using country code + number, digits only, without a plus sign or spaces. Do not put your personal number into the public GitHub repository.
4. Click Save.
5. In the function dropdown, select `setupBookingSystem` and click Run. Approve the permissions requested by Google.
6. Return to your spreadsheet and reload it. It should contain `Slots` and `Requests` tabs, plus a **Teleconsult bookings** menu.

## 3. Add the slots you want to offer

In the **Slots** tab, add one row for every available appointment. Use these columns:

| slotId | date | time | status |
|---|---|---|---|
| S001 | YYYY-MM-DD | 10:00 AM | AVAILABLE |
| S002 | YYYY-MM-DD | 11:00 AM | AVAILABLE |

- Use a unique ID for every slot.
- Enter dates as text in `YYYY-MM-DD` format and times like `10:00 AM` or `04:30 PM`.
- The status must be exactly `AVAILABLE`.
- Leave columns E–H blank. Do not change slots already marked `PENDING` or `CONFIRMED` without checking the Requests tab.

You control the schedule: add slots when you want to open appointments and stop offering a date by removing its unbooked `AVAILABLE` rows. Do not delete rows that are pending or confirmed.

## 4. Deploy the script as a web app

1. In Apps Script, click **Deploy → New deployment**.
2. Choose the deployment type **Web app**.
3. Set **Execute as** to **Me**.
4. Set **Who has access** to **Anyone** (including users who aren't signed in, if that is the option shown).
5. Click **Deploy** and complete the authorization steps.
6. Copy the web app URL. It should end in `/exec`.

Do not share the Apps Script editor or the Google Sheet with the public. Only the deployed web app should be public.

## 5. Connect the public website

In the website's `index.html`, replace the placeholder `BOOKING_API_URL` with the deployed `/exec` URL. Save the file to `main` and wait for GitHub Pages to update.

## 6. Review and approve requests

1. A patient chooses a currently available slot and submits their name and phone number.
2. The backend marks the slot `PENDING` and stores a request in the private Requests tab.
3. The patient is offered a button to open WhatsApp with a prefilled request message to the doctor. The patient must tap Send in WhatsApp.
4. Discuss and confirm or decline with the patient on WhatsApp.
5. In the Requests tab, select the pending request row. Use **Teleconsult bookings → Approve selected request** or **Decline selected request**. This marks the request accordingly; an approved slot becomes `CONFIRMED`, while a declined slot is released back to `AVAILABLE`.

## Important limitations and privacy

- This is an appointment-request workflow, not automatic confirmation. The appointment is confirmed only after you approve it and tell the patient on WhatsApp.
- The free WhatsApp link opens a prefilled message; it does not send messages automatically. Do not enter patient symptoms or medical records in this booking form.
- The spreadsheet must remain private, and access to it should be limited to you. Review privacy/consent language and applicable telemedicine requirements before accepting real patients.
- Test with dummy information first. Test a duplicate request to verify the slot cannot be requested twice.
