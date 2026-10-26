# Teleconsult booking + built-in video

The website shows only appointment slots marked `AVAILABLE`. When a patient requests a slot, Apps Script marks it `PENDING` and saves it. When you approve it, Apps Script automatically creates a private Daily video room, generates separate patient and doctor access links, marks the appointment confirmed, and opens a WhatsApp message for you to send to the patient.

## One-time setup

1. Open [the latest Code.gs](https://raw.githubusercontent.com/bhupinthakkar001-droid/teleconsult-website/main/booking-backend/Code.gs).
2. Copy all the code.
3. Open the private **Teleconsult Bookings** Google Sheet → **Extensions → Apps Script**.
4. Open `Code.gs`, select its contents, and replace it with the copied code.
5. Press **Ctrl + S**.
6. In the function dropdown, select `configureTeleconsultation` and click **Run**.
7. When prompted, enter the doctor's WhatsApp number with country code (digits only), then paste the Daily API key from the Daily dashboard. Both are saved in private Apps Script properties; neither is placed in the public GitHub source. Never paste the Daily API key into GitHub or this chat.
8. Select `setupBookingSystem` and click **Run** once. It creates or updates the `patientVideoLink` and `doctorVideoLink` columns in Requests and preserves existing rows.
9. Click **Deploy → Manage deployments → pencil icon → New version → Deploy**. Keep the same `/exec` URL.
10. Refresh the website using **Ctrl + Shift + R**.

Google may ask you to authorize external requests to Daily when you run the new code. Review the prompts, and only approve them if you recognize this project and the action you initiated.

## Get your Daily API key

Sign in to [Daily](https://dashboard.daily.co/). Open the developer/API-key area and copy your API key. The key goes only into the private Apps Script prompt in Step 7. Daily says the Video SDK starts with 10,000 free participant-minutes per month; check the current [Daily pricing page](https://www.daily.co/pricing/) and keep an eye on usage before taking real consultations.

## Daily workflow after setup

1. Add available appointments in the **Slots** tab: unique slot ID, date in `YYYY-MM-DD`, time such as `10:00 AM`, and status `AVAILABLE`.
2. Patients see only `AVAILABLE` slots. On submission, the chosen slot becomes `PENDING` and is removed from the public list.
3. To approve a request, open **Requests**, select the pending row, then choose **Teleconsult bookings → Approve request & create video room**.
4. A private Daily room is created automatically. The row is marked `CONFIRMED`; the patient link is saved in column K and the doctor's link in column L.
5. A dialog offers **Join as doctor** and **Open WhatsApp to send patient link**. Send the prefilled WhatsApp message yourself; WhatsApp does not send automatically.
6. To decline a request, select its row and choose **Teleconsult bookings → Decline selected request**. The slot becomes `AVAILABLE` again.

The doctor should open their own doctor link; do not forward it to the patient. The patient link is private and should only be sent to that patient. The room is private and tokens expire around the appointment time.

## Important privacy and usage notes

- Keep the Google Sheet private. It contains patient names and phone numbers.
- Do not include symptoms or medical reports in the public booking form.
- The website uses the Daily Prebuilt video interface embedded in `video.html`; the links contain short-lived access tokens.
- Do not enable recording for clinical visits unless you have a lawful reason, proper patient consent, and secure retention arrangements.
- Test end-to-end with dummy details before real patients use the system.
