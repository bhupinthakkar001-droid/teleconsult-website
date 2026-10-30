# Teleconsult booking + UPI payment setup

This system uses GitHub Pages for the patient website and Google Apps Script + a private Google Sheet for booking operations. UPI collection is a zero-gateway, manual-verification workflow. The website does **not** connect to a bank or verify credits automatically.

## Update the Apps Script backend

1. Open the latest [Code.gs](https://raw.githubusercontent.com/bhupinthakkar001-droid/teleconsult-website/main/booking-backend/Code.gs), copy the complete file, and paste it into the bound Apps Script project for the private **Teleconsult Bookings** Google Sheet, replacing the existing `Code.gs`.
2. Open the latest [Admin.html](https://raw.githubusercontent.com/bhupinthakkar001-droid/teleconsult-website/main/booking-backend/Admin.html). In the same Apps Script project, open the existing HTML file named **Admin** and replace its contents. If it does not exist, add an HTML file named exactly `Admin`.
3. Save both files.
4. In the Apps Script editor, confirm the **Google Calendar API v3** service is enabled under Services. This is used to create unique Google Meet links.
5. Run `setupBookingSystem` once. It preserves existing slots and requests and adds the payment columns to the Requests sheet.
6. Deploy → Manage deployments → pencil icon → **New version** → Deploy. Keep the existing Web App `/exec` URL so the patient website remains connected.
7. Open the doctor dashboard using the existing Web App URL with `?action=admin`, then sign in.

## Configure the UPI details

1. In the doctor dashboard, open **Payment setup**.
2. Enter the UPI ID where consultation fees should be received, for example `name@bank`.
3. Enter the payee name exactly as patients should see it.
4. Optionally enter an HTTPS image URL for your UPI QR code. You can leave this blank; the patient page will still show the UPI ID, a copy button and an **Open UPI app to pay** link with the correct amount pre-filled.
5. Press **Save payment details**. The UPI ID and payee name are public payment instructions; never enter a UPI PIN, OTP, bank password or card details.

The patient website is updated through GitHub Pages. Refresh it after configuring payment details.

## Patient payment flow

1. The patient chooses Regular (₹500) or Priority (₹2,000), a date and time, and completes the intake form and required consents.
2. The patient pays the exact amount using the displayed UPI ID, QR image (if configured), or UPI app link.
3. The patient enters the UPI transaction reference / UTR and checks the payment declaration before submitting.
4. The request is held as **PENDING**. The submitted UPI reference is stored in the Requests sheet with the server-calculated fee. Submission of a reference is not proof that funds arrived.

## Verify a payment and confirm an appointment

1. Sign in to the doctor dashboard and open **Pending**.
2. Compare the patient's transaction reference and exact amount with the actual credit in your bank or UPI app. Do not rely only on the screenshot or reference entered by the patient.
3. Once you have independently confirmed the credit, press **Verify payment**.
4. The **Approve & create Meet** button becomes available. Approve the request to create a unique Google Meet link and confirm the slot.
5. Send the confirmation and Meet link to the patient using the dashboard's WhatsApp/SMS actions.

If the payment cannot be confirmed, do not verify or approve the appointment. Decline the booking and handle any refund manually through your bank/UPI app. Cancellations and refunds are not automated by this system.

## Important operational notes

- Payment references are stored in the private Requests sheet. Keep the Google Sheet private.
- A UPI ID and QR image may be visible to any website visitor; use only the account you intend to receive consultation payments into.
- The doctor must manually verify every credit. This implementation does not integrate with Razorpay, Cashfree, Paytm, a bank API or an automatic payment webhook.
- Never ask patients to share a UPI PIN, OTP, card number or banking password.
- The appointment remains unconfirmed until payment is verified and the doctor approves it.
- Keep the Privacy Notice aligned with the information collected by the form.
- The Google Meet call opens in a new tab; it is not embedded in the website.
