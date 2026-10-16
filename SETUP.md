# RealtorOne Content Sprint: setup (free, about 15 minutes)

The page is one file, `index.html`. No server needed. **It works as it is in preview mode:** with `SUBMIT_ENDPOINT` and `PAYMENT_URL` empty, answers are saved in the browser only (localStorage), the payment step shows a clearly marked "preview" box, and no email or message is sent.

Live page: https://anatomyofacloser-cyber.github.io/realtorone-content-sprint/

There are 3 things to switch on. Do them in this order.

## 1. Save signups in a Google Sheet and send the emails (Apps Script)
1. Create a new Google Sheet with the Gmail account you want the emails to come from. Name it "Content Sprint signups".
2. Open **Extensions > Apps Script**. Delete the sample code, paste everything from `apps-script.gs`, and save.
3. At the top of the script, check the config: `WHATSAPP_NUMBER` is already set. Leave `PAYMENT_URL` empty for now (you will add it in step 2). `OWNER_EMAIL` can stay empty: new-signup notices then go to the Gmail account that owns the sheet.
4. Optional test: pick the function `testEmailsToMe`, click **Run**, approve the permissions. You get a sample confirmation email and a sample welcome email in your own inbox. The Execution log shows `CONFIRMATION_SENT / WELCOME_SENT`.
5. Click **Deploy > New deployment**, type **Web app**. Execute as: **Me**. Who has access: **Anyone**. Click **Deploy** and approve (the script needs your sheet and permission to send email as you).
6. Copy the **Web app URL** (ends in `/exec`) and paste it into `index.html`, at the top of the `<script>`:
   ```js
   var SUBMIT_ENDPOINT = "https://script.google.com/macros/s/XXXX/exec";
   ```
7. Test on the live page with your own email. A row appears in the `Responses` tab, and you get the confirmation email.

If you edit `apps-script.gs` later: paste it over the old one, save, then **Deploy > Manage deployments > pencil > Version: New version > Deploy**. The `/exec` URL stays the same.

## 2. Add the Razorpay payment link (AED 400)
1. Log in to the Razorpay Dashboard. **The account must accept AED:** go to **Settings > Payment Methods > International Payments** and request it if it is not on yet. Razorpay may ask for documents and takes a little time to approve. AED is on Razorpay's supported-currency list for Payment Links.
2. Go to **Payment Links > Create Payment Link**.
3. Choose **AED** in the currency dropdown and enter **400** as the amount. If you do not see AED in the dropdown, international payments are not active yet. Ask Razorpay support.
4. Description: `RealtorOne Content Sprint, 30 days`. Leave the customer details to be filled at checkout. Create the link (it looks like `https://rzp.io/...`).
5. Optional: if your link form has a "callback / redirect URL" option, set it to `https://anatomyofacloser-cyber.github.io/realtorone-content-sprint/?paid=1`. People then land on the Welcome screen after paying. If not, no problem: the page also shows an **"I've completed my payment"** button after they tap pay.
6. Paste the link in **both** files:
   - `index.html`: `var PAYMENT_URL = "https://rzp.io/...";`
   - `apps-script.gs`: `var PAYMENT_URL = 'https://rzp.io/...';` (used for the "Secure my spot" button in the confirmation email). Then redeploy the script as a new version.

When `PAYMENT_URL` is empty the page shows "preview: the payment link isn't set yet" and no real payment can happen.

**Important:** the page cannot see whether someone really paid. "I've completed my payment" (or `?paid=1`) only marks the row **PAYMENT_CLAIMED** and emails you. **Always check the payment in Razorpay before you start someone's content.**

## 3. Check the other settings (top of the `<script>` in `index.html`)
```js
var SUBMIT_ENDPOINT = "";   // step 1
var PAYMENT_URL     = "";   // step 2
var WHATSAPP_NUMBER = "917302317285";   // digits with country code. Empty = WhatsApp buttons hidden
var CURRENCY        = "AED";
var PRICE           = 400;
var REGULAR_PRICE   = "";   // optional, e.g. "590–690": shows "Regular price later: AED 590–690". Empty = hidden
var SPOTS           = 10;
```
- `REGULAR_PRICE` is empty on purpose. Fill it only when you have decided the real later price.
- If you change `CURRENCY`, `PRICE` or the WhatsApp number, change the same lines at the top of `apps-script.gs` too.
- Empty values hide their button/line. WhatsApp buttons appear on the payment step and the welcome screen.

## What you see in the sheet (`Responses` tab)
One row per signup, with every questionnaire answer, plus:
- **Payment status:** `PENDING` (signed up) > `LINK_OPENED` (opened the payment link) > `PAID_CLAIMED` (said they paid. Not verified).
- **Delivery status:** `CONFIRMATION_SENT` after signup, `WELCOME_SENT` after "paid", or `EMAIL_FAILED: ...` / `EMAIL_NOT_SENT: daily quota used up`.
- **Owner notified:** `OWNER_EMAIL_SENT` when you got the "new signup" / "payment claimed" email.
- **Delivery channel:** where they want the daily content (WhatsApp, Email or Both). Send the daily 9 AM content yourself on that channel. The script does not send WhatsApp.

## Limits to know
- Emails go out from the Gmail account that owns the sheet. Free Google accounts can send about 100 emails a day. Emails may land in spam at first: tell the first few people to check.
- The page posts in `no-cors` mode, so it cannot read the script's reply. Check the sheet's Delivery status to see what actually happened.
- The sheet holds names, emails and WhatsApp numbers. Keep it private and only contact people who ticked the consent box.
- No trackers or analytics. The only outside request is the Google Fonts stylesheet; the page works without it.
- The example posts on step 3 are samples built from the person's answers. They use `[bracketed]` placeholders, never invented figures. Keep it that way if you edit the text in `buildSamples()` in `index.html`.

## Hosting
Already published with GitHub Pages (branch `main`, folder `/`). To update the page: edit `index.html` and push to `main`; the live page refreshes in about a minute. `.nojekyll` is included.
