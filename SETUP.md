# RealtorOne Content Sprint: setup (free, about 15 minutes)

The page is one file, `index.html`. No server needed. `SUBMIT_ENDPOINT` and `PAYMENT_URL` are both set, so signups go to your Google Sheet and the Pay button opens the real Razorpay link. (If `SUBMIT_ENDPOINT` is ever emptied the page goes back to preview mode: answers stay in the visitor's browser only.)

Live page: https://anatomyofacloser-cyber.github.io/realtorone-content-sprint/

Steps 1 and 2 are done. **Step 4 (automation) is new: you need to redeploy and run `setupTriggers` once.**

## 1. Save signups in a Google Sheet and send the emails (Apps Script)
1. Create a new Google Sheet with the Gmail account you want the emails to come from. Name it "Content Sprint signups".
2. Open **Extensions > Apps Script**. Delete the sample code, paste everything from `apps-script.gs`, and save.
3. At the top of the script, check the config: `WHATSAPP_NUMBER` is already set. `PAYMENT_URL` is already set. `OWNER_EMAIL` can stay empty: new-signup notices then go to the Gmail account that owns the sheet.
4. Optional test: pick the function `testEmailsToMe`, click **Run**, approve the permissions. You get a sample confirmation email and a sample welcome email in your own inbox. The Execution log shows `CONFIRMATION_SENT / WELCOME_SENT`.
5. Click **Deploy > New deployment**, type **Web app**. Execute as: **Me**. Who has access: **Anyone**. Click **Deploy** and approve (the script needs your sheet and permission to send email as you).
6. Copy the **Web app URL** (ends in `/exec`). It is already pasted into `index.html` (`SUBMIT_ENDPOINT`). Only change it there if you ever create a brand-new deployment.
7. Test on the live page with your own email. A row appears in the `Responses` tab, and you get the confirmation email.

If you edit `apps-script.gs` later: paste it over the old one, save, then **Deploy > Manage deployments > pencil > Version: New version > Deploy**. The `/exec` URL stays the same.

## 2. Razorpay payment link (AED 400): already set
`PAYMENT_URL` is **already set** to `https://rzp.io/rzp/ZRf2Lllg` in both `index.html` and `apps-script.gs`. If you ever need a new link, the steps below show how to create one and where to paste it.

1. Log in to the Razorpay Dashboard. **The account must accept AED:** go to **Settings > Payment Methods > International Payments** and request it if it is not on yet. Razorpay may ask for documents and takes a little time to approve. AED is on Razorpay's supported-currency list for Payment Links.
2. Go to **Payment Links > Create Payment Link**.
3. Choose **AED** in the currency dropdown and enter **400** as the amount. If you do not see AED in the dropdown, international payments are not active yet. Ask Razorpay support.
4. Description: `RealtorOne Content Sprint, 30 days`. Leave the customer details to be filled at checkout. Create the link (it looks like `https://rzp.io/...`).
5. Optional: if your link form has a "callback / redirect URL" option, set it to `https://anatomyofacloser-cyber.github.io/realtorone-content-sprint/?paid=1`. People then land on the Welcome screen after paying. If not, no problem: the page also shows an **"I've completed my payment"** button after they tap pay.
6. Paste the link in **both** files:
   - `index.html`: `var PAYMENT_URL = "https://rzp.io/...";`
   - `apps-script.gs`: `var PAYMENT_URL = 'https://rzp.io/...';` (used for the "Secure my spot" button in the confirmation email). Then redeploy the script as a new version.

(If you clear `PAYMENT_URL`, the page goes back to preview.) When `PAYMENT_URL` is empty the page shows "preview: the payment link isn't set yet" and no real payment can happen.

**Important:** the page cannot see whether someone really paid. "I've completed my payment" (or `?paid=1`) only marks the row **PAID_CLAIMED** and emails you. **Always check the payment in Razorpay before you start someone's content.**

## 3. Check the other settings (top of the `<script>` in `index.html`)
```js
var SUBMIT_ENDPOINT = "https://script.google.com/macros/s/.../exec";   // step 1: already set
var PAYMENT_URL     = "https://rzp.io/rzp/ZRf2Lllg";   // step 2: already set
var WHATSAPP_NUMBER = "917302317285";   // digits with country code. Empty = WhatsApp buttons hidden
var CURRENCY        = "AED";
var PRICE           = 400;
var REGULAR_PRICE   = 610;  // shown struck through next to the founding price AED 400. Empty = hidden
var SPOTS           = 10;
```
- `REGULAR_PRICE` is 610: the page shows "Regular price ~~AED 610~~" and "Founding offer — you get it for AED 400". Make it empty to show only AED 400.
- If you change `CURRENCY`, `PRICE`, `REGULAR_PRICE` or the WhatsApp number, change the same lines at the top of `apps-script.gs` too.
- Empty values hide their button/line. WhatsApp buttons appear on the payment step and the welcome screen.

## 4. Automation: payment reminder + daily 9 AM content (new)
**One-time steps (about 3 minutes):**
1. In the Apps Script editor paste the new `apps-script.gs` over the old one and save. Keep your own `OWNER_EMAIL` if you set one.
2. **Deploy > Manage deployments > pencil (Edit) > Version: New version > Deploy.** The `/exec` URL stays the same. (Without this, the live form still runs the OLD script.)
3. Pick the function **`setupTriggers`** and click **Run**. Approve the permissions (the new one is "run when you are not here"). It creates the two triggers and the `Content Today` tab. Running it again is safe: it replaces the triggers, never doubles them. Check under the clock icon (Triggers).

**What runs by itself:**
- **Every hour, `sendReminders`:** emails ONE reminder to people who signed up 24+ hours ago, ticked consent, and whose Payment status is not `PAID` / `PAID_CLAIMED`. It writes the time in **Reminder sent**, so nobody gets a second one. Signups older than 7 days are not reminded (`REMINDER_MAX_AGE_DAYS`). The reminder says their details are saved, shows the AED 400 offer and the Razorpay button, and tells them to reply if they already paid.
- **Every day around 9 AM Dubai, `sendDailyContent`:** emails everyone whose Payment status is exactly **`PAID`**, for 30 days from their **Content start date** (blank = the first run after you set PAID, so the first email comes the same morning or the next 9 AM). It skips anyone with nothing to send today, anyone already emailed today, and anyone past day 30.

**Your daily routine (about 2 minutes):**
1. When money arrives in Razorpay, set that row's **Payment status** to `PAID` (and optionally a **Content start date**).
2. Each day, before 9 AM Dubai, open the **`Content Today`** tab and add a row:
   | Date | For | Link | Note |
   |---|---|---|---|
   | leave blank (= today) or `2026-10-20` | `ALL`, one person's email, or their Submission ID | a Google Drive / Canva / Notion link (https://...) | optional short text |

   Blank date = "for today": after it is sent the script stamps today's date on that row, so it is not sent again tomorrow. You can add several rows (several links) and they all go in one email. A row with no link and no note is ignored.
3. Check **Delivery status** in `Responses` the next morning.

**Test safely first:** run `remindersDryRun` or `dailyDryRun` from the editor. They only write to the Execution log (who WOULD get an email) and send nothing. A row whose name is `TEST`, or whose email ends in `.invalid` / `@example.com`, is saved but never emailed.

**Tuning (top of `apps-script.gs`):** `DAILY_HOUR` (9), `TZ` (`Asia/Dubai`), `SPRINT_DAYS` (30), `REMINDER_AFTER_HOURS` (24), `REMINDER_MAX_AGE_DAYS` (7), `QUOTA_RESERVE` (10), `SKIP_EMAIL_IF_WHATSAPP_ONLY` (false). After changing `DAILY_HOUR` or `TZ`, run `setupTriggers` again. To switch automation off: run `removeTriggers`.

**Honest limits of the automation:**
- Google's daily trigger runs **sometime within the 9:00 to 10:00 hour** (Dubai), not at exactly 9:00:00. The hourly reminder also drifts within its hour.
- If a daily run fails or the quota is used up, that day is **not retried** automatically (the sheet says `CONTENT_NOT_SENT`). Run `sendDailyContent` by hand from the editor if that happens: it never double-sends the same day.
- The script sends **email only**. It cannot send WhatsApp. If someone chose WhatsApp only, either send by hand or set `SKIP_EMAIL_IF_WHATSAPP_ONLY = true` (their status then says `WHATSAPP_MANUAL`).
- Free Gmail limit is about **100 emails a day** (Workspace about 1,500). Each person costs one email a day, plus confirmations and your own notices, so about 70 active members is the practical ceiling on a free account. The script stops at `QUOTA_RESERVE` left and logs it.
- The page cannot verify payments, so `PAID` is always your decision. `PAID_CLAIMED` does not start the daily emails.
- Automatic emails can land in spam or Promotions, especially the first ones.
- Time-driven triggers need the one-time authorization above. If you change the Google account that owns the script, run `setupTriggers` again.

## What you see in the sheet (`Responses` tab)
One row per signup, with every questionnaire answer, plus:
- **Payment status:** `PENDING` (signed up) > `LINK_OPENED` (opened the payment link) > `PAID_CLAIMED` (said they paid. Not verified) > **`PAID`** (you set this by hand after you see the payment in Razorpay).
- **Delivery status:** the latest thing the script did for that row: `CONFIRMATION_SENT`, `WELCOME_SENT`, `REMINDER_SENT`, `CONTENT_SENT day 5/30`, `SPRINT_COMPLETE (30 days)`, `TEST_ROW_NO_EMAIL`, or a problem such as `EMAIL_FAILED: ...`, `REMINDER_FAILED: ...`, `CONTENT_NOT_SENT: daily quota used up`.
- **Owner notified:** `OWNER_EMAIL_SENT` when you got the "new signup" / "payment claimed" email.
- **Reminder sent:** the time the one payment reminder went out (`FAILED ...` if Google rejected the address). Empty = not sent yet.
- **Content start date / Last content sent / Content days sent:** filled by the script (see step 4). You may type a start date (`2026-10-20`) to start someone on a chosen day.
- **Delivery channel:** where they want the daily content (WhatsApp, Email or Both). The script can only send email. WhatsApp content you send by hand.

New columns are added automatically at the end of an existing sheet when the new script runs. Do not rename the header cells in row 1: the script finds columns by their names (you may reorder or hide columns).

## Limits to know
- Emails go out from the Gmail account that owns the sheet. Free Google accounts can send about 100 emails a day. Emails may land in spam at first: tell the first few people to check.
- The page posts in `no-cors` mode, so it cannot read the script's reply. Check the sheet's Delivery status to see what actually happened.
- The sheet holds names, emails and WhatsApp numbers. Keep it private and only contact people who ticked the consent box.
- No trackers or analytics. The only outside request is the Google Fonts stylesheet; the page works without it.
- Step 3 shows one static post and one swipeable carousel (Instagram 4:5 style), built from the person's answers, plus three locked teasers (market insight, community insight, buyer / seller psychology). They use `[bracketed]` placeholders, never invented figures. Keep it that way if you edit the text in `buildExamples()` in `index.html`.

## Logo, popup and poster
- `logo.png` (optimised, white background) is the header logo, favicon source and email logo. Emails load it from `PAGE_URL + 'logo.png'` (top of `apps-script.gs`), so keep the file in the repo.
- When the page opens, a poster popup explains the offer. It shows once per browser session. Its prices follow `PRICE`, `REGULAR_PRICE` and `SPOTS` in `index.html`. People with "reduce motion" turned on see a still poster.
- `poster.png` (1080 × 1620) is the same poster as a shareable image for WhatsApp and Pinterest. `poster-share.jpg` is the link preview image. If you change prices, re-export the poster.

## Hosting
Already published with GitHub Pages (branch `main`, folder `/`). To update the page: edit `index.html` and push to `main`; the live page refreshes in about a minute. `.nojekyll` is included.
