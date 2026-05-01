# NinjaOne Live-Feed Webhook Setup (PowerShell)

This script configures the kind of NinjaOne live-feed webhook described in NinjaOne's [Send NinjaOne Webhooks Via SIEM](https://www.ninjaone.com/blog/ninjaone-com-blog-send-ninjaone-webhooks-via-siem/) blog post — but without you ever having to open NinjaOne's API documentation page.

It uses the OAuth 2.0 **Authorization Code** flow (the same flow the API docs page uses behind the scenes), so you don't share your login credentials with anything and you don't need a separate API client tool.

If you've never touched OAuth or PowerShell before, this guide is written for you.

---

## What the script does

1. Opens NinjaOne in your default browser and asks you to log in / consent.
2. Catches the authorization code Ninja sends back to `https://localhost`.
3. Exchanges that code for an access token.
4. Calls `PUT /v2/webhook` to create (or replace) the webhook that streams Ninja activity events to your SIEM.

You only need to run it once. After that, the webhook keeps streaming on its own.

---

## Part 1 — One-time setup in NinjaOne

You need a NinjaOne **System Administrator** account for everything in this section.

### 1.1 Create the OAuth Client App

1. In NinjaOne, go to **Administration → Apps → API**.
2. Open the **Client App IDs** tab and click **Add**.
3. Fill in the form like this:

   | Field                    | Value                                                            |
   | ------------------------ | ---------------------------------------------------------------- |
   | **Application Platform** | `Web`                                                            |
   | **Name**                 | `Webhook Setup Script` (or anything you'll recognize)            |
   | **Redirect URIs**        | `https://localhost`  *(no port, no trailing slash, exactly this)*|
   | **Scopes**               | `Monitoring`  *(also tick `offline_access` if you want refresh tokens)* |
   | **Allowed Grant Types**  | `Authorization Code`, `Refresh Token`                            |

4. Click **Save**.
5. Ninja will show you a **Client ID** and a **Client Secret**. Copy both somewhere safe **right now** — the secret is shown only once.

> **Heads up about the redirect URI:** It must match `https://localhost` byte-for-byte on both ends — what's saved on the OAuth app and what's in the script's `$Config.RedirectUri`. No port number, no trailing slash. Mismatch = `invalid_client` error every time.

### 1.2 Find your region's login URL

Look at the URL you log into NinjaOne with. That's your `BaseUrl`:

| Region          | Login URL                       |
| --------------- | ------------------------------- |
| United States   | `https://app.ninjarmm.com`      |
| United States 2 | `https://us2.ninjarmm.com`      |
| Canada          | `https://ca.ninjarmm.com`       |
| EU              | `https://eu.ninjarmm.com`       |
| Oceania (APAC)  | `https://oc.ninjarmm.com`       |

### 1.3 Get your SIEM ingest URL and token

For **Splunk HEC** (the example in the blog post):

- Create the HEC in Splunk and note the URL. The script's payload assumes it ends in `/services/collector/raw` — Ninja's outbound POST won't be parsed correctly otherwise.
- Copy the HEC token Splunk generated. The auth header value goes in the form `Splunk <token>`.

For **most other SIEMs** (Sentinel via Logstash, Sumo, Elastic, custom collector, etc.) the auth header is usually `Bearer <token>` instead of `Splunk <token>`. Adjust the `value` field of the header accordingly.

---

## Part 2 — Edit the script

Open `Set-NinjaWebhook.ps1` in any text editor. There are two blocks at the top — that's all you need to touch.

### Block 1: `$Config` (credentials)

```powershell
$Config = [ordered]@{
    BaseUrl       = 'https://<your Login URL>'
    TokenEndpoint = 'https://<your login URL>/ws/oauth/token'
    AuthEndpoint  = 'https://<your login URL>/ws/oauth/authorize'
    ClientId      = '<Your Client ID>'
    ClientSecret  = '<Your Client Secert>'
    RedirectUri   = 'https://localhost'
    Scope         = 'monitoring offline_access'
}
```

Replace **all five** `<...>` placeholders. Use the same host across `BaseUrl`, `TokenEndpoint`, and `AuthEndpoint`. Example for a US instance:

```powershell
BaseUrl       = 'https://app.ninjarmm.com'
TokenEndpoint = 'https://app.ninjarmm.com/ws/oauth/token'
AuthEndpoint  = 'https://app.ninjarmm.com/ws/oauth/authorize'
```

Leave `RedirectUri` exactly as `https://localhost`. Don't add a port. Don't add a slash.

### Block 2: `$WebhookPayload` (what to stream)

```powershell
$WebhookPayload = [ordered]@{
    url     = 'https://<your-siem-host>:8088/services/collector/raw'
    expand  = @('device', 'organization', 'location')
    headers = @(
        [ordered]@{
            name  = 'Authorization'
            value = 'Splunk <your-HEC-token-here>'
        }
    )
    activities = [ordered]@{
        CONDITION = @('*')
        SYSTEM    = @('*')
        ACTIONSET = @('*')
    }
}
```

Three things to change:

- **`url`** — your SIEM ingest URL. For Splunk HEC, must end in `/services/collector/raw`.
- **`headers[0].value`** — `Splunk <token>` for Splunk HEC, or `Bearer <token>` for most other SIEMs.
- **`activities`** — which Ninja activity categories you want streamed. Each value is `@('*')` meaning "all sub-types of this category" (per the v2 webhook schema). Common categories:
  - `CONDITION` — alerts and resets
  - `ACTIONSET` — automations and technician actions
  - `SYSTEM` — system events
  - `ANTIVIRUS`, `MDM`, `PATCH_MANAGEMENT`, `SOFTWARE_PATCH_MANAGEMENT`, `TICKETING`, `SECURITY`, `MONITOR`, `SCHEDULED_TASK`, `REMOTE_TOOLS`, `SCRIPTING`

Keep only what your SIEM actually needs — fewer categories means less ingest noise (and potentially less Splunk license burn).

Save the file.

---

## Part 3 — Run it

1. Open **PowerShell**. Either Windows PowerShell 5.1 (built into Windows) or PowerShell 7+ works.
2. Navigate to the folder containing the script:

   ```powershell
   cd C:\path\to\ninja-webhook
   ```

3. **First-time-running-a-script-on-this-machine?** Allow scripts for just this session (this doesn't change anything permanently):

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```

4. Run it:

   ```powershell
   .\Set-NinjaWebhook.ps1
   ```

### What you'll see, step by step

**A. A browser window opens** to NinjaOne's consent page. Log in if asked, then click **Allow**.

**B. The browser redirects to `https://localhost/?code=...`** and shows an error like *"This site can't be reached"*, *"localhost refused to connect"*, or a certificate warning.

> **This is expected and correct.** Nothing is actually listening on `https://localhost` — Ninja just delivers the auth code to that address so your browser can show it to you. The error page IS the success state.

**C. Copy the entire URL out of the browser's address bar.** It will look like this:

```
https://localhost/?code=eyJhbGciOiJIUzI1NiJ9.long.string&state=abc123def456
```

**D. Switch back to PowerShell and paste it** at the prompt that says `Paste the full https://localhost/?code=... URL here`. Hit Enter.

**E. The script will then:**
- Extract the `code` parameter
- Exchange it at `/ws/oauth/token` for an access token
- PUT the webhook config to `/v2/webhook`
- Print `[ OK ]  204 No Content - webhook configured successfully.`

If anything fails, the script prints the actual response body Ninja returned along with a list of the most likely causes.

---

## Part 4 — Verify it worked

In NinjaOne:

1. Go to **Administration → Notification Channels**.
2. You should see a webhook entry **with no name**. That's your live-feed webhook.

   > **DO NOT rename, edit, or modify it.** Touching that entry breaks the stream. This is per NinjaOne's own SIEM blog post.

3. Trigger a test condition (e.g., manually fire an alert on a test device) and confirm it lands in your SIEM within a few seconds.

**Optional but recommended:** Under **General → Activities**, enable notifications for the *webhook failed* and *webhook disabled* events. That way you'll find out fast if the stream stops for any reason.

---

## Troubleshooting

| Error / symptom                                        | What it means                                                                                                                |
| ------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------- |
| Script exits with `Configuration is incomplete`        | One of the `<placeholder>` values is still in `$Config` or `$WebhookPayload`. Open the script and finish editing.            |
| Browser shows "site can't be reached" on `https://localhost` | Expected. Don't close the tab — just copy the URL from the address bar back into PowerShell.                           |
| `No 'code' parameter found in the URL`                 | You pasted the wrong URL. It must be the one your browser was redirected to (starts `https://localhost/?code=...`), not the original Ninja consent URL. |
| `invalid_grant`                                        | The auth code was already used, or it expired. Auth codes are single-use and live ~60 seconds. Just re-run the script.       |
| `invalid_client`                                       | Wrong Client ID or Client Secret, **or** the Redirect URI on the Ninja OAuth app doesn't exactly match `https://localhost`. Check both.  |
| `unauthorized_client`                                  | The OAuth app doesn't have `Authorization Code` ticked under Allowed Grant Types. Edit the app in Ninja and tick it.         |
| `401 Unauthorized` on the webhook PUT                  | Token invalid or expired between steps. Just re-run the script.                                                              |
| `403 Forbidden` on the webhook PUT                     | Either your Ninja user is not a System Administrator, or the OAuth app's scopes don't include `monitoring`.                  |
| `400 Bad Request` on the webhook PUT                   | Something in the payload is off. Check the `url` format, activity type spelling, and that `headers` is an array of `{name, value}` objects. |
| Webhook configured OK but no events show up in SIEM    | (a) The activity types you picked aren't firing in your tenant. Try adding `CONDITION` and trigger a test alert. (b) Splunk HEC URL doesn't end in `/services/collector/raw`. (c) HEC token wrong or HEC disabled in Splunk. |

---

## Updating the webhook later

Want to add or remove activity types, swap SIEM tokens, or change the URL?

1. Edit the `$WebhookPayload` block in the script.
2. Re-run `.\Set-NinjaWebhook.ps1`.

`PUT /v2/webhook` **replaces** the existing config — it doesn't create duplicates. You'll get a fresh consent prompt; that's normal.

---

## Removing the webhook

There isn't a `DELETE` analog called out in the v2 webhook schema. To stop the stream, the cleanest path is:

- In NinjaOne, go to **Administration → Notification Channels**, find the unnamed webhook entry, and delete it.

That's the same entry the script created and that the blog post warns not to modify under normal operation.

---

## Files in this folder

- `Set-NinjaWebhook.ps1` — the script
- `README.md` — this walkthrough
