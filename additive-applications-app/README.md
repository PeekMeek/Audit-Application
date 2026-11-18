# Additive Applications (trial)

A phone app (PWA) for logging additive manufacturing application ideas:

1. Check your notebook page against the **Required** and **Supporting** lists shown in the app.
2. Take a photo of the page, plus an optional photo of the part.
3. The AI reads your handwriting and fills in each field by its position. Bullets or numbers both work, and you don't need to write the labels.
4. Review the fields. The urgency slider runs from green (1) to red (10), and confidence runs from **!** to **!?!?!?!?**. Fix any lines that landed in the wrong slot with the ↓ / ↑ buttons.
5. **Save & post to Teams.** The message starts with your name.
6. **Export workbook** gives you one `.xlsx` file with a **Summary** tab plus one tab per application in the "Additive Application" form layout, with the part photo in the Image row.

Entries are saved on the phone, so the export always includes everything you've logged on that device.

---

## 1. Put it on GitHub Pages

1. On GitHub, create a new repository, for example `additive-applications`. **Public** is required for free GitHub Pages. Nothing secret is in these files.
2. Upload every file and folder in this zip, keeping the folders (`icons/`, `vendor/`). Use **Add file → Upload files** and drag in the contents of the folder, not the folder itself.
3. Go to **Settings → Pages → Build and deployment**, set the source to **Deploy from a branch**, pick `main` and `/ (root)`, then **Save**.
4. After a minute the app is live at `https://<your-username>.github.io/additive-applications/`.

**Updating later:** upload the changed files, then change `VERSION` at the top of `sw.js` (for example `aa-v1` → `aa-v2`) so phones pick up the new version. Close and reopen the app twice if it still shows the old version.

## 2. Install it on your phone

- **iPhone (Safari):** open the link → Share → **Add to Home Screen**.
- **Android (Chrome):** open the link → ⋮ menu → **Install app** (or **Add to Home screen**).

The first time it opens, it asks for your name, the AI key, and the Teams link. These are stored **only on that phone**. Each person who uses the app sets up their own phone with their own name.

## 3. Handwriting reader (AI key)

Pick one in the app's settings:

| Service | What you need |
|---|---|
| **Anthropic (Claude)** (default) | An API key from console.anthropic.com. The model is pre-filled (`claude-sonnet-5-5`). |
| **OpenAI** | An API key from platform.openai.com and the name of a current vision-capable model. |
| **Azure OpenAI** | Your company resource's endpoint, deployment name, and key. Ask IT. This is the likely route for the company version. |
| **None** | Skips the AI. You type entries in by hand. Excel and Teams still work. |

Each page costs roughly a cent or two to read. Photos go straight from the phone to the service you pick.

> ⚠️ For this trial the key is stored on the phone. That's fine for a few test users, but before rolling out company-wide, the key should sit on a small company server (or use Azure OpenAI with company sign-in) so it isn't on every phone.

## 4. Teams group chat (Workflows)

1. In Teams, open the **group chat** the messages should go to.
2. Click **⋯** (More options) at the top of the chat → **Workflows**. You can also open the **Workflows** app from the Teams sidebar and search the templates.
3. Choose the template **"Post to a chat when a webhook request is received"**.
4. Name it, for example *Additive Applications*, sign in if asked, and confirm the chat it posts to. Click **Add workflow**.
5. Copy the **URL** it shows you.
6. In the app: ⚙︎ → **Workflow webhook URL** → paste → **Send a test message**. A card should appear in the chat within a few seconds.

**Moving to a channel later:** make a new workflow from **"Post to a channel when a webhook request is received"** on that channel, then paste its URL into each phone's settings. Nothing else changes.

**Notes**
- Messages arrive from the **Workflows** bot, so each card starts with **"*Your Name* logged a new additive application"** to show who sent it.
- Treat the webhook URL like a password. Anyone who has it can post into the chat. It's stored in each phone's settings, never in these files.
- Some companies turn off Workflows or block outside apps from posting. If the test message never shows up, check with IT.
- **Part photo in Teams** is off by default. Teams caps message size, and an over-limit post fails without an error. Turn it on in Settings to include a small thumbnail.

## How the notebook is read

- Write the items **in the order shown in the app**. Bullets (•, -, *) and numbers (1., 1)) both work, and so does a mix.
- **Numbered lines are safest.** The AI places each value by its number, so a skipped item can't shift the rest.
- With bullets, keep a bullet with **"—" or "N/A"** for anything you leave blank, so the count stays right. For **Photo**, a check mark is fine.
- Writing **"Required"** and **"Supporting"** (or **R** / **S**) as headings restarts the count, which keeps any mix-up inside one section.
- Lines without a bullet are added to the item above, so long descriptions are fine.
- If a value doesn't fit its field (words where 1–10 should be, a cost with no number), the app flags it and suggests a skipped line. Use **↓** on the first misplaced field to push it and everything below down one slot, or **↑** to remove an extra line. A value removed with ↑ goes into *Other notes*.

## How notebook items map to the Excel form

| Excel row | From the notebook |
|---|---|
| Originator | Your name (from Settings) |
| Department/Area/Product Line | Department/area/product line |
| Workstream Owner | Default from Settings, editable per entry |
| Application Name / Description | Same |
| Current Issue/Challenge | Current issue + Downtime impact + EOL concern + Safety concern |
| Part Info | Part size + Current material/method + Current part cost + Lead time |
| Notes for Additive Solution | Suggested AM material/process |
| Customer Usage | EAU (Estimated Annual Usage) |
| Urgency/Importance (1-10) | Slider. The cell is shaded green → red. |
| Confidence (1-10) | Slider, for example `7  !?!?!?` |
| Customer input/interest | Customer buy-in/interest + Customer quote/input |
| Image | Part photo |

Below the form, each tab also has a **Notebook Details** section with every item exactly as captured, plus the notebook page photo, so you can always check the original.

## Files

| File | What it does |
|---|---|
| `index.html`, `styles.css` | Screens and look |
| `app.js` | App logic: photos, AI call, review, saving, Teams |
| `fields.js` | The field list, the AI instructions, skip checks, and the Teams card |
| `xlsx.js` | Builds the Excel workbook on the phone |
| `vendor/jszip.min.js` | Zip library used to write `.xlsx` files (MIT licence) |
| `sw.js`, `manifest.webmanifest`, `icons/` | Makes it installable and lets it open offline |

To change the field list or the Excel row labels, edit the lists at the top of `fields.js`.

## Planned next

- **Part dimensions from a photo:** photograph the part next to a ruler and have the app estimate its size from the ruler's scale.
- **Company version:** Microsoft sign-in (so names fill in automatically), the AI key moved to a company server or Azure OpenAI, and posting to a Teams channel.
