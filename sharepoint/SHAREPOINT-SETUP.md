# Sending entries to SharePoint Excel

When someone taps **Save & post to Teams**, the same Teams workflow also:

1. Opens that person's workbook: `<Audit folder>/<Name> Notes/3_Workstream_Template_<Name>.xlsx`
2. Adds a **new tab** for the application, copied from your blank "Additive Application" form, with every row filled in and the part photo in the Image row
3. Adds a line to that workbook's **Summary** tab, with a link to the new tab
4. Adds a line to the shared **Audit Summary.xlsx** in the audit folder, with a link to the person's workbook

Posting the same entry again from the app (**Save & post to Teams again**) updates its tab and rows instead of adding duplicates.

The **audit folder** is a setting inside the workflow. Only the workflow's owner (the admin) can see or change it. Phones never need updating when a new audit starts.

---

## 1. Before each audit: folders and files

In the audit folder in SharePoint:

- Each person has a folder named **`<Name> Notes`** containing **`3_Workstream_Template_<Name>.xlsx`**. `<Name>` must match the **Notes folder name** in that person's app settings exactly, for example `Miko` → `Miko Notes/3_Workstream_Template_Miko.xlsx`.
- That workbook has a sheet called **Workstream Example** with a blank "Additive Application" form in **A1:B14**. The script copies that block, including labels, colors and column widths, to make each new tab.
  - *Optional:* add a sheet named **Form Template** with one clean blank form in A1:B14. If it exists, the script uses it instead, so filling in Workstream Example by hand can never affect new tabs.
- Create the shared summary workbook: **New → Excel workbook**, then rename it exactly **`Audit Summary.xlsx`**. Leave it blank. The script adds the headers the first time it runs.

On each phone: **⚙︎ → Notes folder name** (for example `Brenden`), then **Save**.

---

## 2. Add the two Office Scripts (admin, one time)

1. Open any workbook in **Excel for the web**, for example your own Workstream workbook.
2. Go to **Automate** → **New Script**.
3. Delete the sample code and paste the full contents of **`Add Additive Application.ts`**.
4. Click the script name at the top, rename it to exactly **Add Additive Application**, then **Save script**.
5. Repeat steps 2–4 with **`Add to Audit Summary.ts`**, named exactly **Add to Audit Summary**.

The scripts are saved in your own OneDrive (*Documents → Office Scripts*), so only you can change them.

---

## 3. Add the Excel steps to the Teams workflow (admin, one time)

Open the **Workflows** app in Teams, open your webhook workflow, and click **Edit**.

### 3a. The admin settings (two variables)

Click **+** directly below **Initialize variable (Attachments)** → **Add an action** → search **Initialize variable**. Add two of them:

| Name | Type | Value |
|---|---|---|
| `AuditFolder` | String | The audit folder's path inside the Documents library, starting with `/`, for example `/Site Folder/Customer/Audits/Plant/Audit 12012026-12022026` (see the tip in 3c) |
| `LibraryUrl` | String | The web address of the Documents library, for example `https://yourcompany.sharepoint.com/sites/YourSite/Shared Documents`. Open the library in your browser and copy the address up to and including `Shared Documents` (or `Shared%20Documents`). Drop everything after it, such as `/Forms/AllItems.aspx`. |

> Variables have to sit at the top level of the workflow, not inside a condition or loop. Directly under the existing *Initialize variable* steps is the right spot.

### 3b. Only run the Excel steps for real entries

1. Click **+** **below** the big **Attachments is null** box (outside it, at the bottom) → **Add an action** → **Condition**. Rename it **Entry included**.
2. Left box: click **fx** and enter `empty(triggerBody()?['entry'])`. Operator: **is equal to**. Right box: **fx** → `false`.
3. In the condition's **Settings** tab, under **Run after**, tick **Is successful**, **Has failed** and **Is skipped**. That way Excel still gets updated even if the Teams post fails.

Test messages from the app or PowerShell have no entry data, so they skip the Excel steps.

### 3c. Inside **True**: update the person's workbook

1. **Compose**, renamed **Workbook path**. Inputs: **fx** →
   ```
   concat(variables('AuditFolder'), '/', triggerBody()?['entry']?['notesFolder'], ' Notes/3_Workstream_Template_', triggerBody()?['entry']?['notesFolder'], '.xlsx')
   ```
2. **Excel Online (Business) → Run script**, renamed **Add to person workbook**:
   - **Location:** your SharePoint site · **Document Library:** Documents
   - **File:** select the **Outputs** of *Workbook path* from the dynamic content list (⚡ icon)
   - **Script:** Add Additive Application
   - **entryJson:** **fx** → `string(triggerBody()?['entry'])`

   > **Tip for finding the AuditFolder path:** before choosing dynamic content for **File**, click the folder icon and browse to your own workbook once. The box fills in the exact path, such as `/…/Audit 12012026-12022026/Miko Notes/3_Workstream_Template_Miko.xlsx`. Copy everything **before** `/Miko Notes` into the **AuditFolder** variable, then switch **File** back to the *Workbook path* output.

### 3d. Inside **True**: update the shared Audit Summary

3. **Compose**, renamed **Workbook link**. Inputs: **fx** →
   ```
   concat(variables('LibraryUrl'), replace(outputs('Workbook_path'), ' ', '%20'))
   ```
4. **Excel Online (Business) → Run script**, renamed **Add to audit summary**:
   - **Location / Document Library:** same as above
   - **File:** **fx** → `concat(variables('AuditFolder'), '/Audit Summary.xlsx')`
   - **Script:** Add to Audit Summary
   - **entryJson:** **fx** → `string(triggerBody()?['entry'])`
   - **tabName:** dynamic content → **result** from *Add to person workbook*
   - **workbookLink:** dynamic content → **Outputs** of *Workbook link*

### 3e. Inside **True**: tell the admin if something went wrong

5. **Microsoft Teams → Post message in a chat or channel**, renamed **Excel failed alert**:
   - **Post as:** Flow bot · **Post in:** Chat with Flow bot · **Recipient:** your email
   - **Message:** **fx** →
     ```
     concat('⚠️ Couldn''t add "', triggerBody()?['entry']?['rows']?['appName'], '" from ', triggerBody()?['entry']?['originator'], ' to Excel. Check that the folder "', triggerBody()?['entry']?['notesFolder'], ' Notes" and its workbook are in the current audit folder, and that Audit Summary.xlsx is there too.')
     ```
   - **Settings → Run after:** untick **Is successful**, tick **Has failed**, **Is skipped** and **Has timed out**

Click **Save**.

The first time, the Excel steps ask you to sign in. The workflow then writes to the workbooks using **your** SharePoint access, so you need edit permission on every Notes folder and on Audit Summary.xlsx. SharePoint will show you as *Modified By*.

---

## 4. Admin-only control

- **Changing the audit folder:** only the workflow's owners can open or edit it. Keep yourself as the **only owner**. In Power Automate (make.powerautomate.com) → **My flows** → this flow → **Owners**, it should list only you. Don't add co-owners.
- **Everyone else** only has the webhook URL in their app. It can send entries, but it can't see or change the workflow, its settings, or the scripts.
- **The scripts** live in your OneDrive, so only you can edit them.

### Starting a new audit

1. Create the new audit folder with each person's `<Name> Notes` folder and workbook, plus a blank `Audit Summary.xlsx`.
2. In **Workflows → your workflow → Edit**, change the **AuditFolder** value, then **Save**.

That's all. Phones keep working with no changes.

---

## 5. Test it

1. Post a real entry from the app. Within about 30 seconds you should see:
   - the Teams card
   - a new tab, such as `01 Coolant nozzle bracket`, in your workbook
   - a **Summary** tab at the front of your workbook
   - a row in **Audit Summary.xlsx**
2. Open the entry in the app and tap **Save & post to Teams again**. The same tab and rows update, with no duplicates.

**If it doesn't work**, open the run in **Workflows → run history**:
- **Add to person workbook failed with "file not found":** check the *Workbook path* output. The Notes folder name on the phone or the AuditFolder path probably doesn't match SharePoint exactly, including spaces and capital letters.
- **A script error:** take a screenshot of the error and the step's inputs.
