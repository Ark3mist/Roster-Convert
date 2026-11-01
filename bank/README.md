# Roster Bank setup (your Google Drive)

When someone ticks **Share to Roster Bank** and presses **Download PDF**, a copy of that PDF is saved to a private folder in **your** Google Drive:

```
My Drive / Roster Bank / Shafiq Fakhri Othman, 2106537 / Shafiq Fakhri Othman, 2106537, November 2026.pdf
```

Only you can see this folder. People who share can upload, but they can't list, open, or delete anything in it. Sharing the same month again replaces the older copy.

## One-time setup (about 10 minutes)

1. Go to **https://script.google.com** while signed in to the Google account whose Drive should hold the rosters. Click **New project**.
2. Delete the sample code, paste in the whole of [`Code.gs`](Code.gs), and click **Save**. Name the project "Roster Bank".
3. Click **Deploy → New deployment**. Click the ⚙ next to "Select type" and choose **Web app**, then set:
   - **Execute as:** *Me*
   - **Who has access:** *Anyone*. This is needed so other crew can send files. They still can't read anything.
4. Click **Deploy**. Google asks you to **Authorize access**: choose your account, then **Advanced → Go to Roster Bank (unsafe) → Allow**. The warning appears because the script is your own and hasn't been reviewed by Google. It only gets access to your Drive.
5. Copy the **Web app URL**. It looks like `https://script.google.com/macros/s/AKfy…/exec`.
6. In this repo, open **`config.js`**, paste the URL into `bankUrl`, and put your name in `bankOwner`:
   ```js
   window.ROSTER_CONFIG = {
     bankUrl: 'https://script.google.com/macros/s/AKfy…/exec',
     bankOwner: 'Shafiq',
   };
   ```
   Commit the change. You can edit the file directly on github.com with the ✏️ button.
7. Check it works: open the Web app URL in a browser. You should see `{"ok":true,"service":"Roster Bank"}`.

## Good to know

- **Updating the script:** after changing `Code.gs`, go to **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. This keeps the same URL.
- **Turning it off:** set `bankUrl: ''` in `config.js`. The option then disappears from the app. To stop uploads immediately, also archive the deployment in Apps Script.
- **Spam:** the URL is visible to anyone who reads the page source. The script only accepts real PDFs up to 5 MB, and at most 60 uploads an hour.
- **Privacy:** rosters contain names, staff numbers, duties, and hotels. Only keep what people chose to share, and delete a roster if its owner asks.
