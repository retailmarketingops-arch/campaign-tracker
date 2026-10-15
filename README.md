# Campaign Tracker

An app for retail marketing element plans. For each campaign you upload one store-level Excel or CSV file. The app then shows, for every zone:

- total stores, covered stores and total element quantity
- the quantity of each element and how many stores receive it
- breakdowns by Xi State, Retailer Type and Category, or any other column you choose
- an **Export to Excel** button for the zone report

Campaigns stay saved on the device, so you can open any campaign's report at any time. The same code runs as an **Android APK** and as a **web app** that you can install on a desktop or phone.

---

## Easiest way to build: GitHub (free, nothing to install)

GitHub builds the APK for you on its servers. It can also host the web app at a free web address for desktop users.

### 1. Put the code on GitHub
1. Create a free account at https://github.com.
2. Click **+ → New repository**. Name it `campaign-tracker`, choose **Public** or **Private**, then click **Create repository**.
3. On the new repository page, click **uploading an existing file**.
4. Unzip `campaign-tracker.zip` on your computer. Open the `campaign-tracker` folder, select **everything inside it**, including the `.github` folder, and drag it into the browser.
   - On Mac, press `Cmd + Shift + .` in Finder to see hidden folders such as `.github`. On Windows, use **View → Show → Hidden items**.
5. Click **Commit changes**.

### 2. Get the APK
1. In the repository, open the **Actions** tab. If GitHub asks, click **I understand… enable workflows**.
2. The **Build Android APK** run starts by itself after each upload. You can also start it by clicking **Build Android APK → Run workflow**.
3. Wait about 5–8 minutes for the green tick, then open the run.
4. Under **Artifacts**, download **CampaignTracker-apk**. It is a zip file that contains `app-debug.apk`.

### 3. Install on Android phones
1. Send `app-debug.apk` to the phone by WhatsApp, email, Google Drive or a USB cable.
2. Tap the file. Android asks you to allow installs from that app (for example Files or WhatsApp). Allow it, then tap **Install**.
3. If Play Protect warns that the app is from an unknown developer, choose **Install anyway**. This warning appears for any APK that doesn't come from the Play Store.

### 4. Desktop / iPhone: the web app (optional, also free)
1. In the repository, go to **Settings → Pages** and set **Source = GitHub Actions**.
2. Open **Actions → Publish web app (GitHub Pages) → Run workflow**.
3. Your app is now online at `https://<your-github-name>.github.io/campaign-tracker/`.
   - **Windows / Mac (Chrome or Edge):** open that link, then click the **Install** icon in the address bar. The app appears like a normal desktop program and also works offline.
   - **iPhone:** open the link in Safari, tap **Share → Add to Home Screen**. iPhones can't install APKs, so this is the way to use it there.

> If the repository is private, GitHub Pages needs a paid plan. Make the repository public, or use another free host such as Netlify (drag the `www` folder onto https://app.netlify.com/drop).

---

## Other free ways to build the APK

| Option | Best for | How |
|---|---|---|
| **Android Studio** (free, Windows/Mac/Linux) | Someone comfortable installing software | Install Node.js 22+ and Android Studio. In the project folder run `npm install`, `npx cap add android`, `npx @capacitor/assets generate --android` and `npx cap sync android`, then `npx cap open android`. In Android Studio, choose **Build → Build App Bundle(s) / APK(s) → Build APK(s)**. |
| **PWABuilder** (https://www.pwabuilder.com, free) | When the web app is already online (step 4 above) | Paste the GitHub Pages link and choose **Package for stores → Android**. You get an APK or AAB, and you can also make a Windows package. |

---

## Using the app

1. Tap **+ New campaign**, enter the campaign name and choose the data file.
2. Check the columns. The app detects them automatically:
   - **Store code** is used to count stores. Rows without a store code, such as the totals row at the top of your file, are skipped.
   - **Zone** gives you the zone buttons.
   - **Report by**: Xi State, Ret Type and Category are selected by default. Tick any other columns you want as well.
   - **Elements**: every quantity column, such as Poster, Dangler or Standee.
3. Tap **Save campaign**. Open any saved campaign from the home screen.
4. In a campaign:
   - Pick a zone, then switch between the **Xi State**, **Ret Type** and **Category** tabs.
   - **Qty (stores)** means element quantity, with the number of stores receiving it in brackets.
   - **Export to Excel** saves the current zone's report.
   - **Replace file** loads an updated plan for the same campaign. **Columns** changes the detected columns.
5. **Backup & sharing** (home screen): **Save backup file** creates one `.json` file that holds every campaign. Send it to a colleague, who taps **Import backup file** to get the same campaigns.

Data is stored only on the device, never on a server. Uninstalling the app or clearing its data removes the campaigns, so save a backup first.

## File format the app expects

- One row per store, with a header row. The header can sit lower than the first row, because the app finds it automatically.
- Text columns for zone, state, retailer type, category and so on.
- One numeric column per element, holding the quantity for that store. A blank cell counts as 0.

## Project layout

```
www/                  the app itself (HTML/CSS/JS). Open www/index.html to try it in a browser
  vendor/             SheetJS spreadsheet reader (Apache-2.0)
package.json          Capacitor (wraps the web app as an Android app)
capacitor.config.json app id / name
assets/               app icon source used for the Android icons
.github/workflows/    free GitHub builds: APK + web hosting
```

To change the app name or ID before the first build, edit `capacitor.config.json`. To change the icon, replace the PNGs in `assets/` (1024×1024).

### Releasing on the Play Store (optional)
The APK built above is a *debug* build. It is fine for sharing within your team, but the Play Store needs a *signed release* build. In Android Studio, use **Build → Generate Signed Bundle / APK**. A Google Play developer account costs a one-time US$25.
