# MatFinance 📈

MatFinance is a lightweight, mobile-friendly personal finance tracker. It uses a local-first web interface (PWA) that syncs your income and expenses directly to a private Google Sheet.

It features offline support, background syncing, and beautiful charts, all without needing a traditional database or backend server.

## Features

- 📱 PWA Ready: Install it on your iOS or Android home screen for an app-like experience.

- ⚡ Local-First & Offline Support: Add transactions even when you don't have an internet connection. They will queue up and sync automatically when you go back online.

- 📊 Automated Reports: Visualizes your monthly spending with Chart.js.

- 🔒 Private & Secure: Your data never touches a third-party server. It goes straight from your device to your personal Google Sheet.

## 🛠️ How to Setup Your Own Instance

You can host your own version of MatFinance completely for free! Here is how to set it up:

### Step 1: Set up the Google Sheet (Backend)

1. Create a new, blank Google Sheet.

2. Go to Extensions > Apps Script.

3. Create two files in the Apps Script editor:

    - `Code.gs` (Delete the default code and paste the contents of `Code.gs` from this repo).

    - `Index.html` (Create an HTML file and paste the contents of `Index.html` from this repo).

4. Save the project.

5. In the dropdown at the top, select the `setupSpreadsheet` function and click Run. Grant the necessary permissions. This will automatically format your Google Sheet.

### Step 2: Deploy the Web App

1. In the Apps Script editor, click Deploy > New deployment.

2. Select Web app as the type.

3. Set Execute as to `Me`.

4. Set Who has access to `Anyone`.

5. Click Deploy and copy the Web app URL (it ends in `/exec`).

### Step 3: Host the Frontend

1. Open the `app.html` file from this repository.

2. Find the following line:
```
const GAS_APP_URL = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
```

3. Replace the placeholder URL with the actual Web app URL you copied in **Step 2**.

4. Host `app.html` anywhere you like! You can use GitHub Pages, Cloudflare Pages, Netlify, or Vercel for free.

5. Open your hosted `app.html` URL on your phone and tap "Add to Home Screen".

## Contributing

Feel free to fork this repository and submit pull requests if you have ideas for new features or improvements!
Happy Coding and remember `FOSS` is the best.

## License

MIT License