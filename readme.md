# PowerPoint Text CRUD Add-in

A simple, modern PowerPoint Add-in that lets you **Create, Read, Update, and Delete** text snippets from your slides, with persistent storage via a local Node.js/Express backend. This add-in is ideal for demoing Office.js APIs, building productivity tools, or learning how to connect PowerPoint with external APIs.

---

## ✨ Features

- Read selected text from PowerPoint slides/shapes
- Save selected text to a backend API (with unique IDs)
- Load, update, or delete saved texts by ID
- See a list of all saved texts and click to load them
- Clean, responsive UI for the PowerPoint taskpane
- Works with PowerPoint Desktop and PowerPoint Online (with HTTPS hosting)
- Easy to run locally for development and testing

---

## 📦 Folder Structure

```
ppt-text-crud-addin/
├── manifest.xml           # Office Add-in manifest (edit for your ngrok/public URL)
├── package.json           # Node.js dependencies and scripts
├── server/
│   └── server.js          # Express backend (serves static files + API)
└── src/
    ├── taskpane.html      # Taskpane UI
    ├── taskpane.js        # Taskpane logic (CRUD, Office.js)
    └── taskpane.css       # Taskpane styles
```

---

## 🚀 Quick Start

### 1. Prerequisites

- [Node.js](https://nodejs.org/) (LTS recommended)
- [ngrok](https://ngrok.com/) (for HTTPS tunneling, required for PowerPoint Online)
- PowerPoint (Desktop or Online, with ability to sideload add-ins)

---

### 2. Install Dependencies

Open a terminal in the `ppt-text-crud-addin` folder and run:

```sh
npm install
```

---

### 3. Start the Backend Server

```sh
npm start
```

- The server will run on [http://localhost:3000](http://localhost:3000) by default.
- It serves the taskpane UI and the `/api/texts` endpoints.

---

### 4. Expose Your Server with ngrok (for HTTPS)

PowerPoint Online requires HTTPS. In a new terminal, run:

```sh
ngrok http 3000
```

- Copy the HTTPS URL shown by ngrok (e.g., `https://your-ngrok-id.ngrok-free.app`).

---

### 5. Update the Manifest

Edit [`manifest.xml`](manifest.xml):

- Replace all instances of the sample URL (`https://c20e3f272ae6.ngrok-free.app`) with your actual ngrok HTTPS URL.
    - Update `<SourceLocation>`, `<AppDomain>`, `<IconUrl>`, etc.

Example:

```xml
<SourceLocation DefaultValue="https://your-ngrok-id.ngrok-free.app/taskpane.html" />
<AppDomain>https://your-ngrok-id.ngrok-free.app</AppDomain>
```

---

### 6. Sideload the Add-in in PowerPoint

#### PowerPoint Online

1. Go to **Insert → Add-ins → My Add-ins → Manage My Add-ins → Upload My Add-in**.
2. Upload your updated `manifest.xml`.
3. Open a presentation, and launch the add-in from the Home tab.

#### PowerPoint Desktop

1. Go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
2. Add a network share or folder containing your `manifest.xml`.
3. Restart PowerPoint, then **Insert → My Add-ins → Shared Folder**.

---

### 7. Use the Add-in

- Select a text shape or text box in your slide.
- Use the **Read Selected** button to preview text.
- **Save Selected** to store it in the backend.
- Use the ID field and **Load/Update/Delete by ID** buttons to manage saved texts.
- The **Saved Texts** list shows all stored items; click any to load it by ID.
- **Clear Selected Text** removes text from the current selection.

---

## 🛠️ Development Notes

- The backend uses an **in-memory store**. Restarting the server will erase saved texts.
- For production, replace with a database or persistent storage.
- The add-in uses the PowerPoint-specific Office.js APIs for shape/text manipulation.
- All static files are served from `/src` via Express.

---

## 🧩 API Endpoints

- `POST   /api/texts`        — Save new text, returns `{ id, text }`
- `GET    /api/texts`        — List all saved texts
- `GET    /api/texts/:id`    — Get text by ID
- `PUT    /api/texts/:id`    — Update text by ID
- `DELETE /api/texts/:id`    — Delete text by ID

---

## 📝 Customization

- Edit [`src/taskpane.html`](src/taskpane.html), [`src/taskpane.js`](src/taskpane.js), and [`src/taskpane.css`](src/taskpane.css) for UI/logic changes.
- Update [`server/server.js`](server/server.js) for backend logic or to add persistent storage.

---

## ❓ Troubleshooting

- **Add-in not loading?** Make sure your manifest URLs match your ngrok HTTPS URL.
- **API errors?** Check the backend server logs and ensure ngrok is running.
- **PowerPoint Online requires HTTPS** — always use the ngrok HTTPS URL in the manifest.

---

## 📚 References

- [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/)
- [Office.js API Reference](https://learn.microsoft.com/javascript/api/overview/powerpoint)
- [ngrok documentation](https://ngrok.com/docs)

---

## 🏁 License

MIT License

---

**Happy hacking!**