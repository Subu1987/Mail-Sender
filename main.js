const electron = require("electron");
const url = require("url");
const path = require("path");


// App Object and browser Window object Pulled out from electron
const { app, BrowserWindow, Menu, ipcMain, dialog } = electron;
// Set Environment
// process.env.NODE_ENV = 'production';

let mainWindow;
let addWindow;
let uploadWindow;
let downloadWindow;

// A function will execute when the app is ready
app.on("ready", function () {

  createFileUpload();

});
//Handle Create Add Window
function createAddWindow() {
  // An empty object is passed in the browser window constructer because at the starting there is no item in the list
  addWindow = new BrowserWindow({
    width: 300,
    height: 200,
    title: "Add Shopping list Item",
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
    },
  });
  // Process to Load HTML in Main Window
  addWindow.loadURL(
    url.format({
      pathname: path.join(__dirname, "addWindow.html"),
      protocol: "file:",
      slashes: true,
    })
  );
  // Handle Garbage Collection for add item when the app is closed
  addWindow.on("close", function () {
    addWindow = null;
  });
}
// Handle File Upload
function createFileUpload() {
  // An empty object is passed in the browser window constructer because at the starting there is no item in the list
  uploadWindow = new BrowserWindow({
    width: 800,
    height: 600,
    title: "Upload File",
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
    },
    autoHideMenuBar: false,
  });
  // Process to Load HTML in Main Window
  uploadWindow.loadURL(
    url.format({
      pathname: path.join(__dirname, "uploadWindow2.html"),
      protocol: "file:",
      slashes: true,
    })
  );
  // console.log("The Pathname is " + pathname)
  // Handle Garbage Collection for add item when the app is closed
  uploadWindow.on("close", function () {
    uploadWindow = null;
  });
}
// Handle Download File
function createDownloadFile() {
  // An empty object is passed in the browser window constructer because at the starting there is no item in the list
  downloadWindow = new BrowserWindow({
    width: 350,
    height: 250,
    title: "Download File",
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
    },
  });
  // Process to Load HTML in Main Window
  downloadWindow.loadURL(
    url.format({
      pathname: path.join(__dirname, "downloadWindow.html"),
      protocol: "file:",
      slashes: true,
    })
  );
  // Handle Garbage Collection for add item when the app is closed
  downloadWindow.on("close", function () {
    downloadWindow = null;
  });
}

//catch item:add
ipcMain.on("item:add", function (e, item) {
  console.log(item);
  mainWindow.webContents.send("item:add", item);
  addWindow.close();
});
// Create Menu Template. An array for the list of items
const mainMenuTemplate = [
  // {
  //   label: "File",
  //   // Options inside a particular menu
  //   submenu: [
  //     {
  //       label: "Add Item",
  //       click() {
  //         createAddWindow();
  //       },
  //     },
  //     {
  //       label: "Clear items",
  //       click() {
  //         mainWindow.webContents.send("item:clear");
  //       },
  //     },
  //     {
  //       label: "quit",
  //       // Shortcut for the option
  //       accelerator: process.platform == "darwin" ? "command+Q" : "ctrl+Q",
  //       click() {
  //         app.quit();
  //       },
  //     },
  //   ],
  // },
  {
    label: "File Upload",
    click() {
      createFileUpload();
    },
  },
  // {
  //   label: "Download File",
  //   click() {
  //     createDownloadFile();
  //   },
  // },
];
// If we are in mac the main menu template will not shown properly it will show only electron as a default menu. For fixing that we have to add first object in the main menu template as a empty object.
if (process.platform == "darwin") {
  mainMenuTemplate.unshift({});
}

//If our app is not in production then developer tools is added to the main menu template.
if (process.env.NODE_ENV !== "production") {
  mainMenuTemplate.push({
    label: "Developer Tools",
    submenu: [
      {
        label: "Toggle Dev Tools",
        accelerator: process.platform == "darwin" ? "command+I" : "ctrl+I",
        click(item, focusedWindow) {
          focusedWindow.toggleDevTools();
        },
      },
      {
        role: "reload",
      },
    ],
  });
}

// Retrieve event send by the renderer process for upload
ipcMain.on("open-file-dialog-for-file", function (event) {
  if (process.platform === "win32" || process.platform === "linux") {
    dialog
      .showOpenDialog(uploadWindow, {
        properties: ["openFile"],
      })
      .then((files) => {
        // console.log(files);
        if (files) {
          event.sender.send("selected-file", files.filePaths);
        }
      })
      .catch((err) => {
        console.log(err);
      });
  }
  else {
    dialog
      .showOpenDialog(uploadWindow, {
        properties: ["openFile", "openDirectory"],
      })
      .then((files) => {
        // console.log(files);
        if (files) {
          event.sender.send("selected-file", files.filePaths);
        }
      })
      .catch((err) => {
        console.log(err);
      });
  }
});

//Retrieve event send by the renderer process for Download 
ipcMain.on("open-folder-dialog-for-file", function (event) {
  dialog
    .showOpenDialog(downloadWindow, {
      properties: ["openDirectory"],
    })
    .then((files) => {
      if (files) {
        event.sender.send("selected-folder", files.filePaths);
      }
    })
    .catch((err) => {
      console.log(err);
    });
});

ipcMain.on('open-directory-dialog', (event) => {
  dialog.showOpenDialog({ properties: ['openDirectory'] }).then(result => {
    if (!result.canceled) {
      event.sender.send('selected-directory', result.filePaths[0]);
    }
  }).catch(err => {
    console.log(err);
  });
});
