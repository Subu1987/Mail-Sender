const Bottleneck = require("bottleneck");
const fs = require("fs");
const path = require("path");
const ipc = require("electron").ipcRenderer;
const XLSX = require("xlsx");
const nodemailer = require("nodemailer");
// const archiver = require('archiver');
const AdmZip = require("adm-zip");
const { electron } = require("process");
const { dialog, ipcMain, ipcRenderer } = require("electron");

fileName = document.getElementById("fileName");
fileContents = document.getElementById("fileContents");
fileContents2 = document.getElementById("fileContents2");
btnRead = document.getElementById("btnRead");

let pathName = path.join(__dirname, "Files");
var userId = "";
var password = "";
var mailSubject = "";
var mailBody = "";
var dirPath = "";
var Data = [];
var path1 = "";

// chk password visibiity
const togglePassword = document.getElementById("togglePassword");
const passwordInput = document.getElementById("Password");
togglePassword.addEventListener("click", function () {
  if (passwordInput.getAttribute("type") === "password") {
    passwordInput.setAttribute("type", "text");
  } else {
    passwordInput.setAttribute("type", "password");
  }

  togglePassword.classList.toggle("fa-eye-slash");
});

// send files
const element = document.getElementById("sendFiles");
element.addEventListener("click", sendFiles);

// excel sheet upload
function upload(event) {
  // console.log("The Upload Function is called")
  ipc.send("open-file-dialog-for-file");
  ipc.on("selected-file", function (event, path) {
    //event.preventDefault();
    // var path1 = path.basename(pathName);
    path1 = path;
    path1 = path1.toString();
    let index = path1.lastIndexOf("\\");
    let fileName = path1.slice(index + 1);
    document.getElementById("excel").value = fileName;

    hideValidationMessage(document.getElementById("excel"));
  });
}

//Credential Save
function onDirectorySelected(event) {
  ipcRenderer.send("open-directory-dialog");

  ipcRenderer.on("selected-directory", (event, path) => {
    // event.preventDefault();
    console.log(`Selected directory: ${path}`);
    dirPath = `${path}`.toString() + "/";

    console.log(dirPath);
    document.getElementById("dirPath").value = dirPath;

    hideValidationMessage(document.getElementById("dirPath"));
  });
}

// show validation msg
function showValidationMessage(inputElement, message) {
  // Check if a previous validation message exists for the input field
  const existingValidationMessage = inputElement.parentNode.querySelector(
    ".validation-message"
  );
  if (existingValidationMessage) {
    existingValidationMessage.remove();
  }

  const validationMessage = document.createElement("div");
  validationMessage.className = "validation-message";
  validationMessage.textContent = message;
  inputElement.parentNode.appendChild(validationMessage);
}

// hide validation msg
function hideValidationMessage(inputElement) {
  const validationMessage = inputElement.parentNode.querySelector(
    ".validation-message"
  );
  if (validationMessage) {
    validationMessage.remove();
  }
}

function isEmailValid(email) {
  // Regular expression to validate email format
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function formValidation() {
  const email = document.getElementById("Email").value.trim();
  const password = document.getElementById("Password").value.trim();
  const excelFile = document.getElementById("excel").value.trim();
  const dirPath = document.getElementById("dirPath").value.trim();

  const emailInput = document.getElementById("Email");
  emailInput.addEventListener("focus", () => hideValidationMessage(emailInput));
  emailInput.addEventListener("input", () => hideValidationMessage(emailInput));

  const passwordInput = document.getElementById("Password");
  passwordInput.addEventListener("focus", () =>
    hideValidationMessage(passwordInput)
  );
  passwordInput.addEventListener("input", () =>
    hideValidationMessage(passwordInput)
  );

  const excelInput = document.getElementById("excel");
  excelInput.addEventListener("focus", () => hideValidationMessage(excelInput));
  excelInput.addEventListener("input", () => hideValidationMessage(excelInput));

  const dirPathInput = document.getElementById("dirPath");

  if (dirPathInput.value === "") {
    hideValidationMessage(dirPathInput);
  }

  dirPathInput.addEventListener("focus", () =>
    hideValidationMessage(dirPathInput)
  );
  dirPathInput.addEventListener("input", () =>
    hideValidationMessage(dirPathInput)
  );

  // chk email not blank & valid email
  if (email.trim() === "") {
    showValidationMessage(
      document.getElementById("Email"),
      "Please enter an email address."
    );
    return false;
  } else if (!isEmailValid(email)) {
    showValidationMessage(
      document.getElementById("Email"),
      "Please enter a valid email address."
    );
    return false;
  } else {
    hideValidationMessage(document.getElementById("Email"));
  }

  // chk pass not blanked
  if (password.trim() === "") {
    showValidationMessage(
      document.getElementById("Password"),
      "Please enter a password."
    );
    return false;
  } else {
    hideValidationMessage(document.getElementById("Password"));
  }
  passwordInput.addEventListener("focus", () =>
    hideValidationMessage(passwordInput)
  );

  // chk excel not blank & file format correct
  if (excelFile.trim() === "") {
    showValidationMessage(
      document.getElementById("excel"),
      "Please select a file."
    );
    return false;
  } else if (!/\.(csv|xls|xlsx)$/i.test(excelFile)) {
    showValidationMessage(
      document.getElementById("excel"),
      "Please choose a valid Excel CSV, XLS, or XLSX file."
    );
    return false;
  } else {
    hideValidationMessage(document.getElementById("excel"));
  }

  // chk directory path not blank
  if (dirPath.trim() === "") {
    showValidationMessage(
      document.getElementById("dirPath"),
      "Please choose a directory path."
    );
    return false;
  } else {
    hideValidationMessage(document.getElementById("dirPath"));
  }
  dirPathInput.addEventListener("focus", () =>
    hideValidationMessage(dirPathInput)
  );

  return true;
}

function showMessage(message, isSuccess) {
  const messageBox = document.getElementById("messageBox");
  const messageText = document.getElementById("messageText");

  messageText.textContent = message;

  if (isSuccess) {
    messageBox.classList.remove("error");
    messageBox.classList.add("success");
  } else {
    messageBox.classList.remove("success");
    messageBox.classList.add("error");
  }

  // Display the message box
  messageBox.style.display = "block";

  // Hide the message box after a few seconds
  setTimeout(function () {
    messageBox.style.display = "none";
  }, 5000);
}



// Create a rate limiter with a limit of 5 emails per second
const limiter = new Bottleneck({ maxConcurrent: 5, minTime: 200 });

function sendFiles(event) {
  // validate the form 1st
  if (!formValidation()) {
    // if not valid then stop the execution
    return;
  }

  userId = document.getElementById("Email").value.trim();
  password = document.getElementById("Password").value.trim();

  var transporter = nodemailer.createTransport({
    // host: "smtp.mailtrap.io",
    // port: 2525,
    service: "hotmail",
    auth: {
      user: userId,
      pass: password,
    },
  });

  const formElement = document.getElementById("formbox1");
  const loadingElement = document.getElementById("loading");
  loadingElement.style.display = "block";
  formElement.classList.add("blur");

  const workbook = XLSX.readFile(path1);
  const sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  // console.log(data.filter(obj => obj['E mail']));
  Data = data.filter((obj) => obj["E mail"]);
  console.log(Data);
  var filePath = dirPath;

  const zip = new AdmZip();

  // Create an array to track promises for email sending
  const emailPromises = [];

  Data.forEach(iterate);

  function iterate(item) {
    for (var i = 1; i <= 2; i++) {
      zip.addLocalFile(filePath + item[`CODE${i}`] + ".pdf");
    }

    // Get the current date in the format "dd_mm_yyyy"
    const currentDate = new Date();
    const day = String(currentDate.getDate()).padStart(2, "0");
    const month = String(currentDate.getMonth() + 1).padStart(2, "0");
    const year = currentDate.getFullYear();
    const downloadName = `${day}_${month}_${year}.zip`;

    // zipped file
    const data1 = zip.toBuffer();
    zip.writeZip(filePath + downloadName);
      // Now the zip file is fully created, proceed with email sending

      var mailOptions = {
        from: userId,
        to: item["E mail"],
        subject: "Mail Sender",
        attachments: [
          {
            filename: downloadName,
            path: filePath + downloadName,
          },
        ],
      };

      // Use the rate limiter to control the email sending rate
      const emailPromise = limiter.schedule(() => {
        return new Promise((resolve, reject) => {
          transporter.sendMail(mailOptions, function (error, info) {
            if (error) {
              console.log(error);

              const errorMessage = error.response
                ? error.message.replace(/^Error:\s(.+?)\s\[.+/, "$1")
                : "Error sending email. Please try again later.";

              // Delete the zip file after email is sent successfully
              if (fs.existsSync(filePath + downloadName)) {
                fs.unlink(filePath + downloadName, (err) => {
                  if (err) {
                    console.error("Error deleting zip file:", err);
                  } else {
                    console.log("Zip file deleted successfully.");
                  }
                });
              }

              // show the error msg
              showMessage(errorMessage, false);

              loadingElement.style.display = "none";
              formElement.classList.remove("blur");
              reject(error);
            } else {
              console.log("Email sent: " + info.response);

              // Delete the zip file after email is sent successfully
              if (fs.existsSync(filePath + downloadName)) {
                fs.unlink(filePath + downloadName, (err) => {
                  if (err) {
                    console.error("Error deleting zip file:", err);
                  } else {
                    console.log("Zip file deleted successfully.");
                  }
                });
              }

              // show the success msg
              showMessage("Email Send Successfully", true);

              loadingElement.style.display = "none";
              formElement.classList.remove("blur");

              document.getElementById("formbox1").reset();
              resolve(info);
            }
          });
        });
      });

      // Add the email promise to the array
      emailPromises.push(emailPromise);

      filePath = dirPath;
  }
  // Handle all email promises together after they have resolved or rejected
  Promise.all(emailPromises)
    .then((results) => {
      console.log("All emails sent successfully!", results);
    })
    .catch((error) => {
      console.error("Email sending failed!", error);
    });
}
