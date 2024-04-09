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
let mail_ID = "";
let password = "";
let mailSubject = "";
let mailBody = "";
let dirPath = "";
let Data = [];
let path1 = "";

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
    // let path1 = path.basename(pathName);
    path1 = path;
    path1 = path1.toString();
    let index = path1.lastIndexOf("\\");
    let fileName = path1.slice(index + 1);
    document.getElementById("excel").value = fileName;

    hideValidationMessage(document.getElementById("excel"));
  });
}

// get the directory path
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

// email validation
function isEmailValid(mail_ID) {
  // Regular expression to validate email format
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(mail_ID);
}

// form validation
function formValidation() {
        mail_ID = document.getElementById("Email").value.trim();
        password = document.getElementById("Password").value.trim();
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
  if (mail_ID.trim() === "") {
    showValidationMessage(
      document.getElementById("Email"),
      "Please enter an email address."
    );
    return false;
  } else if (!isEmailValid(mail_ID)) {
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

// show the success message
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
  }, 7000);
}



// Create a rate limiter with a limit of 5 emails per second
const limiter = new Bottleneck({ maxConcurrent: 5, minTime: 200 });
let mailCount = 0;

// Function to send files as email attachments
async function sendFiles(event) {
  // Validate the form first
  if (!formValidation()) {
    // If not valid, stop the execution
    return;
  }

  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: mail_ID,
      pass: password,
    },
    // Retry configuration
    retry: {
      retries: 5, // Number of retries
      factor: 2, // Exponential backoff factor
      minTimeout: 1000, // Minimum time between retries in milliseconds
    }
  });

  const formElement = document.getElementById("formbox1");
  const loadingElement = document.getElementById("loading");
  const mailCountBox = document.getElementById('mailCount');
  let mailCountNo = document.getElementById('mailCountNo');
  loadingElement.style.display = "block";
  mailCountBox.style.display = "block";
  formElement.classList.add("blur");

  const workbook = XLSX.readFile(path1);
  const sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

  // Filter data to keep only objects with the "E mail" property
  const filteredData = data.filter(obj => obj.hasOwnProperty("E mail"));

  // If there are no valid objects, show an error message and stop further execution
  if (filteredData.length === 0) {
    showMessage("Invalid Excel format. No valid data with 'E mail' property found.", false);
    loadingElement.style.display = "none";
    mailCountBox.style.display = "none";
    formElement.classList.remove("blur");
    return;
  }

  // dynamic the no of files go though
  let maxCodeFileNo = 0;
  filteredData.forEach(obj => {
    Object.keys(obj).forEach(key => {
      if (key.startsWith("CODE")) {
        // Extract number after "CODE"
        const codeFileNo = parseInt(key.substring(4));
        if (codeFileNo > maxCodeFileNo) {
          maxCodeFileNo = codeFileNo;
        }
      }
    })
  })

  try {
    await Promise.all(
      filteredData.map(async (item) => {
        // Create a new instance of AdmZip for each recipient
        const zip = new AdmZip();
        let filePath = dirPath; // Set the correct filePath for each email

        for (let i = 1; i <= maxCodeFileNo; i++) {
          if (!item[`CODE${i}`]) {
            continue;
          } else {
            zip.addLocalFile(filePath + item[`CODE${i}`] + ".pdf");
          }

        }

        // Get the current date in the format "dd_mm_yyyy"
        // const currentDate = new Date();
        // const day = String(currentDate.getDate()).padStart(2, "0");
        // const month = String(currentDate.getMonth() + 1).padStart(2, "0");
        // const year = currentDate.getFullYear();
        const downloadName = item['TRADER NAME'] + ".zip";

        // Create the zipped file
        zip.writeZip(filePath + downloadName);

        // Now the zip file is fully created, proceed with email sending
        const mailOptions = {
          from: mail_ID,
          to: item["E mail"],
          subject: "File Attachment",
          text: `Hi ${item['TRADER NAME']}, please check the attachments`,
          attachments: [
            {
              filename: downloadName,
              path: filePath + downloadName,
            },
          ],
        };

        // Use the rate limiter to control the email sending rate
        try {
          await limiter.schedule(() => {
            return new Promise((resolve, reject) => {
              transporter.sendMail(mailOptions, function (error, info) {
                if (error) {
                  console.log(error);

                  const errorMessage = error.response
                    ? error.message.replace(/^Error:\s(.+?)\s\[.+/, "$1")
                    : "Error sending email. Please try again later.";


                  mailCount = 0;
                  mailCountBox.style.display = "none";
                  // Show the error message
                  showMessage(errorMessage, false);

                  reject(error);
                } else {
                  console.log("Email sent: " + info.response);
                  mailCount++;
                  mailCountNo.textContent = mailCount;

                  // Delete the zip file after email is sent successfully
                  // if (fs.existsSync(filePath + downloadName)) {
                  //   fs.unlink(filePath + downloadName, (err) => {
                  //     if (err) {
                  //       console.error("Error deleting zip file:", err);
                  //     } else {
                  //       console.log("Zip file deleted successfully.");
                  //     }
                  //   });
                  // }

                  // Resolve the promise on success
                  resolve(info);
                }
              });
            });
          });
        } catch (error) {
          console.error("Email sending failed!", error);
          throw error; // Throw the error again to handle it outside this catch block
        } finally {
          filePath = dirPath; // Reset filePath for the next email iteration
        }
      })
    );
    mailCountBox.style.display = "none";
    // Show the success message after all emails are sent successfully
    showMessage(`All Emails Sent Successfully : [\u0020${mailCount}\u0020]`, true);
    mailCount = 0;
  } catch (error) {
    // Show the error message if any email encounters an error
    showMessage(error, false);
  } finally {
    // Hide the loading box and reset the form
    loadingElement.style.display = "none";
    formElement.classList.remove("blur");
    mailCountNo.textContent = 0;
    document.getElementById("formbox1").reset();
  }
}
