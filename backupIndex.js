const Bottleneck = require("bottleneck");
const { ipcRenderer } = require("electron");
const XLSX = require("xlsx");
const nodemailer = require("nodemailer");
const AdmZip = require("adm-zip");
const fs = require("fs");

const formElement = document.getElementById("formbox1");
const loadingElement = document.getElementById("loading");
const mailCountBox = document.getElementById('mailCount');
const mailCountNo = document.getElementById('mailCountNo');

let mail_ID = "";
let password = "";
let dirPath = "";
let path1 = "";
let logFilePath = "";

// chk password visibility
const togglePassword = document.getElementById("togglePassword");
const passwordInput = document.getElementById("Password");
togglePassword.addEventListener("click", function () {
  passwordInput.setAttribute("type", passwordInput.getAttribute("type") === "password" ? "text" : "password");
  togglePassword.classList.toggle("fa-eye-slash");
});

// send files
document.getElementById("sendFiles").addEventListener("click", sendFiles);

// excel sheet upload
function upload(event) {
  ipcRenderer.send("open-file-dialog-for-file");
  ipcRenderer.on("selected-file", function (event, path) {
    path1 = path.toString();
    document.getElementById("excel").value = path1.slice(path1.lastIndexOf("\\") + 1);
    hideValidationMessage(document.getElementById("excel"));
  });
}

// get the directory path
function onDirectorySelected(event) {
  ipcRenderer.send("open-directory-dialog");
  ipcRenderer.on("selected-directory", (event, path) => {
    dirPath = `${path}/`;
    document.getElementById("dirPath").value = dirPath;
    hideValidationMessage(document.getElementById("dirPath"));

    // Get today's date in the format yyyy-mm-dd
    const today = new Date().toLocaleDateString('en-US', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    }).replace(/\//g, '-');

    // Set the log file path with today's date
    logFilePath = `${dirPath}/email_logs_${today}.txt`;

  });
}

// show validation message
function showValidationMessage(inputElement, message) {
  const existingValidationMessage = inputElement.parentNode.querySelector(".validation-message");
  if (existingValidationMessage) {
    existingValidationMessage.remove();
  }
  const validationMessage = document.createElement("div");
  validationMessage.className = "validation-message";
  validationMessage.textContent = message;
  inputElement.parentNode.appendChild(validationMessage);
}

// hide validation message
function hideValidationMessage(inputElement) {
  const validationMessage = inputElement.parentNode.querySelector(".validation-message");
  if (validationMessage) {
    validationMessage.remove();
  }
}

// email validation
function isEmailValid(mail_ID) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(mail_ID);
}

// form validation
function formValidation() {
  mail_ID = document.getElementById("Email").value.trim();
  password = passwordInput.value.trim();
  const excelFile = document.getElementById("excel").value.trim();
  dirPath = document.getElementById("dirPath").value.trim();

  const emailInput = document.getElementById("Email");
  emailInput.addEventListener("focus", () => hideValidationMessage(emailInput));
  emailInput.addEventListener("input", () => hideValidationMessage(emailInput));

  passwordInput.addEventListener("focus", () => hideValidationMessage(passwordInput));
  passwordInput.addEventListener("input", () => hideValidationMessage(passwordInput));

  const excelInput = document.getElementById("excel");
  excelInput.addEventListener("focus", () => hideValidationMessage(excelInput));
  excelInput.addEventListener("input", () => hideValidationMessage(excelInput));

  const dirPathInput = document.getElementById("dirPath");
  dirPathInput.addEventListener("focus", () => hideValidationMessage(dirPathInput));
  dirPathInput.addEventListener("input", () => hideValidationMessage(dirPathInput));

  if (dirPathInput.value === "") {
    hideValidationMessage(dirPathInput);
  }

  if (mail_ID.trim() === "") {
    showValidationMessage(emailInput, "Please enter an email address.");
    return false;
  } else if (!isEmailValid(mail_ID)) {
    showValidationMessage(emailInput, "Please enter a valid email address.");
    return false;
  }

  if (password.trim() === "") {
    showValidationMessage(passwordInput, "Please enter a password.");
    return false;
  }

  if (excelFile.trim() === "") {
    showValidationMessage(excelInput, "Please select a file.");
    return false;
  } else if (!/\.(csv|xls|xlsx)$/i.test(excelFile)) {
    showValidationMessage(excelInput, "Please choose a valid Excel CSV, XLS, or XLSX file.");
    return false;
  }

  if (dirPath.trim() === "") {
    showValidationMessage(dirPathInput, "Please choose a directory path.");
    return false;
  }

  return true;
}

// show the success message
function showMessage(message, isSuccess) {
  const messageBox = document.getElementById("messageBox");
  const messageText = document.getElementById("messageText");

  messageText.textContent = message;
  messageBox.classList.remove(isSuccess ? "error" : "success");
  messageBox.classList.add(isSuccess ? "success" : "error");
  messageBox.style.display = "block";

  setTimeout(() => messageBox.style.display = "none", 5000);
}

// Create a rate limiter with a limit of 5 emails per second
const limiter = new Bottleneck({ maxConcurrent: 5, minTime: 200 });
let mailCount = 0;

// Function to send files as email attachments
async function sendFiles() {
  if (!formValidation()) return;

  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: { user: mail_ID, pass: password },
    retry: { retries: 5, factor: 2, minTimeout: 1000 }
  });

  loadingElement.style.display = "block";
  mailCountBox.style.display = "block";
  formElement.classList.add("blur");

  const workbook = XLSX.readFile(path1);
  const sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

  const filteredData = data.filter(obj => obj.hasOwnProperty("E mail"));

  if (filteredData.length === 0) {
    showMessage("Invalid Excel format. No valid data with 'E mail' property found.", false);
    loadingElement.style.display = "none";
    mailCountBox.style.display = "none";
    formElement.classList.remove("blur");
    return;
  }

  let maxCodeFileNo = 0;
  filteredData.forEach(obj => {
    Object.keys(obj).forEach(key => {
      if (key.startsWith("CODE")) {
        const codeFileNo = parseInt(key.substring(4));
        if (codeFileNo > maxCodeFileNo) {
          maxCodeFileNo = codeFileNo;
        }
      }
    })
  });

  try {
    await Promise.all(
      filteredData.map(async (item) => {
        // Create a new instance of AdmZip for each recipient
        const zip = new AdmZip();
        let filePath = dirPath; // Set the correct filePath for each email
        let filesToSend = [];

        for (let i = 1; i <= maxCodeFileNo; i++) {
          if (!item[`CODE${i}`]) {
            continue;
          } else {
            const fileName = item[`CODE${i}`] + ".pdf";
            if (fs.existsSync(filePath + fileName)) {
              zip.addLocalFile(filePath + item[`CODE${i}`] + ".pdf");
              filesToSend.push(item[`CODE${i}`]); // Collect files to be sent for logging
            } else {
              // showMessage(`File ${fileName} not found. Skipping...`, false);
              console.warn(`File ${fileName} not found. Skipping...`);
            }
          }
        }

        const downloadName = item['TRADER NAME'] + ".zip";
        zip.writeZip(filePath + downloadName);

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

        try {
          await limiter.schedule(() => {
            return new Promise((resolve, reject) => {
              transporter.sendMail(mailOptions, function (error, info) {
                if (error) {
                  console.log(error);
                  mailCount = 0;
                  mailCountBox.style.display = "none";
                  showMessage(error.response ? error.message.replace(/^Error:\s(.+?)\s\[.+/, "$1") : "Error sending email. Please try again later.", false);
                  reject(error);
                } else {
                  console.log("Email sent: " + info.response);
                  mailCount++;
                  mailCountNo.textContent = mailCount;
                  resolve(info);
                }
              });
            }).then(() => {
              // Log the email sent along with the files for this recipient
              let currentDate = new Date().toLocaleString('en-US', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: 'numeric',
                minute: 'numeric',
                second: 'numeric',
                hour12: true
              });
              let logEntry = `${currentDate} - Email sent to: ${item['E mail']}, Files: ${filesToSend.join(", ")}`;
              fs.appendFileSync(logFilePath, logEntry + '\n');

            });
          });
        } catch (error) {
          console.error("Email sending failed!", error);
          throw error;
        } finally {
          // Delete the zip file after email is sent successfully
          fs.unlinkSync(filePath + downloadName);
          filePath = dirPath; // Reset filePath for the next email iteration
        }

      })
    );
    mailCountBox.style.display = "none";
    showMessage(`All Emails Sent Successfully : [\u0020${mailCount}\u0020]`, true);
    mailCount = 0;
  } catch (error) {
    showMessage(error, false);
  } finally {
    loadingElement.style.display = "none";
    formElement.classList.remove("blur");
    mailCountNo.textContent = 0;
    document.getElementById("formbox1").reset();
  }
}
