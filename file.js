const Bottleneck = require("bottleneck");
const fs = require("fs");
const path = require("path");
const ipc = require("electron").ipcRenderer;
const XLSX = require("xlsx");
const nodemailer = require("nodemailer");
const AdmZip = require("adm-zip");
const { electron } = require("process");
const { dialog, ipcMain, ipcRenderer } = require("electron");

// Rest of your code...

async function sendFiles(event) {
  // validate the form first
  if (!formValidation()) {
    // If not valid, stop the execution
    return;
  }

  userId = document.getElementById("Email").value.trim();
  password = document.getElementById("Password").value.trim();

  const transporter = nodemailer.createTransport({
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
  Data = data.filter((obj) => obj["E mail"]);
  console.log(Data);
  const filePath = dirPath;

  const zip = new AdmZip();

  try {
    for (const item of Data) {
      for (let i = 1; i <= 2; i++) {
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
      await limiter.schedule(async () => {
        try {
          const info = await transporter.sendMail(mailOptions);
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

          showMessage("Email Send Successfully", true);
        } catch (error) {
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

          showMessage(errorMessage, false);
        }
      });
    }

    // All emails sent successfully
    console.log("All emails sent successfully!");
  } catch (error) {
    console.error("Email sending failed!", error);
  } finally {
    loadingElement.style.display = "none";
    formElement.classList.remove("blur");
    document.getElementById("formbox1").reset();
  }
}

// Rest of your code...
