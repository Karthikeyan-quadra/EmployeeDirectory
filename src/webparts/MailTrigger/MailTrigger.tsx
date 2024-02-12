import "@pnp/sp/lists";
import { IEmailProperties } from "@pnp/sp/presets/all";
import { SPFI } from "@pnp/sp";
import { addSP } from "../Services/PnpConfig";
import "@pnp/sp/sputilities";

export async function Requestmail(msg: any, userMail: any, displayName:any, senderName:any,senderJob:any,senderDept:any) {
  try {
    // Check if the message is not empty
    if (!msg.trim()) {
      throw new Error("Message cannot be empty");
    }

    const emailProps: IEmailProperties = {
      To: [userMail],
      CC: [],
      BCC: [],
      Subject: `Message from ${senderName}`,
      Body: `
        <html>
          <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
          </head>
          <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
            <p>Dear ${displayName},</p>
            <p>${msg}</p>

            <p style="margin-top:100px; font-size:14px; margin-bottom:5px;"><b>Thanks & Regards,</b></p>
            <p style="margin-top:5px; font-size:14px; margin-bottom:0px;"><b>${senderName} | ${senderJob} - ${senderDept}</b></p>
            <p style="font-size:14px; margin-top:5px;"><span style="color:rgb(255,102,0);"><b>Quadrasystems.net (India)</b></span><b> Private Limited</b></p>

          </body>
        </html>
      `,
    };

    // Ensure that addSP returns a valid SPFI object
    const sp: SPFI = addSP();

    if (!sp) {
      throw new Error("Unable to initialize SPFI object");
    }

    await sp.utility.sendEmail(emailProps);
    console.log("Email sent successfully");
  } catch (error) {
    console.error("Error sending email:", error);
    throw error; // Rethrow the error if needed for further handling
  }
}
