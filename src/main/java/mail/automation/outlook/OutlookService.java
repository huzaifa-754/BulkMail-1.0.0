package mail.automation.outlook;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
//import com.jacob.com.Variant;
import mail.automation.model.MailData;
import mail.automation.util.HtmlBuilder;

import java.io.File;

public class OutlookService {

    /**
     * Send mail via Outlook (Jacob)
     */
    public void sendMail(MailData mailData, String tableHtml, String[] logos) {
        try {
            // Initialize Outlook application
            ActiveXComponent outlook = new ActiveXComponent("Outlook.Application");
            Dispatch mail = outlook.invoke("CreateItem", 0).toDispatch(); // 0 = MailItem

            // Set To, CC, Subject
            Dispatch.put(mail, "To", safe(mailData.getTo()));
            Dispatch.put(mail, "CC", safe(mailData.getCc()));
            Dispatch.put(mail, "Subject", safe(mailData.getSubject()));

            // Attachments
            Dispatch attachments = Dispatch.get(mail, "Attachments").toDispatch();

            addAttachmentIfExists(attachments, mailData.getAttachment1());
            addAttachmentIfExists(attachments, mailData.getAttachment2());

            // ================= INLINE LOGOS =================
            // logos[0] = logo1, logos[1] = logo2
            // addInlineAttachment(attachments, logos != null ? logos[0] : null, "logo1");
            // addInlineAttachment(attachments, logos != null ? logos[1] : null, "logo2");

            // Add logos as INLINE (CID)
            addInlineImage(mail, logos[0], "logo1");
            addInlineImage(mail, logos[1], "logo2");

            // Build HTML Body
            String htmlBody = HtmlBuilder.buildBody(mailData, tableHtml);

            Dispatch.put(mail, "HTMLBody", htmlBody);

            // Send mail
            Dispatch.call(mail, "Send");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Safe string handler (avoid nulls)
     */
    private String safe(String value) {
        return value == null ? "" : value;
    }

    /**
     * Add normal attachment
     */
    private void addAttachmentIfExists(Dispatch attachments, String filePath) {
        try {
            if (filePath != null && !filePath.trim().isEmpty()) {
                File file = new File(filePath.trim());

                if (file.exists()) {
                    Dispatch.call(attachments, "Add", file.getAbsolutePath());
                } else {
                    System.out.println("Attachment not found: " + filePath);
                }
            }
        } catch (Exception e) {
            System.out.println("Error adding attachment: " + filePath + " - " + e.getMessage());
        }
    }

    /**
     * Add inline image (for logos using CID)
     */
    // private void addInlineAttachment(Dispatch attachments, String filePath,
    // String cid) {
    // try {
    // if (filePath != null && !filePath.trim().isEmpty()) {
    // File file = new File(filePath.trim());

    // if (file.exists()) {
    // // Attach with Content-ID for inline display
    // Dispatch.call(attachments, "Add", file.getAbsolutePath(), 1, 0, cid);
    // } else {
    // System.out.println("Logo not found: " + filePath);
    // }
    // }
    // } catch (Exception e) {
    // System.out.println("Error adding inline image: " + filePath + " - " +
    // e.getMessage());
    // }
    // }

    private void addInlineImage(Dispatch mail, String path, String cid) {
        try {
            if (path == null || path.isEmpty())
                return;

            File f = new File(path);
            if (!f.exists()) {
                System.out.println("Logo not found: " + path);
                return;
            }

            Dispatch attachments = Dispatch.get(mail, "Attachments").toDispatch();
            Dispatch attachment = Dispatch.call(attachments, "Add", path).toDispatch();

            // 🔥 VERY IMPORTANT (CID binding)
            Dispatch propertyAccessor = Dispatch.call(attachment, "PropertyAccessor").toDispatch();

            Dispatch.call(propertyAccessor, "SetProperty",
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid);

            // Hide attachment
            Dispatch.put(attachment, "Position", 0);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}