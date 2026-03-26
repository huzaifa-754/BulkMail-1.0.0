package mail.automation.util;

import mail.automation.model.MailData;

public class HtmlBuilder {

    public static String buildBody(MailData mail, String tableHtml) {

        StringBuilder html = new StringBuilder();

        html.append("<html><body>");

        html.append("").append(mail.getBody()).append(",<br><br>");

        html.append("This is an automated email.<br><br>");

        html.append("<b>Please find details below:</b><br><br>");

        // ✅ Dynamic Table
        html.append(tableHtml).append("<br><br>");

        html.append("</body></html>");

        return html.toString();
    }
}