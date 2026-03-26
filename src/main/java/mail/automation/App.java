package mail.automation;

import mail.automation.excel.ExcelReader;
import mail.automation.model.MailData;
import mail.automation.outlook.OutlookService;

import javax.swing.*;
import java.util.List;
import java.util.Map;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;

public class App {

    /**
     * Optimized BulkMail sending with:
     * - Pre-read body & signature
     * - GST table read once in memory
     * - Stop button support
     * - Progress logging
     * - Multi-threaded sending
     */
    public void runWithLogging(Consumer<String> logger, AtomicBoolean stopFlag, JProgressBar progressBar,
            String excelPath) {

        try {
            ExcelReader reader = new ExcelReader();
            List<MailData> mails = reader.readBulkMail(excelPath);
            OutlookService outlook = new OutlookService();

            // 1️⃣ Pre-read body & signature once
            String body = reader.buildBodyContent(excelPath);
            String signature = reader.buildSignature(excelPath);
            String[] logos = reader.getLogoPaths(excelPath);

            // 2️⃣ Pre-read GST table into memory
            List<Map<String, String>> gstRows = reader
                    .readGstSheet(excelPath, "GST");

            // 3️⃣ ExecutorService for multi-threaded sending
            int threadCount = 3; // adjust based on Outlook safety
            ExecutorService executor = Executors.newFixedThreadPool(threadCount);

            AtomicInteger count = new AtomicInteger(0);
            int total = mails.size();

            for (MailData mail : mails) {
                executor.submit(() -> {
                    try {
                        // Stop check
                        if (stopFlag.get())
                            return;

                        // Filter GST table for this mail
                        String tableHtml = reader.convertRowsToHtmlTable(gstRows, mail.getFilterValue(), "SUPPLIER2");

                        String finalHtml = "<html><body style='font-family:Calibri;font-size:11pt;'>" + body +
                                tableHtml +
                                "<br><br>" +
                                signature +
                                "</body></html>";

                        // Send mail
                        outlook.sendMail(mail, finalHtml, logos);

                        // Update count
                        int sent = count.incrementAndGet();

                        // Log message
                        logger.accept("Mail sent to: " + mail.getTo() + " (" + sent + "/" + total + ")");

                        // Update Swing progress bar safely
                        if (progressBar != null) {
                            SwingUtilities.invokeLater(() -> progressBar.setValue((int) ((sent * 100.0) / total)));
                        }

                        // Optional small delay to avoid Outlook issues
                        Thread.sleep(200);

                    } catch (Exception e) {
                        logger.accept("Error sending to: " + mail.getTo() + " - " + e.getMessage());
                    }
                });
            }

            // Shutdown executor and wait for tasks to finish
            executor.shutdown();
            executor.awaitTermination(30, TimeUnit.MINUTES);

            logger.accept("BulkMail completed!");

        } catch (Exception e) {
            logger.accept("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }
}