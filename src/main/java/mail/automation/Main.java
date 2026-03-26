package mail.automation;

import java.util.concurrent.atomic.AtomicBoolean;

import javax.swing.JProgressBar;

import mail.automation.excel.ExcelReader;

public class Main {

    public static void main(String[] args) {
        System.out.println("Starting Application...");

        App app = new App();
        // ✅ Prepare stop flag
        AtomicBoolean stopFlag = new AtomicBoolean(false);

        // ✅ Optional progress bar (null if running from console)
        JProgressBar progressBar = null;

        // Let user pick Excel file at runtime
        ExcelReader reader = new ExcelReader();
        String excelPath = reader.pickExcelFile();

        if (excelPath != null) {
            // Call App logic
            app.runWithLogging(System.out::println, stopFlag, progressBar, excelPath);
        } else {
            System.out.println("No Excel file selected. Exiting.");
        }
    }

}