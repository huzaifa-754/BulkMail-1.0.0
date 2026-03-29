package mail.automation;

// import java.util.concurrent.atomic.AtomicBoolean;

// import javax.swing.JProgressBar;
import javax.swing.SwingUtilities;

// import mail.automation.excel.ExcelReader;
import mail.automation.ui.ControlPanel;

public class Main {

    public static void main(String[] args) {
        System.out.println("Starting Application...");

        SwingUtilities.invokeLater(() -> {
        ControlPanel cp = new ControlPanel();
        cp.setVisible(true);
        });

        // App app = new App();
        // // ✅ Prepare stop flag
        // AtomicBoolean stopFlag = new AtomicBoolean(false);

        // // ✅ Optional progress bar (null if running from console)
        // JProgressBar progressBar = null;

        // // Let user pick Excel file at runtime
        // ExcelReader reader = new ExcelReader();
        // String excelPath = reader.pickExcelFile();

        // if (excelPath != null) {
        //     // Call App logic
        //     app.runWithLogging(System.out::println, stopFlag, progressBar, excelPath);
        // } else {
        //     System.out.println("No Excel file selected. Exiting.");
        // }
    }

}