package mail.automation.ui;

import mail.automation.App;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.util.concurrent.atomic.AtomicBoolean;

public class ControlPanel extends JFrame {

    private JButton selectExcelButton;
    private JButton sendBulkMailButton;
    private JButton stopButton;
    private JButton autoReplyButton;
    private JProgressBar progressBar;
    private JTextArea logArea;

    private AtomicBoolean stopFlag = new AtomicBoolean(false);
    private String excelPath = null; // dynamically selected Excel file

    public ControlPanel() {
        setTitle("BulkMail Control Panel");
        setSize(700, 450);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // ---------------- TOP PANEL ----------------
        JPanel topPanel = new JPanel();
        selectExcelButton = new JButton("Select Excel File");
        sendBulkMailButton = new JButton("Send Bulk Mail");
        stopButton = new JButton("Stop");
        autoReplyButton = new JButton("AutoReply (Coming Soon)");

        topPanel.add(selectExcelButton);
        topPanel.add(sendBulkMailButton);
        topPanel.add(stopButton);
        topPanel.add(autoReplyButton);

        add(topPanel, BorderLayout.NORTH);

        // ---------------- CENTER LOG AREA ----------------
        logArea = new JTextArea();
        logArea.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(logArea);
        add(scrollPane, BorderLayout.CENTER);

        // ---------------- BOTTOM PROGRESS BAR ----------------
        progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        add(progressBar, BorderLayout.SOUTH);

        // ---------------- ACTIONS ----------------
        selectExcelButton.addActionListener((ActionEvent e) -> selectExcelFile());
        sendBulkMailButton.addActionListener((ActionEvent e) -> startBulkMail());
        stopButton.addActionListener((ActionEvent e) -> stopBulkMail());
        autoReplyButton.addActionListener((ActionEvent e) -> log("AutoReply feature coming soon!"));
    }

    // ---------------- SELECT EXCEL ----------------
    private void selectExcelFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        int option = fileChooser.showOpenDialog(this);
        if (option == JFileChooser.APPROVE_OPTION) {
            excelPath = fileChooser.getSelectedFile().getAbsolutePath();
            log("Selected Excel file: " + excelPath);
        } else {
            log("No Excel file selected.");
        }
    }

    // ---------------- SEND BULKMAIL ----------------
    private void startBulkMail() {
        if (excelPath == null) {
            log("Please select an Excel file first!");
            return;
        }
        sendBulkMailButton.setEnabled(false);
        stopFlag.set(false);
        log("BulkMail started...");
        progressBar.setValue(0);

        new Thread(() -> {
            App app = new App();
            // runWithLogging handles logging & progress internally
            app.runWithLogging(this::log, stopFlag, progressBar, excelPath);

            // After all mails are done
            SwingUtilities.invokeLater(() -> {
                progressBar.setValue(progressBar.getMaximum());
                log("BulkMail completed!");
                sendBulkMailButton.setEnabled(true);
            });
        }).start();
    }

    // ---------------- STOP BULKMAIL ----------------
    private void stopBulkMail() {
        stopFlag.set(true);
        log("Stop requested...");
    }

    // ---------------- LOGGING ----------------
    public void log(String message) {
        SwingUtilities.invokeLater(() -> logArea.append(message + "\n"));
    }

    // ---------------- MAIN ----------------
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ControlPanel cp = new ControlPanel();
            cp.setVisible(true);
        });
    }
}