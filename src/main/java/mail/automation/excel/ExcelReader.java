package mail.automation.excel;

import mail.automation.model.MailData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.InputStream;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ExcelReader {

    private DataFormatter formatter = new DataFormatter();

    /**
     * Let user pick an Excel file at runtime
     */
    public String pickExcelFile() {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Select Excel File");
        int result = chooser.showOpenDialog(null);

        if (result == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();
            return file.getAbsolutePath();
        } else {
            System.out.println("No file selected!");
            return null;
        }
    }

    public List<MailData> readBulkMail(String filePath) {
        List<MailData> list = new ArrayList<>();
        if (filePath == null)
            return list;

        try (InputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("BulkMail");
            if (sheet == null)
                return list;

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                if (row == null)
                    continue;

                MailData mail = new MailData();

                mail.setTo(getCellValue(row.getCell(0))); // A
                mail.setCc(getCellValue(row.getCell(1))); // B
                mail.setSubject(getCellValueWithFormula(row.getCell(6), workbook)); // G
                mail.setBody(getCellValue(row.getCell(7))); // H
                mail.setAttachment1(getCellValue(row.getCell(8))); // I
                mail.setAttachment2(getCellValue(row.getCell(9))); // J
                mail.setFilterValue(getCellValue(row.getCell(2))); // Column C
                if (mail.getTo() != null && !mail.getTo().isEmpty()) {
                    list.add(mail);

                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return list;

    }

    private String getCellValue(Cell cell) {
        if (cell == null)
            return "";

        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell).trim();
    }

    // Subject Getter
    private String getCellValueWithFormula(Cell cell, Workbook workbook) {
        if (cell == null)
            return "";

        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        return formatter.formatCellValue(cell, evaluator).trim();
    }

    private boolean isRowEmpty(Row row) {
        if (row == null)
            return true;

        DataFormatter formatter = new DataFormatter();

        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);

            if (cell != null &&
                    cell.getCellType() != CellType.BLANK &&
                    !formatter.formatCellValue(cell).trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }

    // Build Body
    public String buildBodyContent(String filePath) {

        StringBuilder html = new StringBuilder();
        DataFormatter formatter = new DataFormatter();

        try (FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("GSTMailBody");

            String body1 = formatter.formatCellValue(sheet.getRow(4).getCell(1));
            String body2 = formatter.formatCellValue(sheet.getRow(5).getCell(1));

            html.append("<div style='font-family:Calibri;font-size:11pt;'>");

            html.append(body1).append("<br>");
            html.append(body2).append("<br><br>");
            html.append("</div>"); // ✅ THIS WAS MISSING

        } catch (Exception e) {
            e.printStackTrace();
        }

        return html.toString();
    }

    // Read GST Sheet For Table
    // private DataFormatter formatter = new DataFormatter();

    // Read the GST sheet into a list of maps (columnName -> value)
    public List<Map<String, String>> readGstSheet(String filePath, String sheetName) {
        List<Map<String, String>> rows = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null)
                return rows;

            // Read header
            Row headerRow = sheet.getRow(0);
            if (headerRow == null)
                return rows;
            int lastCol = headerRow.getLastCellNum();
            List<Integer> visibleCols = new ArrayList<>();
            List<String> headerNames = new ArrayList<>();

            for (int col = 0; col < lastCol; col++) {
                if (!sheet.isColumnHidden(col)) {
                    visibleCols.add(col);
                    headerNames.add(formatter.formatCellValue(headerRow.getCell(col)));
                }
            }

            int lastRow = sheet.getLastRowNum();
            for (int rowNum = 1; rowNum <= lastRow; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null || row.getZeroHeight())
                    continue;
                if (isRowEmpty(row))
                    continue;

                Map<String, String> rowData = new LinkedHashMap<>();
                for (int i = 0; i < visibleCols.size(); i++) {
                    int colIndex = visibleCols.get(i);
                    Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowData.put(headerNames.get(i), formatter.formatCellValue(cell));
                }
                rows.add(rowData);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return rows;
    }

    // Convert Filtered Rows To Html Table
    public String convertRowsToHtmlTable(List<Map<String, String>> rows, String filterValue, String gstColumnName) {

        StringBuilder html = new StringBuilder();
        if (rows.isEmpty())
            return "";

        // Normalize column name
        gstColumnName = gstColumnName.trim();

        // Build header
        html.append("<table border='1' style='border-collapse:collapse;font-family:Calibri;font-size:10pt;'>");
        html.append("<tr style='background-color:#ED7D31;color:white;'>");

        for (String header : rows.get(0).keySet()) {
            html.append("<th>").append(header).append("</th>");
        }
        html.append("</tr>");

        // Normalize filter
        String filter = (filterValue == null) ? "" : filterValue;
        filter = filter.trim().replace(".0", "").replaceAll("\\s+", "");

        for (Map<String, String> row : rows) {

            // 🔥 PRINT HEADERS (IMPORTANT DEBUG)
            System.out.println("Available Columns: " + row.keySet());

            String gstValue = row.get(gstColumnName);

            // ❗ If column mismatch
            if (gstValue == null) {
                System.out.println("❌ Column NOT FOUND: " + gstColumnName);
                continue;
            }

            // Normalize GST value
            gstValue = gstValue.trim().replace(".0", "").replaceAll("\\s+", "");

            // Debug
            // System.out.println("GST COLUMN VALUE: [" + gstValue + "]");
            // System.out.println("FILTER VALUE: [" + filter + "]");
            // System.out.println("MATCH: " + gstValue.equalsIgnoreCase(filter));
            // System.out.println("----------------------");

            if (!gstValue.equalsIgnoreCase(filter))
                continue;

            // Build row
            html.append("<tr>");
            for (String value : row.values()) {
                html.append("<td>").append(value).append("</td>");
            }
            html.append("</tr>");
        }

        html.append("</table>");
        return html.toString();
    }

    // Signature With Logo
    public String buildSignature(String filePath) {

        StringBuilder html = new StringBuilder();
        DataFormatter formatter = new DataFormatter();

        try (FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("GSTMailBody");

            String[] sig = new String[12];

            // Safe read (avoids null pointer)
            for (int i = 0; i < 10; i++) {
                Row row = sheet.getRow(9 + i);
                sig[i] = (row != null && row.getCell(1) != null)
                        ? formatter.formatCellValue(row.getCell(1))
                        : "";
            }

            Row row20 = sheet.getRow(20);
            sig[10] = (row20 != null && row20.getCell(1) != null)
                    ? formatter.formatCellValue(row20.getCell(1))
                    : "";

            // ================= HTML =================
            html.append("<table style='width:100%; font-family:Calibri; font-size:11pt;'>");
            html.append("<tr>");

            // ===== LEFT SIDE =====
            html.append("<td style='width:120px; vertical-align:top;'>")

                    // B10 ABOVE LOGO
                    .append("<div style='font-weight:bold; margin-bottom:5px;'>")
                    .append(sig[0])
                    .append("</div>")

                    // LOGO1 (CID)
                    .append("<img src='cid:logo1' width='90' style='display:block; margin-top:-5px;'/>")

                    .append("</td>");

            // ===== RIGHT SIDE =====
            html.append("<td style='vertical-align:top; padding-left:10px;'>")

                    .append("<br>")

                    .append("<div>").append(sig[1]).append("</div>")
                    .append("<div style='font-weight:bold;'>").append(sig[2]).append("</div>")
                    .append("<div style='color:#F7941E;'>").append(sig[3]).append("</div>");

            for (int i = 4; i <= 9; i++) {
                html.append("<div>").append(sig[i]).append("</div>");
            }

            html.append("<br>")
                    .append("<div style='color:#F7941E;'>").append(sig[10]).append("</div>")

                    .append("<br>")

                    // LOGO2 (CID)
                    .append("<img src='cid:logo2' width='120'/>")

                    .append("</td>");

            html.append("</tr>");
            html.append("</table>");

            // ❌ DO NOT close body/html here
            // html.append("</body></html>");

        } catch (Exception e) {
            e.printStackTrace();
        }

        return html.toString();
    }

    // ================= LOGO PATH READER =================
    public String[] getLogoPaths(String filePath) {
        String[] logos = new String[2];

        try (FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet("GSTMailBody");

            if (sheet != null) {

                // B22 -> row 21, col 1
                Row row21 = sheet.getRow(21);
                logos[0] = (row21 != null && row21.getCell(1) != null)
                        ? formatter.formatCellValue(row21.getCell(1))
                        : "";

                // B23 -> row 22, col 1
                Row row22 = sheet.getRow(22);
                logos[1] = (row22 != null && row22.getCell(1) != null)
                        ? formatter.formatCellValue(row22.getCell(1))
                        : "";
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return logos;
    }
}