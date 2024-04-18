import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {

    public static void main(String[] args) {
        // Create a new Excel workbook
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create a new sheet with the name "Sheet1"
            Sheet sheet = workbook.createSheet("Sheet1");

            // Write column headers
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Name", "Age", "Email"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Write data rows
            Object[][] data = {
                    {"John Doe", 30, "john@test.com"},
                    {"Jane Doe", 28, "jane@test.com"},
                    {"Bob Smith", 35, "bob@test.com"},
                    {"Swapnil", 37, "swapnil@example.com"}
            };
            int rowNum = 1;
            for (Object[] rowData : data) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (Object field : rowData) {
                    Cell cell = row.createCell(colNum++);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
            }

            // Write the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")) {
                workbook.write(fileOut);
                System.out.println("Excel file has been generated successfully!");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
