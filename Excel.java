import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {

    public static void main(String[] args) {
        // Create a new workbook
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create a sheet
            Sheet sheet = workbook.createSheet("Sheet1");

            // Data to be written to Excel
            Object[][] data = {
                    {"Name", "Age", "Location"},
                    {"John", 30, "New York"},
                    {"Alice", 25, "London"},
                    {"Bob", 35, "Paris"}
            };

            // Write data to cells
            int rowNum = 0;
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
                System.out.println("Excel file has been created successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
