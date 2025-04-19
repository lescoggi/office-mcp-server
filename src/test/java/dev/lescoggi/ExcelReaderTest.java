package dev.lescoggi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class ExcelReaderTest {

    @Test
    public void testReadExcelFile() throws IOException {
        // Load the file from the classpath
        ClassLoader classLoader = getClass().getClassLoader();
        File excelFile = new File(classLoader.getResource("example.xlsx").getFile());
        try (FileInputStream fis = new FileInputStream(excelFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Access the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Read the first row and first cell
            Row row = sheet.getRow(0);
            Cell cell = row.getCell(0);

            // Assert the value of the first cell
            String expectedValue = "Name";
            assertEquals(expectedValue, cell.getStringCellValue());
        }
    }
}
