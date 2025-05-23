package dev.lescoggi;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.UUID;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import io.quarkus.test.junit.QuarkusTest;

@QuarkusTest
public class ExcelFeaturesTest {

    private File tempDir;
    private OfficeMcpServerExcelFeatures excelFeatures;
    private String workbookPath;
    private final String SHEET_NAME = "TestSheet";

    @BeforeEach
    void setUp() throws Exception {
        // Create a temporary directory manually
        String tempDirPath = System.getProperty("java.io.tmpdir");
        tempDir = new File(tempDirPath, "excel-test-" + UUID.randomUUID());
        tempDir.mkdirs();
        
        excelFeatures = new OfficeMcpServerExcelFeatures();
        workbookPath = tempDir.getAbsolutePath() + "/test.xlsx";
    }
    
    @AfterEach
    void tearDown() {
        // Clean up the test files
        if (tempDir != null && tempDir.exists()) {
            File[] files = tempDir.listFiles();
            if (files != null) {
                for (File file : files) {
                    file.delete();
                }
            }
            tempDir.delete();
        }
    }

    @Test
    void testCreateExcelWorkbook() throws Exception {
        excelFeatures.createExcelWorkbook(workbookPath);
        assertTrue(Files.exists(Paths.get(workbookPath)));
    }

    @Test
    void testCreateExcelSheet() throws Exception {
        // First create the workbook
        excelFeatures.createExcelWorkbook(workbookPath);
        
        // Add a new sheet
        assertNotNull(excelFeatures.createExcelSheet(workbookPath, SHEET_NAME));
        
        // Verify the file exists
        assertTrue(Files.exists(Paths.get(workbookPath)));
    }
    
    @Test
    void testAddExcelRow() throws Exception {
        // Create workbook and add data
        excelFeatures.createExcelWorkbook(workbookPath);
        String rowData = "Cell1,Cell2,Cell3";
        
        // Add a row (to Sheet1 which is created by default)
        assertNotNull(excelFeatures.addExcelRow(workbookPath, "Sheet1", rowData));
    }
    
    @Test
    void testReadExcelCell() throws Exception {
        // Create workbook and add data
        excelFeatures.createExcelWorkbook(workbookPath);
        String rowData = "TestValue";
        excelFeatures.addExcelRow(workbookPath, "Sheet1", rowData);
        
        // Read the cell
        assertNotNull(excelFeatures.readExcelCell(workbookPath, "Sheet1", 0, 0));
    }
    
    @Test
    void testCloseExcelWorkbook() throws Exception {
        // Create workbook
        excelFeatures.createExcelWorkbook(workbookPath);
        
        // Close the workbook
        assertNotNull(excelFeatures.closeExcelWorkbook(workbookPath));
    }
    
    @Test
    void testGetExcelSheetCount() throws Exception {
        // Create workbook with multiple sheets
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.createExcelSheet(workbookPath, "Sheet2");
        excelFeatures.createExcelSheet(workbookPath, "Sheet3");
        
        // Get sheet count
        assertNotNull(excelFeatures.getExcelSheetCount(workbookPath));
    }
    
    @Test
    void testGetExcelRowCount() throws Exception {
        // Create workbook and add rows
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.addExcelRow(workbookPath, "Sheet1", "Row1");
        excelFeatures.addExcelRow(workbookPath, "Sheet1", "Row2");
        
        // Get row count
        assertNotNull(excelFeatures.getExcelRowCount(workbookPath, "Sheet1"));
    }
    
    @Test
    void testGetExcelColumnCount() throws Exception {
        // Create workbook and add row with multiple columns
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.addExcelRow(workbookPath, "Sheet1", "Col1,Col2,Col3,Col4");
        
        // Get column count
        assertNotNull(excelFeatures.getExcelColumnCount(workbookPath, "Sheet1"));
    }
}