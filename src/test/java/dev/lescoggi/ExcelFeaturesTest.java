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

    @Test
    void testCreateFormattedTable() throws Exception {
        // First create a workbook and sheet
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.createExcelSheet(workbookPath, SHEET_NAME);
        
        // Test creating a formatted table with blue theme and alternating rows
        String tableData = "Name,Age,Department;John,30,IT;Jane,25,HR;Bob,35,Finance;Alice,28,Marketing";
        var response = excelFeatures.createFormattedTable(workbookPath, SHEET_NAME, tableData, "blue", true);
        
        assertNotNull(response);
        assertTrue(Files.exists(Paths.get(workbookPath)));
    }

    @Test
    void testCreateFormattedTableWithDifferentThemes() throws Exception {
        // First create a workbook and sheet
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.createExcelSheet(workbookPath, SHEET_NAME);
        
        String tableData = "Product,Price,Stock;Laptop,999.99,50;Mouse,29.99,100;Keyboard,79.99,75";
        
        // Test green theme
        var greenResponse = excelFeatures.createFormattedTable(workbookPath, SHEET_NAME, tableData, "green", true);
        assertNotNull(greenResponse);
        
        // Test orange theme  
        var orangeResponse = excelFeatures.createFormattedTable(workbookPath, SHEET_NAME, tableData, "orange", false);
        assertNotNull(orangeResponse);
    }

    @Test
    void testApplyTableFormattingToExistingData() throws Exception {
        // Create workbook and add some data first
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.createExcelSheet(workbookPath, SHEET_NAME);
        
        // Add header row
        excelFeatures.addExcelRow(workbookPath, SHEET_NAME, "Column A,Column B,Column C");
        excelFeatures.addExcelRow(workbookPath, SHEET_NAME, "Data1,Data2,Data3");
        excelFeatures.addExcelRow(workbookPath, SHEET_NAME, "Data4,Data5,Data6");
        
        // Apply table formatting to the range
        var response = excelFeatures.applyTableFormatting(workbookPath, SHEET_NAME, "A0:C2", "blue", true);
        assertNotNull(response);
    }

    @Test
    void testApplyConditionalFormattingFeature() throws Exception {
        // Create workbook and add numeric data
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.createExcelSheet(workbookPath, SHEET_NAME);
        
        excelFeatures.addExcelRow(workbookPath, SHEET_NAME, "10,20,30");
        excelFeatures.addExcelRow(workbookPath, SHEET_NAME, "40,50,60");
        excelFeatures.addExcelRow(workbookPath, SHEET_NAME, "70,80,90");
        
        // Apply conditional formatting to highlight values greater than 50
        var response = excelFeatures.applyConditionalFormatting(
            workbookPath, SHEET_NAME, "A0:C2", "greater_than", "50", "red");
        assertNotNull(response);
    }

    @Test
    void testApplyCustomBordersFeature() throws Exception {
        // Create workbook and add some data
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.createExcelSheet(workbookPath, SHEET_NAME);
        
        excelFeatures.addExcelRow(workbookPath, SHEET_NAME, "Border,Test,Data");
        excelFeatures.addExcelRow(workbookPath, SHEET_NAME, "With,Custom,Styling");
        
        // Apply thick blue borders
        var response = excelFeatures.applyCustomBorders(
            workbookPath, SHEET_NAME, "A0:C1", "thick", "blue");
        assertNotNull(response);
    }

    @Test
    void testFormattedTableWithoutAlternatingRows() throws Exception {
        // Create workbook and sheet
        excelFeatures.createExcelWorkbook(workbookPath);
        excelFeatures.createExcelSheet(workbookPath, SHEET_NAME);
        
        String tableData = "Header1,Header2;Value1,Value2;Value3,Value4";
        
        // Create table without alternating rows
        var response = excelFeatures.createFormattedTable(workbookPath, SHEET_NAME, tableData, "green", false);
        assertNotNull(response);
    }
}