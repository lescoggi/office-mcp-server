package dev.lescoggi;

import org.eclipse.microprofile.config.inject.ConfigProperty;

import io.quarkiverse.mcp.server.TextContent;
import io.quarkiverse.mcp.server.Tool;
import io.quarkiverse.mcp.server.ToolArg;
import io.quarkiverse.mcp.server.ToolResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class OfficeMcpServerExcelFeatures {

    @ConfigProperty(name = "office.files.path")
    String officeFilesPath;

    @Tool(description = "Create a new Excel workbook", name = "create_excel_workbook")
    ToolResponse createExcelWorkbook(@ToolArg(description = "Path to create new Excel workbook") String filepath) {
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOut = new FileOutputStream(filepath)) {
            workbook.createSheet("Sheet1");
            workbook.write(fileOut);
            return ToolResponse.success(
                new TextContent("Excel workbook created at: " + filepath));
        } catch (IOException e) {
            return ToolResponse.error("Failed to create Excel workbook: " + e.getMessage());
        }
    }

    @Tool(description = "Create a new sheet in an Excel workbook", name = "create_excel_sheet")
    ToolResponse createExcelSheet(@ToolArg(description = "Path to the Excel workbook") String filepath,
                                   @ToolArg(description = "Name of the new sheet") String sheetName) {
        try (FileInputStream fileIn = new FileInputStream(filepath); Workbook workbook = new XSSFWorkbook(fileIn)) {
            workbook.createSheet(sheetName);
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                workbook.write(fileOut);
            }
            return ToolResponse.success(
                new TextContent("Sheet '" + sheetName + "' created in workbook: " + filepath));
        } catch (IOException e) {
            return ToolResponse.error("Failed to create sheet: " + e.getMessage());
        }
    }

    @Tool(description = "Add a row to an Excel sheet", name = "add_excel_row")
    ToolResponse addExcelRow(@ToolArg(description = "Path to the Excel workbook") String filepath,
                              @ToolArg(description = "Name of the sheet") String sheetName,
                              @ToolArg(description = "Row data") String rowData) {
        try (FileInputStream fileIn = new FileInputStream(filepath); Workbook workbook = new XSSFWorkbook(fileIn)) {
            var sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return ToolResponse.error("Sheet '" + sheetName + "' does not exist.");
            }
            var row = sheet.createRow(sheet.getLastRowNum() + 1);
            if (rowData.contains(",")) {
                String[] cellValues = rowData.split(",");
                for (int i = 0; i < cellValues.length; i++) {
                    row.createCell(i).setCellValue(cellValues[i]);
                }
            } else {
                row.createCell(0).setCellValue(rowData);
            }
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                workbook.write(fileOut);
            }
            return ToolResponse.success(
                new TextContent("Row added to sheet '" + sheetName + "' in workbook: " + filepath));
        } catch (IOException e) {
            return ToolResponse.error("Failed to add row: " + e.getMessage());
        }
    }

    @Tool(description = "Read a cell from an Excel sheet", name = "read_excel_cell")
    ToolResponse readExcelCell(@ToolArg(description = "Path to the Excel workbook") String filepath,
                                @ToolArg(description = "Name of the sheet") String sheetName,
                                @ToolArg(description = "Row number") int rowNum,
                                @ToolArg(description = "Column number") int colNum) {
        try (Workbook workbook = new XSSFWorkbook(filepath)) {
            var sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return ToolResponse.error("Sheet '" + sheetName + "' does not exist.");
            }
            var row = sheet.getRow(rowNum);
            if (row == null) {
                return ToolResponse.error("Row " + rowNum + " does not exist in sheet '" + sheetName + "'.");
            }
            var cell = row.getCell(colNum);
            if (cell == null) {
                return ToolResponse.error("Cell (" + rowNum + ", " + colNum + ") does not exist in sheet '" + sheetName + "'.");
            }
            return ToolResponse.success(new TextContent(cell.toString()));
        } catch (IOException e) {
            return ToolResponse.error("Failed to read cell: " + e.getMessage());
        }
    }

    @Tool(description = "Close an Excel workbook", name = "close_excel_workbook")
    ToolResponse closeExcelWorkbook(@ToolArg(description = "Path to the Excel workbook") String filepath) {
        try (Workbook workbook = new XSSFWorkbook(filepath)) {
            workbook.close();
            return ToolResponse.success(
                new TextContent("Excel workbook closed: " + filepath));
        } catch (IOException e) {
            return ToolResponse.error("Failed to close Excel workbook: " + e.getMessage());
        }
    }

    @Tool(description = "Get the number of sheets in an Excel workbook", name = "get_excel_sheet_count")
    ToolResponse getExcelSheetCount(@ToolArg(description = "Path to the Excel workbook") String filepath) {
        try (Workbook workbook = new XSSFWorkbook(filepath)) {
            int sheetCount = workbook.getNumberOfSheets();
            return ToolResponse.success("Workbook has " + sheetCount + " sheets.");
        } catch (IOException e) {
            return ToolResponse.error("Failed to get sheet count: " + e.getMessage());
        }
    }

    @Tool(description = "Get the number of rows in a sheet", name = "get_excel_row_count")
    ToolResponse getExcelRowCount(@ToolArg(description = "Path to the Excel workbook") String filepath,
                                    @ToolArg(description = "Name of the sheet") String sheetName) {
        try (Workbook workbook = new XSSFWorkbook(filepath)) {
            var sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return ToolResponse.error("Sheet '" + sheetName + "' does not exist.");
            }
            int rowCount = sheet.getPhysicalNumberOfRows();
            return ToolResponse.success("Sheet has " + rowCount + " rows.");
        } catch (IOException e) {
            return ToolResponse.error("Failed to get row count: " + e.getMessage());
        }
    }

    @Tool(description = "Get the number of columns in a sheet", name = "get_excel_column_count")
    ToolResponse getExcelColumnCount(@ToolArg(description = "Path to the Excel workbook") String filepath,
                                       @ToolArg(description = "Name of the sheet") String sheetName) {
        try (Workbook workbook = new XSSFWorkbook(filepath)) {
            var sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return ToolResponse.error("Sheet '" + sheetName + "' does not exist.");
            }
            int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
            return ToolResponse.success("Sheet has " + columnCount + " columns.");
        } catch (IOException e) {
            return ToolResponse.error("Failed to get column count: " + e.getMessage());
        }
    }

}