package dev.lescoggi;

import org.eclipse.microprofile.config.inject.ConfigProperty;

import io.quarkiverse.mcp.server.TextContent;
import io.quarkiverse.mcp.server.Tool;
import io.quarkiverse.mcp.server.ToolArg;
import io.quarkiverse.mcp.server.ToolResponse;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.HashMap;

public class OfficeMcpServerExcelFeatures {

    @ConfigProperty(name = "office.files.path")
    String officeFilesPath;
    
    // Color theme definitions
    private static final Map<String, ColorTheme> COLOR_THEMES = new HashMap<>();
    
    static {
        // Blue theme
        COLOR_THEMES.put("blue", new ColorTheme(
            new byte[]{(byte)68, (byte)114, (byte)196}, // Header: #4472C4
            new byte[]{(byte)217, (byte)225, (byte)242}, // Alt row 1: #D9E1F2
            new byte[]{(byte)255, (byte)255, (byte)255}, // Alt row 2: #FFFFFF
            new byte[]{(byte)68, (byte)114, (byte)196}   // Border: #4472C4
        ));
        
        // Green theme
        COLOR_THEMES.put("green", new ColorTheme(
            new byte[]{(byte)112, (byte)173, (byte)71}, // Header: #70AD47
            new byte[]{(byte)226, (byte)239, (byte)218}, // Alt row 1: #E2EFDA
            new byte[]{(byte)255, (byte)255, (byte)255}, // Alt row 2: #FFFFFF
            new byte[]{(byte)112, (byte)173, (byte)71}   // Border: #70AD47
        ));
        
        // Orange theme
        COLOR_THEMES.put("orange", new ColorTheme(
            new byte[]{(byte)255, (byte)192, (byte)0}, // Header: #FFC000
            new byte[]{(byte)255, (byte)242, (byte)204}, // Alt row 1: #FFF2CC
            new byte[]{(byte)255, (byte)255, (byte)255}, // Alt row 2: #FFFFFF
            new byte[]{(byte)255, (byte)192, (byte)0}   // Border: #FFC000
        ));
    }
    
    private static class ColorTheme {
        byte[] headerColor;
        byte[] altRow1Color;
        byte[] altRow2Color;
        byte[] borderColor;
        
        ColorTheme(byte[] headerColor, byte[] altRow1Color, byte[] altRow2Color, byte[] borderColor) {
            this.headerColor = headerColor;
            this.altRow1Color = altRow1Color;
            this.altRow2Color = altRow2Color;
            this.borderColor = borderColor;
        }
    }

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

    @Tool(description = "Create a formatted table with colorful styling", name = "create_formatted_table")
    ToolResponse createFormattedTable(@ToolArg(description = "Path to the Excel workbook") String filepath,
                                      @ToolArg(description = "Name of the sheet") String sheetName,
                                      @ToolArg(description = "Table data as comma-separated values, rows separated by semicolons") String tableData,
                                      @ToolArg(description = "Color theme: blue, green, orange, or custom") String theme,
                                      @ToolArg(description = "Whether to enable alternating row colors (true/false)") boolean alternatingRows) {
        try (FileInputStream fileIn = new FileInputStream(filepath); 
             XSSFWorkbook workbook = new XSSFWorkbook(fileIn)) {
            
            var sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return ToolResponse.error("Sheet '" + sheetName + "' does not exist.");
            }

            ColorTheme colorTheme = COLOR_THEMES.getOrDefault(theme.toLowerCase(), COLOR_THEMES.get("blue"));
            
            // Parse table data
            String[] rows = tableData.split(";");
            if (rows.length == 0) {
                return ToolResponse.error("No table data provided.");
            }

            int startRow = sheet.getLastRowNum() + 1;
            
            // Create styles
            XSSFCellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(new XSSFColor(colorTheme.headerColor, null));
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            setBordersAndColors(headerStyle, colorTheme);

            XSSFCellStyle altRow1Style = null;
            XSSFCellStyle altRow2Style = null;
            
            if (alternatingRows) {
                altRow1Style = workbook.createCellStyle();
                altRow1Style.setFillForegroundColor(new XSSFColor(colorTheme.altRow1Color, null));
                altRow1Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                setBordersAndColors(altRow1Style, colorTheme);

                altRow2Style = workbook.createCellStyle();
                altRow2Style.setFillForegroundColor(new XSSFColor(colorTheme.altRow2Color, null));
                altRow2Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                setBordersAndColors(altRow2Style, colorTheme);
            }

            // Add rows
            for (int i = 0; i < rows.length; i++) {
                String[] cells = rows[i].split(",");
                Row row = sheet.createRow(startRow + i);
                
                for (int j = 0; j < cells.length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(cells[j].trim());
                    
                    if (i == 0) {
                        // Header row
                        cell.setCellStyle(headerStyle);
                    } else if (alternatingRows) {
                        // Alternating rows
                        cell.setCellStyle((i % 2 == 1) ? altRow1Style : altRow2Style);
                    }
                }
            }

            // Auto-size columns
            for (int i = 0; i < rows[0].split(",").length; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                workbook.write(fileOut);
            }

            return ToolResponse.success(
                new TextContent("Formatted table created in sheet '" + sheetName + "' with " + theme + " theme."));
        } catch (IOException e) {
            return ToolResponse.error("Failed to create formatted table: " + e.getMessage());
        }
    }
    
    private void setBordersAndColors(XSSFCellStyle style, ColorTheme theme) {
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        // For XSSF, we need to set border colors using the XSSF-specific methods
        style.setBottomBorderColor(new XSSFColor(theme.borderColor, null));
        style.setTopBorderColor(new XSSFColor(theme.borderColor, null));
        style.setLeftBorderColor(new XSSFColor(theme.borderColor, null));
        style.setRightBorderColor(new XSSFColor(theme.borderColor, null));
    }

    @Tool(description = "Apply table formatting to existing data range", name = "apply_table_formatting")
    ToolResponse applyTableFormatting(@ToolArg(description = "Path to the Excel workbook") String filepath,
                                      @ToolArg(description = "Name of the sheet") String sheetName,
                                      @ToolArg(description = "Table range (e.g., A1:D10)") String tableRange,
                                      @ToolArg(description = "Color theme (blue, green, orange)") String colorTheme,
                                      @ToolArg(description = "Whether to enable alternating row colors") boolean alternatingRows) {
        try (FileInputStream fileIn = new FileInputStream(filepath); 
             XSSFWorkbook workbook = new XSSFWorkbook(fileIn)) {
            
            var sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return ToolResponse.error("Sheet '" + sheetName + "' does not exist.");
            }
            
            // Parse the table range
            CellRangeAddress range = CellRangeAddress.valueOf(tableRange);
            
            // Get the color theme
            ColorTheme theme = COLOR_THEMES.get(colorTheme.toLowerCase());
            if (theme == null) {
                return ToolResponse.error("Invalid color theme: " + colorTheme + ". Available themes: blue, green, orange");
            }
            
            // Create styles
            XSSFCellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(new XSSFColor(theme.headerColor, null));
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            setBordersAndColors(headerStyle, theme);
            
            // Apply header style to first row
            var headerRow = sheet.getRow(range.getFirstRow());
            if (headerRow != null) {
                for (int i = range.getFirstColumn(); i <= range.getLastColumn(); i++) {
                    var cell = headerRow.getCell(i);
                    if (cell != null) {
                        cell.setCellStyle(headerStyle);
                    }
                }
            }
            
            // Create alternating row styles if requested
            if (alternatingRows) {
                XSSFCellStyle altRow1Style = workbook.createCellStyle();
                altRow1Style.setFillForegroundColor(new XSSFColor(theme.altRow1Color, null));
                altRow1Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                setBordersAndColors(altRow1Style, theme);

                XSSFCellStyle altRow2Style = workbook.createCellStyle();
                altRow2Style.setFillForegroundColor(new XSSFColor(theme.altRow2Color, null));
                altRow2Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                setBordersAndColors(altRow2Style, theme);
                
                // Apply alternating row colors
                for (int r = range.getFirstRow() + 1; r <= range.getLastRow(); r++) {
                    var row = sheet.getRow(r);
                    if (row != null) {
                        XSSFCellStyle style = (r % 2 == 0) ? altRow2Style : altRow1Style;
                        for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                            var cell = row.getCell(c);
                            if (cell != null) {
                                cell.setCellStyle(style);
                            }
                        }
                    }
                }
            } else {
                // Apply basic border style to data rows
                XSSFCellStyle borderStyle = workbook.createCellStyle();
                setBordersAndColors(borderStyle, theme);
                
                for (int r = range.getFirstRow() + 1; r <= range.getLastRow(); r++) {
                    var row = sheet.getRow(r);
                    if (row != null) {
                        for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                            var cell = row.getCell(c);
                            if (cell != null) {
                                cell.setCellStyle(borderStyle);
                            }
                        }
                    }
                }
            }
            
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                workbook.write(fileOut);
            }
            return ToolResponse.success(
                new TextContent("Table formatting applied to range " + tableRange + " in sheet '" + sheetName + "' with " + colorTheme + " theme."));
        } catch (IOException e) {
            return ToolResponse.error("Failed to apply table formatting: " + e.getMessage());
        }
    }

    @Tool(description = "Apply conditional formatting to highlight cells based on values", name = "apply_conditional_formatting")
    ToolResponse applyConditionalFormatting(@ToolArg(description = "Path to the Excel workbook") String filepath,
                                           @ToolArg(description = "Name of the sheet") String sheetName,
                                           @ToolArg(description = "Cell range to apply formatting (e.g., A1:D10)") String cellRange,
                                           @ToolArg(description = "Condition type: greater_than, less_than, equal_to, between") String conditionType,
                                           @ToolArg(description = "Value(s) for condition (comma-separated for 'between')") String conditionValue,
                                           @ToolArg(description = "Highlight color theme: red, green, yellow, blue") String highlightColor) {
        try (FileInputStream fileIn = new FileInputStream(filepath); 
             XSSFWorkbook workbook = new XSSFWorkbook(fileIn)) {
            
            var sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return ToolResponse.error("Sheet '" + sheetName + "' does not exist.");
            }
            
            CellRangeAddress range = CellRangeAddress.valueOf(cellRange);
            
            // Define highlight colors
            byte[] color;
            switch (highlightColor.toLowerCase()) {
                case "red":
                    color = new byte[]{(byte)255, (byte)199, (byte)206}; // Light red
                    break;
                case "green":
                    color = new byte[]{(byte)198, (byte)239, (byte)206}; // Light green
                    break;
                case "yellow":
                    color = new byte[]{(byte)255, (byte)235, (byte)156}; // Light yellow
                    break;
                case "blue":
                    color = new byte[]{(byte)180, (byte)198, (byte)231}; // Light blue
                    break;
                default:
                    return ToolResponse.error("Invalid highlight color: " + highlightColor + ". Available colors: red, green, yellow, blue");
            }
            
            XSSFCellStyle highlightStyle = workbook.createCellStyle();
            highlightStyle.setFillForegroundColor(new XSSFColor(color, null));
            highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            // Parse condition values
            String[] values = conditionValue.split(",");
            double value1 = Double.parseDouble(values[0].trim());
            double value2 = values.length > 1 ? Double.parseDouble(values[1].trim()) : 0;
            
            // Apply conditional formatting by checking each cell
            for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
                var row = sheet.getRow(r);
                if (row != null) {
                    for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                        var cell = row.getCell(c);
                        if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                            double cellValue = cell.getNumericCellValue();
                            boolean shouldHighlight = false;
                            
                            switch (conditionType.toLowerCase()) {
                                case "greater_than":
                                    shouldHighlight = cellValue > value1;
                                    break;
                                case "less_than":
                                    shouldHighlight = cellValue < value1;
                                    break;
                                case "equal_to":
                                    shouldHighlight = Math.abs(cellValue - value1) < 0.001;
                                    break;
                                case "between":
                                    shouldHighlight = cellValue >= Math.min(value1, value2) && cellValue <= Math.max(value1, value2);
                                    break;
                                default:
                                    return ToolResponse.error("Invalid condition type: " + conditionType + ". Available types: greater_than, less_than, equal_to, between");
                            }
                            
                            if (shouldHighlight) {
                                cell.setCellStyle(highlightStyle);
                            }
                        }
                    }
                }
            }
            
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                workbook.write(fileOut);
            }
            
            return ToolResponse.success(
                new TextContent("Conditional formatting applied to range " + cellRange + " in sheet '" + sheetName + "'."));
        } catch (IOException | NumberFormatException e) {
            return ToolResponse.error("Failed to apply conditional formatting: " + e.getMessage());
        }
    }

    @Tool(description = "Apply custom border styles to a table range", name = "apply_custom_borders")
    ToolResponse applyCustomBorders(@ToolArg(description = "Path to the Excel workbook") String filepath,
                                   @ToolArg(description = "Name of the sheet") String sheetName,
                                   @ToolArg(description = "Cell range (e.g., A1:D10)") String cellRange,
                                   @ToolArg(description = "Border style: thin, medium, thick, double") String borderStyle,
                                   @ToolArg(description = "Border color: black, blue, red, green") String borderColor) {
        try (FileInputStream fileIn = new FileInputStream(filepath); 
             XSSFWorkbook workbook = new XSSFWorkbook(fileIn)) {
            
            var sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return ToolResponse.error("Sheet '" + sheetName + "' does not exist.");
            }
            
            CellRangeAddress range = CellRangeAddress.valueOf(cellRange);
            
            // Define border styles
            BorderStyle style;
            switch (borderStyle.toLowerCase()) {
                case "thin":
                    style = BorderStyle.THIN;
                    break;
                case "medium":
                    style = BorderStyle.MEDIUM;
                    break;
                case "thick":
                    style = BorderStyle.THICK;
                    break;
                case "double":
                    style = BorderStyle.DOUBLE;
                    break;
                default:
                    return ToolResponse.error("Invalid border style: " + borderStyle + ". Available styles: thin, medium, thick, double");
            }
            
            // Define border colors
            byte[] color;
            switch (borderColor.toLowerCase()) {
                case "black":
                    color = new byte[]{0, 0, 0};
                    break;
                case "blue":
                    color = new byte[]{(byte)68, (byte)114, (byte)196};
                    break;
                case "red":
                    color = new byte[]{(byte)255, 0, 0};
                    break;
                case "green":
                    color = new byte[]{(byte)112, (byte)173, (byte)71};
                    break;
                default:
                    return ToolResponse.error("Invalid border color: " + borderColor + ". Available colors: black, blue, red, green");
            }
            
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setBorderBottom(style);
            cellStyle.setBorderTop(style);
            cellStyle.setBorderLeft(style);
            cellStyle.setBorderRight(style);
            cellStyle.setBottomBorderColor(new XSSFColor(color, null));
            cellStyle.setTopBorderColor(new XSSFColor(color, null));
            cellStyle.setLeftBorderColor(new XSSFColor(color, null));
            cellStyle.setRightBorderColor(new XSSFColor(color, null));
            
            // Apply to all cells in range
            for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
                var row = sheet.getRow(r);
                if (row != null) {
                    for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                        var cell = row.getCell(c);
                        if (cell != null) {
                            cell.setCellStyle(cellStyle);
                        }
                    }
                }
            }
            
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                workbook.write(fileOut);
            }
            
            return ToolResponse.success(
                new TextContent("Custom borders applied to range " + cellRange + " in sheet '" + sheetName + "'."));
        } catch (IOException e) {
            return ToolResponse.error("Failed to apply custom borders: " + e.getMessage());
        }
    }

}