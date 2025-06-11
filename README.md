# Office MCP Server

The **Office MCP Server** is an unofficial [Model Context Protocol (MCP) Server](https://modelcontextprotocol.io) based Java server designed to manage and process requests from AI agents for Word, Excel, etc. files.

## Framework

This project is built using **Quarkus**, a Kubernetes-native Java framework tailored for building lightweight, high-performance microservices. Quarkus enables fast startup times and low memory usage, making it ideal for cloud-native applications.

## Supported Methods

The server supports the following MCP tools and resources:

### Excel Features

- **Tool: Get Filename**: Retrieve the filename of an Excel file.
  - **Argument**: `filename` - The name of the Excel file.
- **Tool: Create Excel Workbook**: Create a new Excel workbook.
  - **Argument**: `filepath` - Path to create the new Excel workbook.
- **Tool: Create Excel Sheet**: Create a new sheet in an Excel workbook.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the new sheet.
- **Tool: Add Excel Row**: Add a row to an Excel sheet.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the sheet.
    - `rowData` - Data for the new row.
- **Tool: Read Excel Cell**: Read a cell from an Excel sheet.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the sheet.
    - `rowNum` - Row number (0-based).
    - `colNum` - Column number (0-based).
- **Tool: Close Excel Workbook**: Close an Excel workbook.
  - **Argument**: `filepath` - Path to the Excel workbook.
- **Tool: Get Excel Sheet Count**: Get the number of sheets in an Excel workbook.
  - **Argument**: `filepath` - Path to the Excel workbook.
- **Tool: Get Excel Row Count**: Get the number of rows in a sheet.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the sheet.
- **Tool: Get Excel Column Count**: Get the number of columns in a sheet.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the sheet.

#### ✨ NEW: Colorful Table Formatting Features

- **Tool: Create Formatted Table**: Create a new table with colorful styling and formatting.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the sheet.
    - `tableData` - Table data as comma-separated values, rows separated by semicolons.
    - `theme` - Color theme: `blue`, `green`, `orange` (default: blue).
    - `alternatingRows` - Whether to enable alternating row colors (true/false).
  
- **Tool: Apply Table Formatting**: Apply colorful formatting to existing data range.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the sheet.
    - `tableRange` - Table range (e.g., A1:D10).
    - `colorTheme` - Color theme: `blue`, `green`, `orange`.
    - `alternatingRows` - Whether to enable alternating row colors.

- **Tool: Apply Conditional Formatting**: Highlight cells based on values or conditions.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the sheet.
    - `cellRange` - Cell range to apply formatting (e.g., A1:D10).
    - `conditionType` - Condition type: `greater_than`, `less_than`, `equal_to`, `between`.
    - `conditionValue` - Value(s) for condition (comma-separated for 'between').
    - `highlightColor` - Highlight color: `red`, `green`, `yellow`, `blue`.

- **Tool: Apply Custom Borders**: Apply custom border styles and colors to table ranges.
  - **Arguments**:
    - `filepath` - Path to the Excel workbook.
    - `sheetName` - Name of the sheet.
    - `cellRange` - Cell range (e.g., A1:D10).
    - `borderStyle` - Border style: `thin`, `medium`, `thick`, `double`.
    - `borderColor` - Border color: `black`, `blue`, `red`, `green`.

#### Color Themes

The server supports three predefined color themes for table formatting:

1. **Blue Theme**:
   - Header: Professional blue (#4472C4) with white text
   - Alternating rows: Light blue (#D9E1F2) and white
   - Borders: Blue (#4472C4)

2. **Green Theme**:
   - Header: Professional green (#70AD47) with white text
   - Alternating rows: Light green (#E2EFDA) and white
   - Borders: Green (#70AD47)

3. **Orange Theme**:
   - Header: Vibrant orange (#FFC000) with black text
   - Alternating rows: Light orange (#FFF2CC) and white
   - Borders: Orange (#FFC000)

#### Formatting Examples

**Create a formatted table with blue theme and alternating rows:**
```
Table Data: "Name,Age,Department;John,30,IT;Jane,25,HR;Bob,35,Finance"
Theme: "blue"
Alternating Rows: true
```

**Apply conditional formatting to highlight high values:**
```
Cell Range: "B2:B10"
Condition: "greater_than"
Value: "50000"
Highlight Color: "green"
```

**Apply custom thick red borders to a table:**
```
Cell Range: "A1:D5"
Border Style: "thick"
Border Color: "red"
```

### Word Features

- **Tool: Create Word Document**: Create a new Word document.
  - **Argument**: `filepath` - Path to create the new Word document.
- **Tool: Add Text to Word Document**: Add text to a Word document.
  - **Arguments**:
    - `filepath` - Path to the Word document.
    - `text` - Text to add.

### PowerPoint Features

- **Tool: Create PowerPoint Presentation**: Create a new PowerPoint presentation.
  - **Argument**: `filepath` - Path to create the new PowerPoint presentation.
- **Tool: Add Slide to PowerPoint**: Add a new slide to a PowerPoint presentation.
  - **Argument**: `filepath` - Path to the PowerPoint presentation.
- **Tool: Add Text to PowerPoint Slide**: Add text to a specific slide in a PowerPoint presentation.
  - **Arguments**:
    - `filepath` - Path to the PowerPoint presentation.
    - `slideIndex` - Slide index (0-based).
    - `text` - Text to add.
- **Tool: Read Slide Titles from PowerPoint**: Read titles from all slides in a PowerPoint presentation.
  - **Argument**: `filepath` - Path to the PowerPoint presentation.
- **Tool: Get PowerPoint Slide Count**: Get the number of slides in a PowerPoint presentation.
  - **Argument**: `filepath` - Path to the PowerPoint presentation.

## How to Debug and Run Standalone

To run the Office MCP Server, follow these steps:

1. **Build the application**:
   Ensure you have Maven installed. Run the following command to build the project:
   ```sh
   $ ./mvnw clean package
   ```

2. **Run the application**:
   After building, execute the following command to start the server:
   ```sh
   $ java -jar target/office-mcp-server-0.0.1-SNAPSHOT-runner.jar
   ```

3. **Access the application**:
   The server will start on the default port `8080`. You can access the APIs at:
   ```
   http://localhost:8080
   ```

4. **Run in development mode**:
   For development purposes, you can use Quarkus's dev mode:
   ```sh
   $ ./mvnw quarkus:dev
   ```

## Configuration

The application can be configured using the `application.properties` file located in the `src/main/resources` directory. Key configuration options include:

- `quarkus.mcp.server.sse.root-path`: Set the SSE root path.

## Testing

To run the tests, use the following command:
```sh
$ ./mvnw test
```

## Installing the MCP Server

### Configuring the MCP server with VS Code

Here are the steps to configure in VS Code:

- Install GitHub Copilot
- Install this MCP Server using the command palette: `MCP: Add Server...`
- Configure GitHub Copilot to run in `Agent` mode, by clicking on the arrow at the bottom of the the chat window
- On top of the chat window, you should see the `office-mcp-server` server configured as a tool

### Configuring the MCP server with Claude Desktop

Claude Desktop makes it easy to configure and chat with the MCP server. If you want a more advanced usage, we recommend using VS Code (see next section).

You need to add the server to your `claude_desktop_config.json` file. Please note that you need to point to the location
where you downloaded the `office-mcp-server-0.0.1-SNAPSHOT-runner.jar` file.

```json
{
    "mcpServers": {
        "office-mcp-server": {
            "command": "java",
            "args": [
                "-jar",
              "~/Downloads/office-mcp-server-0.0.1-SNAPSHOT-runner.jar"
            ]
        }
    }
}
```

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Support

This project is provided as-is without any warranty. If you encounter issues or have questions, please open an issue on the GitHub repository.