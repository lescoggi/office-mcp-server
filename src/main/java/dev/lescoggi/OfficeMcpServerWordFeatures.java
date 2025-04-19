package dev.lescoggi;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.eclipse.microprofile.config.inject.ConfigProperty;

import io.quarkiverse.mcp.server.TextContent;
import io.quarkiverse.mcp.server.Tool;
import io.quarkiverse.mcp.server.ToolArg;
import io.quarkiverse.mcp.server.ToolResponse;

public class OfficeMcpServerWordFeatures {

    @ConfigProperty(name = "office.files.path")
    String officeFilesPath;

    @Tool(description = "Create a new Word document", name = "create_word_document")
    ToolResponse createWordDocument(@ToolArg(description = "Path to create new Word document") String filepath) {
        try (XWPFDocument document = new XWPFDocument(); FileOutputStream fileOut = new FileOutputStream(filepath)) {
            document.write(fileOut);
            return ToolResponse.success(
                new TextContent("Word document created at: " + filepath));
        } catch (Exception e) {
            return ToolResponse.error("Failed to create Word document: " + e.getMessage());
        }
    }

    @Tool(description = "Add text to a Word document", name = "add_text_to_word_document")
    ToolResponse addTextToWordDocument(@ToolArg(description = "Path to the Word document") String filepath,
                                        @ToolArg(description = "Text to add") String text) {
        try (FileInputStream fileIn = new FileInputStream(filepath); XWPFDocument document = new XWPFDocument(fileIn)) {
            document.createParagraph().createRun().setText(text);
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                document.write(fileOut);
            }
            return ToolResponse.success(
                new TextContent("Text added to Word document at: " + filepath));
        } catch (Exception e) {
            return ToolResponse.error("Failed to add text to Word document: " + e.getMessage());
        }
    }

    @Tool(description = "Read text from a Word document", name = "read_text_from_word_document")
    ToolResponse readTextFromWordDocument(@ToolArg(description = "Path to the Word document") String filepath) {
        try (FileInputStream fileIn = new FileInputStream(filepath); XWPFDocument document = new XWPFDocument(fileIn)) {
            StringBuilder text = new StringBuilder();
            for (var paragraph : document.getParagraphs()) {
                text.append(paragraph.getText());
            }
            return ToolResponse.success(
                new TextContent("Text from Word document at " + filepath + ": " + text));
        } catch (Exception e) {
            return ToolResponse.error("Failed to read text from Word document: " + e.getMessage());
        }
    }
}
