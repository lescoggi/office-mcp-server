package dev.lescoggi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.eclipse.microprofile.config.inject.ConfigProperty;

import io.quarkiverse.mcp.server.TextContent;
import io.quarkiverse.mcp.server.Tool;
import io.quarkiverse.mcp.server.ToolArg;
import io.quarkiverse.mcp.server.ToolResponse;

public class OfficeMcpServerPowerPointFeatures {

    @ConfigProperty(name = "office.files.path")
    String officeFilesPath;

    @Tool(description = "Create a new PowerPoint presentation", name = "create_powerpoint_presentation")
    ToolResponse createPowerPointPresentation(@ToolArg(description = "Path to create new PowerPoint presentation") String filepath) {
        try (XMLSlideShow presentation = new XMLSlideShow(); FileOutputStream fileOut = new FileOutputStream(filepath)) {
            presentation.write(fileOut);
            return ToolResponse.success(
                new TextContent("PowerPoint presentation created at: " + filepath));
        } catch (Exception e) {
            return ToolResponse.error("Failed to create PowerPoint presentation: " + e.getMessage());
        }
    }

    @Tool(description = "Add a slide to a PowerPoint presentation", name = "add_slide_to_powerpoint")
    ToolResponse addSlideToPowerPoint(@ToolArg(description = "Path to the PowerPoint presentation") String filepath) {
        try (FileInputStream fileIn = new FileInputStream(filepath); XMLSlideShow presentation = new XMLSlideShow(fileIn)) {
            XSLFSlide slide = presentation.createSlide();
            
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                presentation.write(fileOut);
            }
            
            return ToolResponse.success(
                new TextContent("Slide added to PowerPoint presentation at: " + filepath));
        } catch (IOException e) {
            return ToolResponse.error("Failed to add slide to PowerPoint presentation: " + e.getMessage());
        }
    }

    @Tool(description = "Add text to a PowerPoint slide", name = "add_text_to_powerpoint_slide")
    ToolResponse addTextToPowerPointSlide(
            @ToolArg(description = "Path to the PowerPoint presentation") String filepath,
            @ToolArg(description = "Slide index (0-based)") int slideIndex, 
            @ToolArg(description = "Text to add") String text) {
        try (FileInputStream fileIn = new FileInputStream(filepath); XMLSlideShow presentation = new XMLSlideShow(fileIn)) {
            if (slideIndex >= presentation.getSlides().size()) {
                return ToolResponse.error("Slide index " + slideIndex + " is out of bounds. The presentation has " + 
                                         presentation.getSlides().size() + " slides.");
            }
            
            XSLFSlide slide = presentation.getSlides().get(slideIndex);
            XSLFTextShape textShape = slide.createTextBox();
            textShape.setText(text);
            
            // Set default position for text box
            textShape.setAnchor(new java.awt.Rectangle(50, 50, 400, 200));
            
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                presentation.write(fileOut);
            }
            
            return ToolResponse.success(
                new TextContent("Text added to slide " + slideIndex + " in PowerPoint presentation at: " + filepath));
        } catch (IOException e) {
            return ToolResponse.error("Failed to add text to PowerPoint slide: " + e.getMessage());
        }
    }

    @Tool(description = "Read slide titles from a PowerPoint presentation", name = "read_slide_titles_from_powerpoint")
    ToolResponse readSlideTitlesFromPowerPoint(@ToolArg(description = "Path to the PowerPoint presentation") String filepath) {
        try (FileInputStream fileIn = new FileInputStream(filepath); XMLSlideShow presentation = new XMLSlideShow(fileIn)) {
            int slideCount = presentation.getSlides().size();
            if (slideCount == 0) {
                return ToolResponse.success(
                    new TextContent("PowerPoint presentation at " + filepath + " has no slides."));
            }
            
            StringBuilder titles = new StringBuilder();
            titles.append("Presentation has " + slideCount + " slides.\n");
            
            int slideIndex = 0;
            for (XSLFSlide slide : presentation.getSlides()) {
                titles.append("Slide ").append(slideIndex++).append(": ");
                
                // Try to find a title in the slide - simplified approach 
                String title = "No title";
                List<XSLFShape> shapes = slide.getShapes();
                for (XSLFShape shape : shapes) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape) shape;
                        String text = textShape.getText();
                        if (text != null && !text.isEmpty()) {
                            title = text;
                            break;
                        }
                    }
                }
                
                titles.append(title).append("\n");
            }
            
            return ToolResponse.success(
                new TextContent(titles.toString()));
        } catch (IOException e) {
            return ToolResponse.error("Failed to read slide titles from PowerPoint presentation: " + e.getMessage());
        }
    }
    
    @Tool(description = "Get slide count from a PowerPoint presentation", name = "get_powerpoint_slide_count")
    ToolResponse getPowerPointSlideCount(@ToolArg(description = "Path to the PowerPoint presentation") String filepath) {
        try (FileInputStream fileIn = new FileInputStream(filepath); XMLSlideShow presentation = new XMLSlideShow(fileIn)) {
            int slideCount = presentation.getSlides().size();
            return ToolResponse.success(
                new TextContent("PowerPoint presentation at " + filepath + " has " + slideCount + " slides."));
        } catch (IOException e) {
            return ToolResponse.error("Failed to get slide count from PowerPoint presentation: " + e.getMessage());
        }
    }
}