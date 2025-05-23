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
public class PowerPointFeaturesTest {

    private File tempDir;
    private OfficeMcpServerPowerPointFeatures powerPointFeatures;
    private String presentationPath;

    @BeforeEach
    void setUp() throws Exception {
        // Create a temporary directory manually
        String tempDirPath = System.getProperty("java.io.tmpdir");
        tempDir = new File(tempDirPath, "ppt-test-" + UUID.randomUUID());
        tempDir.mkdirs();
        
        powerPointFeatures = new OfficeMcpServerPowerPointFeatures();
        presentationPath = tempDir.getAbsolutePath() + "/test.pptx";
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
    void testCreatePowerPointPresentation() throws Exception {
        powerPointFeatures.createPowerPointPresentation(presentationPath);
        assertTrue(Files.exists(Paths.get(presentationPath)));
    }

    @Test
    void testAddSlideToPowerPoint() throws Exception {
        // First create the presentation
        powerPointFeatures.createPowerPointPresentation(presentationPath);
        
        // Add a slide
        powerPointFeatures.addSlideToPowerPoint(presentationPath);
        
        // Verify the slide was added by checking the file exists
        assertTrue(Files.exists(Paths.get(presentationPath)));
    }
    
    @Test
    void testAddTextToPowerPointSlide() throws Exception {
        // Create presentation and add a slide
        powerPointFeatures.createPowerPointPresentation(presentationPath);
        powerPointFeatures.addSlideToPowerPoint(presentationPath);
        
        // Add text to the slide
        String testText = "Test Presentation Text";
        assertNotNull(powerPointFeatures.addTextToPowerPointSlide(presentationPath, 0, testText));
        
        // Verify the file still exists
        assertTrue(Files.exists(Paths.get(presentationPath)));
    }
    
    @Test
    void testReadSlideTitlesFromPowerPoint() throws Exception {
        // Create presentation with slides
        powerPointFeatures.createPowerPointPresentation(presentationPath);
        powerPointFeatures.addSlideToPowerPoint(presentationPath);
        powerPointFeatures.addSlideToPowerPoint(presentationPath);
        
        // Read slide titles
        assertNotNull(powerPointFeatures.readSlideTitlesFromPowerPoint(presentationPath));
    }
    
    @Test
    void testGetPowerPointSlideCount() throws Exception {
        // Create presentation with two slides
        powerPointFeatures.createPowerPointPresentation(presentationPath);
        powerPointFeatures.addSlideToPowerPoint(presentationPath);
        powerPointFeatures.addSlideToPowerPoint(presentationPath);
        
        // Get slide count
        assertNotNull(powerPointFeatures.getPowerPointSlideCount(presentationPath));
    }
}