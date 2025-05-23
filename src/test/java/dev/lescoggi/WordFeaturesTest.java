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
public class WordFeaturesTest {

    private File tempDir;
    private OfficeMcpServerWordFeatures wordFeatures;
    private String documentPath;

    @BeforeEach
    void setUp() throws Exception {
        // Create a temporary directory manually
        String tempDirPath = System.getProperty("java.io.tmpdir");
        tempDir = new File(tempDirPath, "word-test-" + UUID.randomUUID());
        tempDir.mkdirs();
        
        wordFeatures = new OfficeMcpServerWordFeatures();
        documentPath = tempDir.getAbsolutePath() + "/test.docx";
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
    void testCreateWordDocument() throws Exception {
        wordFeatures.createWordDocument(documentPath);
        assertTrue(Files.exists(Paths.get(documentPath)));
    }

    @Test
    void testAddTextToWordDocument() throws Exception {
        // First create the document
        wordFeatures.createWordDocument(documentPath);
        
        // Add text to the document
        String testText = "Test Document Text";
        assertNotNull(wordFeatures.addTextToWordDocument(documentPath, testText));
        
        // Verify the file exists
        assertTrue(Files.exists(Paths.get(documentPath)));
    }
    
    @Test
    void testReadTextFromWordDocument() throws Exception {
        // Create document and add text
        wordFeatures.createWordDocument(documentPath);
        String testText = "Test Document Text For Reading";
        wordFeatures.addTextToWordDocument(documentPath, testText);
        
        // Read text from the document
        assertNotNull(wordFeatures.readTextFromWordDocument(documentPath));
    }
}