package org.artofsolving.jodconverter;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.document.DocumentFormatRegistry;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.artofsolving.jodconverter.streams.OOoInputStream;
import org.artofsolving.jodconverter.streams.OOoOutputStream;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Set;

import static org.testng.Assert.assertNotNull;
import static org.testng.Assert.assertTrue;

@Test(groups = "Functional")
public class OfficeDocumentConverterStreamTest {

    public static final String FILTER_NAME_KEY = "FilterName";

    public void runAllPossibleConversions() throws IOException {
        OfficeManager officeManager = new DefaultOfficeManagerConfiguration().buildOfficeManager();
        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
        DocumentFormatRegistry formatRegistry = converter.getFormatRegistry();
        officeManager.start();

        OOoInputStream inputStream = null;
        OOoOutputStream oOoOutputStream = null;
        FileOutputStream fileOutputStream = null;
        try {
            File dir = new File("src/test/resources/documents");
            File[] files = dir.listFiles(new FilenameFilter() {
                public boolean accept(File dir, String name) {
                    return !name.startsWith(".");
                }
            });
            for (File inputFile : files) {
                // Must support this input file format
                String inputExtension = FilenameUtils.getExtension(inputFile.getName());
                DocumentFormat inputFormat = formatRegistry.getFormatByExtension(inputExtension);
                assertNotNull(inputFormat, "unknown input format: " + inputExtension);
                inputStream = new OOoInputStream(FileUtils.readFileToByteArray(inputFile));

                // Get set of output formats
                Set<DocumentFormat> outputFormats = formatRegistry.getOutputFormats(inputFormat.getInputFamily());

                for (DocumentFormat outputFormat : outputFormats) {
                    File outputFile = File.createTempFile("test", "." + outputFormat.getExtension());
                    outputFile.deleteOnExit();
                    fileOutputStream = new FileOutputStream(outputFile);


                    System.out.printf("-- converting %s to %s... ", inputFormat.getExtension(), outputFormat.getExtension());
                    String filterName = (String) outputFormat.getStoreProperties(inputFormat.getInputFamily()).get(FILTER_NAME_KEY);
                    oOoOutputStream = new OOoOutputStream();
                    converter.convert(inputStream,  oOoOutputStream, filterName);

                    // Write from OOoOutputStream to FileOutputStream
                    oOoOutputStream.writeTo(fileOutputStream);
                    System.out.printf("done.\n");

                    assertTrue(outputFile.isFile() && outputFile.length() > 0);

                    inputStream.close();
                    oOoOutputStream.close();
                    fileOutputStream.close();;
                    //TODO use file detection to make sure outputFile is in the expected format
                }
            }
        } finally {
            officeManager.stop();
            if (inputStream != null) {
                inputStream.close();
            }
            if (oOoOutputStream != null) {
                oOoOutputStream.close();
            }
            if (fileOutputStream != null) {
                fileOutputStream.close();
            }
        }
    }
}
