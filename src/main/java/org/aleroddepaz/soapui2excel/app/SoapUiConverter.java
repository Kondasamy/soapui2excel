package org.aleroddepaz.soapui2excel.app;

import java.io.File;
import java.util.List;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import javax.xml.transform.stream.StreamSource;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import org.aleroddepaz.soapuimodel.ObjectFactory;
import org.aleroddepaz.soapuimodel.Project;
import org.aleroddepaz.soapuimodel.TestCase;
import org.aleroddepaz.soapuimodel.TestSuite;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class SoapUiConverter {

    private static final Logger LOGGER = LoggerFactory
            .getLogger(SoapUiConverter.class);

    private static Unmarshaller unmarshaller;

    static {
        try {
            unmarshaller = JAXBContext.newInstance(ObjectFactory.class)
                    .createUnmarshaller();
        } catch (JAXBException e) {
            LOGGER.error("Error during JAXB Unmarshaller creation: " + e.getMessage());
        }
    }

    public void convertToExcel(File inputFile, File outputFile)
            throws Exception {
        LOGGER.debug("Converting {} to {}...", inputFile.getAbsolutePath(),
                outputFile.getAbsolutePath());

        WritableWorkbook workbook = Workbook.createWorkbook(outputFile);
        JAXBElement<Project> element = unmarshaller.unmarshal(new StreamSource(
                inputFile), Project.class);
        Project project = element.getValue();
        int sheetNum = 0;
        for (TestSuite testSuite : project.getTestSuite()) {
            WritableSheet sheet = workbook.createSheet(testSuite.getName(),
                    sheetNum++);
            List<TestCase> testCases = testSuite.getTestCase();
            int row = 0;
            for (TestCase testCase : testCases) {
                writeLabel(sheet, testCase.getName(), row, 0);
                if (testCase.getDescription() != null) {
                    writeLabel(sheet, testCase.getDescription().trim(), row, 1);
                }
                row++;
            }
        }
        workbook.write();
        workbook.close();
    }

    private static void writeLabel(WritableSheet sheet, String string, int row,
            int column) throws WriteException, RowsExceededException {
        Label label = new Label(column, row, string);
        sheet.addCell(label);
    }

}
