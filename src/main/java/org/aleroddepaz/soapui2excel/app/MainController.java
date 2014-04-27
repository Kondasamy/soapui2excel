package org.aleroddepaz.soapui2excel.app;

import java.io.File;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;

public class MainController {

    private final SoapUiConverter converter = new SoapUiConverter();
    private final ExtensionFilter inputFilter = new ExtensionFilter(
            "SoapUI projects (*.xml)", "*.xml");
    private final ExtensionFilter outputFilter = new ExtensionFilter(
            "Microsoft Excel files (*.xls)", "*.xls");
    private final FileChooser inputFileChooser = new FileChooser();
    private final FileChooser outputFileChooser = new FileChooser();
    private File inputFile;

    protected Stage getScene() {
        return (Stage) targetAction.getScene().getWindow();
    }

    @FXML
    protected Label targetAction;

    @FXML
    protected TextField inputPath;

    @FXML
    protected void initialize() {
        inputFileChooser.getExtensionFilters().add(inputFilter);
        outputFileChooser.getExtensionFilters().add(outputFilter);
    }

    @FXML
    protected void handleOpenAction(ActionEvent event) {
        inputFile = inputFileChooser.showOpenDialog(getScene());
        if (inputFile != null) {
            inputPath.setText(inputFile.getAbsolutePath());
        }
    }

    @FXML
    protected void handleConvertAction(ActionEvent event) {
        if (inputFile != null) {
            String filename = inputFile.getName();
            if (filename.indexOf(".") > 0) {
                filename = filename.substring(0, filename.lastIndexOf("."));
            }
            outputFileChooser.setInitialFileName(filename + ".xls");
            File outputFile = outputFileChooser.showSaveDialog(getScene());
            if (outputFile != null) {
                try {
                    converter.convertToExcel(inputFile, outputFile);
                    targetAction.setStyle("-fx-text-fill: green;");
                    targetAction.setText("Proyecto convertido con éxito");
                } catch (Exception e) {
                    targetAction.setStyle("-fx-text-fill: firebrick;");
                    targetAction.setText("Error al convertir el proyecto: "
                            + e.getMessage());
                }
            }
        } else {
            targetAction.setStyle("-fx-text-fill: firebrick;");
            targetAction.setText("Selecciona un proyecto SoapUI para convertir");
        }
    }

    @FXML
    protected void handleMenuExit(ActionEvent event) {
        getScene().close();
    }
}
