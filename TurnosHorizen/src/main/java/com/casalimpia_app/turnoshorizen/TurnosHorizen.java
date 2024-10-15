/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.casalimpia_app.turnoshorizen;

import static com.casalimpia_app.turnoshorizen.model.FileProcessor.processExcelFiles;
import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.VBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.geometry.Pos;
import javafx.scene.layout.HBox;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author jcavilaa
 */
public class TurnosHorizen extends Application {

    private File file1; // Primer archivo Excel
    private File file2; // Segundo archivo Excel

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        FileChooser fileChooser = new FileChooser();

        // Botón 1 para seleccionar archivo 1
        Button selectFile1Button = new Button("Seleccione Asistencia");
        selectFile1Button.setOnAction(e -> {
            file1 = fileChooser.showOpenDialog(primaryStage);
            if (file1 != null) {
                System.out.println("Archivo seleccionado: " + file1.getName());
            }
        });

        // Botón 2 para seleccionar archivo 2
        Button selectFile2Button = new Button("Seleccione Información de turnos");
        selectFile2Button.setOnAction(e -> {
            file2 = fileChooser.showOpenDialog(primaryStage);
            if (file2 != null) {
                System.out.println("Archivo seleccionado: " + file2.getName());
            }
        });

        // Botón para iniciar el procesamiento
        Button processFilesButton = new Button("Procesar Archivos");
        processFilesButton.setOnAction(e -> {
            if (file1 != null && file2 != null) {
                try {
                    processExcelFiles(file1, file2); // Procesar los archivos
                } catch (IOException ex) {
                    ex.printStackTrace();
                } catch (InvalidFormatException ex) {
                    Logger.getLogger(TurnosHorizen.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else {
                System.out.println("Seleccione ambos archivos antes de continuar.");
            }
        });

        // Layout de la interfaz gráfica
        VBox layout = new VBox(10, selectFile1Button, selectFile2Button, processFilesButton);
        layout.setAlignment(Pos.CENTER);
        Scene scene = new Scene(layout, 410, 110);

        primaryStage.setTitle("Procesamiento de Archivos Excel");
        primaryStage.setResizable(false);
        primaryStage.setScene(scene);
        primaryStage.show();
    }
}
