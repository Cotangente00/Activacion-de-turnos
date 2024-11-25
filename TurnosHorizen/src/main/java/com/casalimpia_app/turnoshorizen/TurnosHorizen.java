/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.casalimpia_app.turnoshorizen;

import com.casalimpia_app.turnoshorizen.model.FileProcessor;
import static com.casalimpia_app.turnoshorizen.model.FileProcessor.processExcelFiles;
import static com.casalimpia_app.turnoshorizen.procesamiento_hojas.writeData.validacionTurnos;
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
import javafx.scene.control.Label;
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

        // Etiquetas para mostrar los archivos seleccionados
        Label file1Label = new Label("Archivo 1 no seleccionado.");
        Label file2Label = new Label("Archivo 2 no seleccionado.");

        // Botón 1 para seleccionar archivo 1
        Button selectFile1Button = new Button("Seleccione Asistencia");
        selectFile1Button.setOnAction(e -> {
            file1 = fileChooser.showOpenDialog(primaryStage);
            if (file1 != null) {
                file1Label.setText("Archivo seleccionado: " + file1.getName());
            } else {
                file1Label.setText("Archivo 1 no seleccionado.");
            }
        });

        // Botón 2 para seleccionar archivo 2
        Button selectFile2Button = new Button("Seleccione Información de turnos");
        selectFile2Button.setOnAction(e -> {
            file2 = fileChooser.showOpenDialog(primaryStage);
            if (file2 != null) {
                file2Label.setText("Archivo seleccionado: " + file2.getName());
            } else {
                file2Label.setText("Archivo 2 no seleccionado.");
            }
        });

        // Botón para iniciar el procesamiento
        Button processFilesButton = new Button("Procesar Archivos");
        processFilesButton.setOnAction(e -> {
            if (file1 != null && file2 != null) {
                try {
                    FileProcessor.processExcelFiles(file1, file2, primaryStage); // Procesar los archivos
                } catch (IOException ex) {
                    ex.printStackTrace();
                } catch (InvalidFormatException ex) {
                    Logger.getLogger(TurnosHorizen.class.getName()).log(Level.SEVERE, null, ex);
                } catch (Exception ex) {
                    Logger.getLogger(TurnosHorizen.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else {
                System.out.println("Seleccione ambos archivos antes de continuar.");
            }
        });

        // Botón para eliminar los archivos seleccionados
        Button clearFilesButton = new Button("Eliminar Archivos Seleccionados");
        clearFilesButton.setOnAction(e -> {
            file1 = null;
            file2 = null;
            file1Label.setText("Archivo 1 no seleccionado.");
            file2Label.setText("Archivo 2 no seleccionado.");
            System.out.println("Archivos seleccionados eliminados.");
        });

        // Layout de la interfaz gráfica
        VBox layout = new VBox(10, selectFile1Button, file1Label, selectFile2Button, file2Label, processFilesButton, clearFilesButton);
        layout.setAlignment(Pos.CENTER);
        Scene scene = new Scene(layout, 450, 250);

        primaryStage.setTitle("Procesamiento de Archivos Excel");
        primaryStage.setResizable(false);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    // Método ficticio para procesar los archivos (implementa tu lógica aquí)
    private void processExcelFiles(File file1, File file2) throws IOException, InvalidFormatException {
        // Aquí va tu lógica de procesamiento
        System.out.println("Procesando archivos: " + file1.getName() + " y " + file2.getName());
    }
}