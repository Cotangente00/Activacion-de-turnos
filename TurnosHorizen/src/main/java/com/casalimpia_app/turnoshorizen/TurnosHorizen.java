/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.casalimpia_app.turnoshorizen;

import com.casalimpia_app.turnoshorizen.model.FileProcessor;
import static com.casalimpia_app.turnoshorizen.model.FileProcessor.processExcelFiles;
import static com.casalimpia_app.turnoshorizen.procesamiento_hojas.writeData.validacionTurnos;
import com.toedter.calendar.JDateChooser;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

public class TurnosHorizen extends Application {

    private File file1; // Primer archivo Excel
    private File file2; // Segundo archivo Excel
    private LocalDate fechaReferencia; // Variable para la fecha seleccionada

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        FileChooser fileChooser = new FileChooser();

        // Etiquetas para mostrar los archivos seleccionados
        Label file1Label = new Label("Archivo 1 no seleccionado.");
        Label file2Label = new Label("Archivo 2 no seleccionado.");
        Label fechaLabel = new Label("Fecha no seleccionada.");

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

        // Botón para seleccionar la fecha
        Button selectDateButton = new Button("Seleccione Fecha de Referencia");
        selectDateButton.setOnAction(e -> {
            JDateChooser dateChooser = new JDateChooser();
            dateChooser.setDateFormatString("yyyy-MM-dd");

            // Mostrar un cuadro de diálogo para seleccionar la fecha
            JOptionPane.showMessageDialog(null, dateChooser, "Seleccione una fecha", JOptionPane.PLAIN_MESSAGE);

            // Obtener la fecha seleccionada
            Date selectedDate = dateChooser.getDate();
            if (selectedDate != null) {
                fechaReferencia = selectedDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                fechaLabel.setText("Fecha seleccionada: " + fechaReferencia.toString());
            } else {
                fechaLabel.setText("Fecha no seleccionada.");
            }
        });

        // Botón para iniciar el procesamiento
        Button processFilesButton = new Button("Procesar Archivos");
        processFilesButton.setOnAction(e -> {
            if (file1 != null && file2 != null && fechaReferencia != null) {
                try {
                    FileProcessor.processExcelFiles(file1, file2, primaryStage, fechaReferencia); // Procesar los archivos
                } catch (IOException ex) {
                    ex.printStackTrace();
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            } else {
                System.out.println("Seleccione ambos archivos y una fecha antes de continuar.");
            }
        });

        // Botón para eliminar los archivos seleccionados
        Button clearFilesButton = new Button("Eliminar Archivos Seleccionados");
        clearFilesButton.setOnAction(e -> {
            file1 = null;
            file2 = null;
            fechaReferencia = null;
            file1Label.setText("Archivo 1 no seleccionado.");
            file2Label.setText("Archivo 2 no seleccionado.");
            fechaLabel.setText("Fecha no seleccionada.");
            System.out.println("Archivos y fecha seleccionados eliminados.");
        });

        // Layout de la interfaz gráfica
        VBox layout = new VBox(10, selectFile1Button, file1Label, selectFile2Button, file2Label, selectDateButton, fechaLabel, processFilesButton, clearFilesButton);
        layout.setAlignment(javafx.geometry.Pos.CENTER);
        Scene scene = new Scene(layout, 500, 300);

        primaryStage.setTitle("Procesamiento de Archivos Excel");
        primaryStage.setResizable(false);
        primaryStage.setScene(scene);
        primaryStage.show();
    }
}