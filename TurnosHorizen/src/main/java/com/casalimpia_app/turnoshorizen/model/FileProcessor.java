/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.turnoshorizen.model;
import static com.casalimpia_app.turnoshorizen.procesamiento_hojas.writeData.validacionTurnos;

import static com.casalimpia_app.turnoshorizen.procesamiento_hojas.service.coincidencias;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException; 
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author jcavilaa
 */
public class FileProcessor {
    // Método para procesar los archivos Excel
    public static void processExcelFiles(File file1, File file2) throws IOException, InvalidFormatException, Exception {
        // Leer ambos archivos
        try (FileInputStream fis1 = new FileInputStream(file1);    
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook wb1 = WorkbookFactory.create(fis1);
            Workbook wb2 = WorkbookFactory.create(fis2)) {
            
            coincidencias(wb1, wb2);
            validacionTurnos(wb1, wb2);
            
            // Guardar los cambios en un nuevo archivo Excel
            try (FileOutputStream fos = new FileOutputStream("O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/archivo modificado.xlsx")) {
                wb1.write(fos);
                fos.close();
                wb1.close();
                wb2.close();
                System.out.println("Archivo modificado guardado como 'archivo_modificado.xlsx'.");
            }
        }
    }
}
