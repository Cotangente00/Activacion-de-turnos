/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.turnoshorizen.model;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException; 
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author jcavilaa
 */
public class FileProcessor {
    // Método para procesar los archivos Excel
    public static void processExcelFiles(File file1, File file2) throws IOException, InvalidFormatException {
        // Leer ambos archivos
        try (FileInputStream fis1 = new FileInputStream(file1);    
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook workbook1 = WorkbookFactory.create(fis1);
            Workbook workbook2 = WorkbookFactory.create(fis2)) {

            // Supongamos que solo trabajas con la primera hoja de ambos archivos
            Sheet sheet1 = workbook1.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(0);

            // Aquí es donde implementas las condiciones personalizadas para modificar archivo1
            // basado en el contenido de archivo2. Esto es solo un ejemplo sencillo:
            for (Row row1 : sheet1) {
                Cell cell1 = row1.getCell(0); // Supongamos que trabajamos con la primera columna
                if (cell1 != null) {
                    for (Row row2 : sheet2) {
                        Cell cell2 = row2.getCell(0); // Compara con la primera columna de archivo2
                        if (cell2 != null && cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                            // Si se cumple la condición, modifica archivo1
                            Cell modifyCell = row1.createCell(1); // Modificamos una celda en archivo1
                            modifyCell.setCellValue("Modificado"); // Modificación de ejemplo
                        }
                    }
                }
            }

            // Guardar los cambios en un nuevo archivo Excel
            try (FileOutputStream fos = new FileOutputStream("archivo_modificado.xlsx")) {
                workbook1.write(fos);
                System.out.println("Archivo modificado guardado como 'archivo_modificado.xlsx'.");
            }
        }
    }
}
