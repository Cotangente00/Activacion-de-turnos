/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.turnoshorizen.procesamiento_hojas;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
//import java.util.Map;
import java.util.Set;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author jcavilaa
 */
public class service {
    public static void coincidencias(Workbook wb1, Workbook wb2) throws Exception {
        Sheet ws1 = wb1.getSheet("TURNOS MES DE OCTUBRE");
        Sheet ws2 = wb2.getSheetAt(0);
        
        // Obtener las columnas de interés como iteradores
        Iterator<Row> rowIterator1 = ws1.iterator();
        rowIterator1.next();
        rowIterator1.next(); // Empezar directamente desde la fila 3        
        Iterator<Row> rowIterator2 = ws2.iterator();
        rowIterator2.next(); // Saltar el encabezado de la información de los turnos
        
        // Crear listas para almacenar los nombres de ambas hojas
        List<Double> numDocAsistencia = new ArrayList<>();
        List<Double> numDocTurnos = new ArrayList<>();
        List<Double> numerosCoincidencias = new ArrayList<>();
        List<Date> fechasPrincipales = new ArrayList<>();  
        List<Date> fechasIniciales = new ArrayList<>();
        List<Date> fechasFinales = new ArrayList<>(); 
        //List<Double> numeroDocumento = new ArrayList<>();
        
        //fecha de referencia para ejecutar pruebas
        LocalDate fechaReferenciaPrueba = LocalDate.of(2024, 10, 02);
        

        // Llenar las listas con los nombres
        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Double numeroAsistencia = row.getCell(0 ).getNumericCellValue(); 
            numDocAsistencia.add(numeroAsistencia);
        }
        System.out.println("Conjunto de nombres de asistencias:" + numDocAsistencia);
        
        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell cell = row.getCell(6);
            // Verificar si la celda existe y si su tipo es numérico
            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                Double numeroTurno = cell.getNumericCellValue();
                numDocTurnos.add(numeroTurno);
            }
        }
        System.out.println("Conjunto de nombres de turnos:" + numDocTurnos);
        
        // Recorrer lista de turnos y buscar coincidencias en la lista de asistencias
        rowIterator2 = ws2.iterator(); // Reiniciar iterador para capturar filas de nuevo
        rowIterator2.next(); // Saltar encabezado

        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell cell = row.getCell(6);
            // Verificar si la celda existe y si su tipo es numérico
            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                
                Double numeroTurno = cell.getNumericCellValue();
                if (numDocAsistencia.contains(numeroTurno)) {
                    numerosCoincidencias.add(numeroTurno);

                    // Obtener la fecha principal de la columna A
                    Cell fechaPrincipalCell = row.getCell(0);
                    if (fechaPrincipalCell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(fechaPrincipalCell) || fechaPrincipalCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        Date fecha = fechaPrincipalCell.getDateCellValue();
                        fechasPrincipales.add(fecha);
                    }
                    // Obtener la fecha inicial de la columna 
                    Cell fechaInicialCell = row.getCell(2);
                    if (fechaInicialCell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(fechaInicialCell) || fechaInicialCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        Date fecha = fechaInicialCell.getDateCellValue();
                        fechasIniciales.add(fecha);
                    }

                    // Obtener la fecha final de la columna
                    Cell fechaFinalCell = row.getCell(4);
                    if (fechaFinalCell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(fechaFinalCell) || fechaFinalCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        Date fecha = fechaFinalCell.getDateCellValue();
                        fechasFinales.add(fecha);
                    }
                }
            }
        }

        // Mostrar resultados en la terminal
        for (int i = 0; i < numerosCoincidencias.size(); i++) {
            System.out.println(numerosCoincidencias.get(i) + ": " + fechasPrincipales.get(i) + " // " + fechasIniciales.get(i) + " // " + fechasFinales.get(i));
        }
            
        /*
        // Crear un nuevo libro de trabajo y una nueva hoja
        XSSFWorkbook wb3 = new XSSFWorkbook();
        Sheet sheet = wb3.createSheet("Resultados");
        Row row;
        int rowNum = 0;

        // Escribir los encabezados
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Nombres coincidencias");
        row.createCell(1).setCellValue("Nombres fechas");

        // Escribir los datos de los conjuntos
        for (String nombreAsistencia : numDocAsistencia) {
            row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(nombreAsistencia);
        }

        for (String nombreTurno : nombresTurnos) {
            row = sheet.createRow(rowNum++);
            row.createCell(1).setCellValue(nombreTurno);
        }

        // Crear el archivo Excel
        FileOutputStream outputStream = new FileOutputStream("O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/resultados.xlsx");
        wb3.write(outputStream);
        outputStream.close();
        */
        
        
        // Crear una nueva hoja en wb2
        Sheet newSheet = wb2.createSheet("Resultados");

        // Crear los encabezados
        Row headerRow = newSheet.createRow(0);
        headerRow.createCell(0).setCellValue("Num. Doc.");
        headerRow.createCell(1).setCellValue("Fecha Principal");
        headerRow.createCell(2).setCellValue("Fecha Inicial");
        headerRow.createCell(3).setCellValue("Fecha Final");

        // Escribir los datos en el nuevo libro
        int rowNum = 1;
        for (int i = 0; i < numerosCoincidencias.size(); i++) {
            Row dataRow = newSheet.createRow(rowNum++);
            dataRow.createCell(0).setCellValue(numerosCoincidencias.get(i));
            dataRow.createCell(1).setCellValue(fechasPrincipales.get(i));
            dataRow.createCell(2).setCellValue(fechasIniciales.get(i));
            dataRow.createCell(3).setCellValue(fechasFinales.get(i));
        }
        
        // Nueva sección para encontrar la fecha objetivo
        Row row = ws1.getRow(1); // Suponiendo que la fila de interés es la 2 (índice 1)
        int colIndex = 7; // Columna H (índice 7)
        Cell cell;

        //LocalDate fechaReferenciaPrueba = LocalDate.of(2024, 10, 03);
        LocalDate fechaObjetivo = fechaReferenciaPrueba.minusDays(1); // Fecha objetivo (día anterior)
        String fechaObjetivoStr = fechaObjetivo.format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        Date fechaObjetivoEncontrada = null;        
        List<Object> coordenadas = new ArrayList<>();
        
        while (colIndex < row.getLastCellNum()) {
            cell = row.getCell(colIndex);
            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                Date fechaCelda = cell.getDateCellValue();
                LocalDate fechaCeldaLocalDate = fechaCelda.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                String fechaCeldaStr = fechaCeldaLocalDate.format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                if (fechaCeldaStr.equals(fechaObjetivoStr)) {
                    // Obtener el nombre de la columna utilizando CellReference
                    String columna = CellReference.convertNumToColString(colIndex);
                    coordenadas.add(columna); // Añadir la columna 
                    coordenadas.add(2); // la fila siempre es dos
                    System.out.println("Fecha objetivo encontrada en la celda " + columna + "2: " + fechaCelda);
                    fechaObjetivoEncontrada = fechaCelda;
                    // Acceder a las coordenadas:
                    System.out.println(coordenadas);
                    System.out.print(fechaObjetivoEncontrada + " " + ((Object)fechaObjetivoEncontrada).getClass().getSimpleName());
                    //System.out.println(fechaReferenciaPrueba);
                    break; // Detener la iteración
                }
            }
            colIndex++;
        }
        // Nueva sección: iterar sobre la columna B de la hoja 'Resultados'
        Sheet resultadosSheet = wb2.getSheet("Resultados");
        if (resultadosSheet == null) {
            throw new IllegalStateException("La hoja 'Resultados' no se encontró.");
        }

        // Crear una nueva hoja para almacenar las filas coincidentes
        Sheet newSheet2 = wb2.createSheet("FilasCoincidentes");
        Row newHeaderRow = newSheet2.createRow(0);
        newHeaderRow.createCell(0).setCellValue("Num. Doc.");
        newHeaderRow.createCell(1).setCellValue("Fecha Principal");
        newHeaderRow.createCell(2).setCellValue("Fecha Inicial");
        newHeaderRow.createCell(3).setCellValue("Fecha Final");

        // Comenzar la iteración en la fila 2 (índice 1)
        Iterator<Row> rowIterator = resultadosSheet.iterator();
        rowIterator.next(); // Saltar el encabezado

        int newRowNum = 1; // Para la nueva hoja
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Cell fechaCell = currentRow.getCell(1); // Columna B es índice 1

            if (fechaCell != null && fechaCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                Date fechaEnResultados = fechaCell.getDateCellValue();
                // Comparar la fecha con la fechaObjetivoEncontrada
                if (fechaEnResultados.equals(fechaObjetivoEncontrada)) {
                    // Copiar la fila completa a la nueva hoja
                    Row newRow = newSheet2.createRow(newRowNum++);
                    newRow.createCell(0).setCellValue(currentRow.getCell(0).getNumericCellValue()); // Columna A (Double)
                    newRow.createCell(1).setCellValue(fechaEnResultados); // Columna B (Fecha)
                    newRow.createCell(2).setCellValue(currentRow.getCell(2).getDateCellValue()); // Columna C (Fecha)
                    newRow.createCell(3).setCellValue(currentRow.getCell(3).getDateCellValue()); // Columna D (Fecha y hora)
                }
            }
        }
        
        // Guardar el nuevo libro
        FileOutputStream outputStream = new FileOutputStream("O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/Results.xlsx");
        wb2.write(outputStream);
        outputStream.close();
        wb2.close();
        //System.out.println("Archivo Excel creado exitosamente: resultados.xlsx");
        
    }
    
    public static void main(String[] args) throws Exception {
        String inputFilePath1 = "O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/Asistencia Octubre-2024-Turnos Horizen (2).xlsx";
        String inputFilePath2 = "O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/informe horizen 11 octubre.xlsx";
        /*
        String outputFilePath1 = "O:/aa/result2.xlsx";
        String outputFilePath2 = "O:/aa/result2.xlsx";
        */
        Workbook wb1, wb2;
        try (FileInputStream fis1 = new FileInputStream(new File(inputFilePath1))) {
            wb1 = WorkbookFactory.create(fis1);  // Apache POI detecta automáticamente si es .xls o .xlsx
        }
        
        try (FileInputStream fis2 = new FileInputStream(new File(inputFilePath2))) {
            wb2 = WorkbookFactory.create(fis2);
        }

        try {
            coincidencias(wb1, wb2);
            /*
            wb1.write(new FileOutputStream(outputFilePath1));
            wb1.close();
            wb2.write(new FileOutputStream(outputFilePath2));
            wb2.close();
            */
            System.out.println("Archivo procesado exitosamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
