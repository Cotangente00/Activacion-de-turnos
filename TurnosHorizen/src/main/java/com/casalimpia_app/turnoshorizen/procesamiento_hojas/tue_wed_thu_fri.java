/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.turnoshorizen.procesamiento_hojas;

//import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author jcavilaa
 */
public class tue_wed_thu_fri {
    public static void escribirDatos(Sheet wsFilasCoincidentes, Sheet ws1, List<Double> activaTurno, List<Double> noActivaTurno, int columnaIndex, Date fechaObjetivo, int diaSemana){
        // Iterar sobre las filas de wsFilasCoincidentes comenzando desde la fila 2
        for (int i = 1; i <= wsFilasCoincidentes.getLastRowNum(); i++) {
            Row row = wsFilasCoincidentes.getRow(i);

            // Si la fila es nula, continuar con la siguiente
            if (row == null) {
                continue;
            }

            // Leer la celda de la columna A (número de documento)
            Cell cellA = row.getCell(0);
            if (cellA == null || cellA.getCellType() == Cell.CELL_TYPE_BLANK) {
                // Detener la iteración si la celda A está vacía
                break;
            }

            // Obtener el valor del número de documento (como tipo numérico)
            double numeroDocumento = cellA.getNumericCellValue();

            // Leer las celdas de las columnas C y D
            Cell cellC = row.getCell(2); // Columna C
            Cell cellD = row.getCell(3); // Columna D

            boolean isCellCEmpty = (cellC == null || cellC.getCellType() == Cell.CELL_TYPE_BLANK);
            boolean isCellDEmpty = (cellD == null || cellD.getCellType() == Cell.CELL_TYPE_BLANK);

            // Evaluar las condiciones
            if (isCellCEmpty && isCellDEmpty) {
                noActivaTurno.add(numeroDocumento);
            } else if (!isCellCEmpty && !isCellDEmpty) {
                activaTurno.add(numeroDocumento);
            } else if (!isCellCEmpty && isCellDEmpty) {
                activaTurno.add(numeroDocumento);
            } else if (isCellCEmpty && !isCellDEmpty) {
                System.out.println("Error");
            }
        }       // Imprimir las listas para verificar la información recopilada
        System.out.println("Números de documento en noActivaTurno: " + noActivaTurno);
        System.out.println("Números de documento en activaTurno: " + activaTurno);
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // Iterar sobre las filas de ws1 comenzando desde la fila 3
        for (int i = 2; i <= ws1.getLastRowNum(); i++) {
            Row row = ws1.getRow(i);

            // Si la fila es nula, continuar con la siguiente
            if (row == null) {
                continue;
            }

            // Leer la celda de la columna A (número de documento)
            Cell cellA = row.getCell(0);
            if (cellA == null || cellA.getCellType() == Cell.CELL_TYPE_BLANK) {
                // Saltarse la celda si se encuentra una vacía
                continue;
            }

            double numeroDocumento = cellA.getNumericCellValue();

            // Crear la celda en la columna correspondiente si no existe
            Cell cellResultado = row.getCell(columnaIndex);
            if (cellResultado == null) {
                cellResultado = row.createCell(columnaIndex);
            }



            // Comparar el número de documento con las listas y escribir en la columna correspondiente
            if (activaTurno.contains(numeroDocumento)) {
                cellResultado.setCellValue("Activa turno");
            } else if (noActivaTurno.contains(numeroDocumento)) {
                cellResultado.setCellValue("No activa turno");
            }
        }       //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // Iteración para agregar el valor de "Bloqueada" para las expertas que tienen fecha de retiro
        for (int i = 2; i <= ws1.getLastRowNum(); i ++){
            Row row = ws1.getRow(i);

            // Si la fila es nula, continuar con la siguiente

            if (row == null){
                continue;
            }

            Cell cellG = row.getCell(6);
            if (cellG == null || cellG.getCellType() == Cell.CELL_TYPE_BLANK || cellG.getCellType() == Cell.CELL_TYPE_STRING){
                // Saltarse la celda si se encuentra una vacía o una celda de tipo texto
                continue;
            }
            // Obtener las fechas
            Date fechasColumnaG = cellG.getDateCellValue();
            // Crear la celda en la columna correspondiente si no existe
            Cell cellResultado = row.getCell(columnaIndex);
            if (cellResultado == null) {
                cellResultado = row.createCell(columnaIndex);
            }

            if (fechaObjetivo.equals(fechasColumnaG)){
                cellResultado.setCellValue("Bloqueada");
            }
        }       
        
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Iteración para asignar el valor de "No carga turno" para los valores vacíos restantes
        for (int i = 2; i <= ws1.getLastRowNum(); i++){
            Row row = ws1.getRow(i);

            // Si la fila es nula, detener el ciclo
            if (row == null){
                break;
            }

            // Obtener la columna con la que se está trabajando
            Cell cellObjetivo = row.getCell(columnaIndex);

            // Obtener la columna de las cedula para detener el cilo
            Cell numeroDocumento = row.getCell(0);
            int validacion = 0;

            if (numeroDocumento == null || numeroDocumento.getCellType() == Cell.CELL_TYPE_BLANK){
                validacion++;
                continue;
            }

            // Crear la celda en la columna correspondiente si no existe
            Cell cellResultado = row.getCell(columnaIndex);
            if (cellResultado == null) {
                cellResultado = row.createCell(columnaIndex);
            }

            if (cellObjetivo == null || cellObjetivo.getCellType() == Cell.CELL_TYPE_BLANK && validacion == 0){
                cellResultado.setCellValue("No carga de turno");
            } else {
                continue;
            }
        }       //---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // Iteración para "ignorar" todas las celdas cuyos valores en la columna "Fecha de creación", (columna C) sea mayor a la fecha objetivo encontrada. seteando un valor null o vaciando todos los caracteres de la misma.
        for (int i = 2; i <= ws1.getLastRowNum(); i++){
            Row row = ws1.getRow(i);

            // Si la fila es nula, detener el ciclo
            if (row == null){
                break;
                }

                Cell cellC = row.getCell(2);
                if (cellC == null || cellC.getCellType() == Cell.CELL_TYPE_BLANK){
                    // Detener la iteración si se encuentra una celda vacía
                    continue;
                }

                // Obtener las fechas de creación de usuaios Horizen
                Date fechasCreacion = cellC.getDateCellValue();
                // Celda en donde se escriben los datos
                Cell cellResultado = row.getCell(columnaIndex);
                if (cellResultado == null) {
                    cellResultado = row.createCell(columnaIndex);
                }

                if (fechasCreacion.after(fechaObjetivo)){
                    cellResultado.setCellType(CellType.BLANK);
                } 

            }       System.out.println(fechaObjetivo);
            
        //---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // Iteración para ignorar todas las celdas (setearlas a nulas), cuyas horas son 200 (Columna F) y la fecha objetivo sea perteneciente a un Sábado o domingo

        for (int i = 2; i <= ws1.getLastRowNum(); i++){
            Row row = ws1.getRow(i);

            // Si se encuentra una fila vacía, se detiene la iteración
            if (row == null){
                break;
            }

            Cell cellF = row.getCell(5);
            if (cellF == null || cellF.getCellType() == Cell.CELL_TYPE_BLANK){
                // Detener la iteración si se encuentra una celda vacía
                break;
            }

            // Celda en donde se escriben los datos
            Cell cellResultado = row.getCell(columnaIndex);
            if (cellResultado == null) {
                cellResultado = row.createCell(columnaIndex);
            }

            double horas = cellF.getNumericCellValue();
            if (horas == 200 && (diaSemana == Calendar.SATURDAY || diaSemana == Calendar.SUNDAY)){
                cellResultado.setCellType(CellType.BLANK);
            } else {
                continue;
            }

        }       
        
        //---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        // Iteración para setear a null todas las celdas, cuyas fechas de retiro (columna G) sean menores a la fecha objetivo trabajada
        for (int i = 2; i <= ws1.getLastRowNum(); i++){
            Row row = ws1.getRow(i);

            // Si se encuentra una fila vacía, saltarse a la siguiente
            if (row == null){
                break;
            }

            //Celda en donde se escriben los datos
            Cell cellResultado = row.getCell(columnaIndex);
            //Si esta celda no existe, se crea una nueva
            if (cellResultado == null) {
                cellResultado = row.createCell(columnaIndex);
            }

            Cell cellG = row.getCell(6);
            if (cellG == null || cellG.getCellType() == Cell.CELL_TYPE_BLANK){
                // Si se encuentra con una celda vacía, continuar a la siguiente
                continue;
            } else if (cellG.getCellType() == Cell.CELL_TYPE_STRING){
                cellResultado.setCellType(CellType.BLANK);
                continue;
            }
            // Obtener las fechas de retiro de la columna F
            Date fechasRetiro = cellG.getDateCellValue();
            if (fechasRetiro.before(fechaObjetivo) || cellG.getCellType() == Cell.CELL_TYPE_STRING){
                cellResultado.setCellType(CellType.BLANK);
            }
        }
    }
}

