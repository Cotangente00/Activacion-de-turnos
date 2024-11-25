/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.turnoshorizen.procesamiento_hojas;

import static com.casalimpia_app.turnoshorizen.model.consultasSIC.supernumerarios;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
//import java.util.Map;
import static jdk.nashorn.internal.objects.NativeString.length;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author jcavilaa
 */
public class service {
    public static void coincidencias(Workbook wb1, Workbook wb2) throws Exception {
        Sheet ws1 = wb1.getSheet("TURNOS MES DE NOVIEMBRE");
        Sheet ws2 = wb2.getSheetAt(0);
        System.out.println("Hoja asistencia: " + ws1);
        System.out.println("Hoja turnos: " + ws2);
        // Obtener las columnas de interés como iteradores
        Iterator<Row> rowIterator1 = ws1.iterator();
        rowIterator1.next();
        rowIterator1.next(); // Empezar directamente desde la fila 3        
        Iterator<Row> rowIterator2 = ws2.iterator();
        rowIterator2.next(); // Saltar el encabezado de la información de los turnos
        
        // Crear listas para almacenar los números de documentos de ambas hojas
        List<Double> numDocAsistencia = new ArrayList<>();
        List<Double> numDocTurnos = new ArrayList<>();
        
        // Listas de las filas coincidentes por primera vez
        List<Double> numerosCoincidencias = new ArrayList<>();
        List<Date> fechasPrincipales = new ArrayList<>();  
        List<Date> fechasIniciales = new ArrayList<>();
        List<Date> fechasFinales = new ArrayList<>();
        
        
        // Listas de las filas coincidentes por segunda vez
        List<Double> numerosCoincidencias2 = new ArrayList<>();
        List<Date> fechasPrincipales2 = new ArrayList<>();  
        List<Date> fechasIniciales2 = new ArrayList<>();
        List<Date> fechasFinales2 = new ArrayList<>();
        
        // Listas de los casos ND
        List<Double> numDocNDTurnos = new ArrayList<>();
        List<String> nombreNDTurnos = new ArrayList<>();
        List<String> estadosNDTurnos = new ArrayList<>();
        
        List<Double> numDocNDAsistencia = new ArrayList<>();
        //List<Double> numeroDocumento = new ArrayList<>();
        
        //fecha de referencia para ejecutar pruebas
        LocalDate fechaReferenciaPrueba = LocalDate.of(2024, 11, 16);
        

        // Listas para almacenar los números de documento de los supernumerarios
        List<String> numDocSupernumerarios = new ArrayList<>();
        List<String> horario = new ArrayList<>();
        
        
        // Llenar las listas con los números de documento 
        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Double numeroAsistencia = row.getCell(0).getNumericCellValue(); 
            numDocAsistencia.add(numeroAsistencia);
        }
        //System.out.println("Conjunto de números de documento de la asistencia:" + numDocAsistencia);
        
        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell cell = row.getCell(6);
            // Verificar si la celda existe y si su tipo es numérico
            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                Double numeroTurno = cell.getNumericCellValue();
                numDocTurnos.add(numeroTurno);
            }
        }
        //System.out.println("Conjunto de números de documentos de los turnos:" + numDocTurnos);
        
        //Comprobar y añadir las ND para el archivo de asistencia
        for (Double numero : numDocAsistencia){
            if (numDocTurnos.contains(numero)){
                
            } else {
                numDocNDAsistencia.add(numero);
            }
        }
        System.out.println( "ND asistencia: " + numDocNDAsistencia);
        
        
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        // Recorrer lista de turnos y buscar coincidencias en la lista de asistencias
        rowIterator2 = ws2.iterator(); // Reiniciar iterador para capturar filas de nuevo
        rowIterator2.next(); // Saltar encabezado
        
        //Ejecutar la iteración por primera vez para agregar todas las expertas nuevas 
        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next(); 
            Cell numero = row.getCell(6);
            Cell nombre = row.getCell(7);
            Cell estado = row.getCell(11);
            // Verificar si la celda existe y si su tipo es numérico
            if (numero != null && numero.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                String nombreTurno = nombre.getStringCellValue();
                Double numeroTurno = numero.getNumericCellValue();
                String estadoTurno = estado.getStringCellValue();
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
                } else {
                    //Añadir las ND para el archivo de los turnos (nuevas)
                    numDocNDTurnos.add(numeroTurno);
                    nombreNDTurnos.add(nombreTurno);
                    estadosNDTurnos.add(estadoTurno);
                    //Actualizar la lista original con los números de documentos de las nuevas
                    numDocAsistencia.add(numeroTurno);
                }
            }
        }
        
        for (Double numero : numDocAsistencia){
            System.out.println( "Lista de números de documento de la asistencia actualizada: " + numero);
        }
        /*
        for (Double numero : numDocNDTurnos){
            System.out.println( "ND Números de documento de los turnos(nuevas): " + numero);
        }
        for (String nombre : nombreNDTurnos){
            System.out.println( "ND Nombres de los turnos(nuevas): " + nombre);
        }
        */
        
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        // Crear un estilo para la fuente en negrita
        CellStyle style = wb1.createCellStyle();
        Font font = wb1.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short)9);
        style.setFont(font);
        
        
        // Crear un estilo para la alineación centrada
        CellStyle centeredStyle = wb1.createCellStyle();
        centeredStyle.setAlignment(HorizontalAlignment.CENTER);
        Font font2 = wb1.createFont();
        font2.setFontHeightInPoints((short)9);
        font2.setBold(true);
        centeredStyle.setFont(font2);
        //Sección para escribir los datos de las nuevas en ws1
        Sheet SheetTurnos = wb1.createSheet("ND TURNOS");
        //int indice = 0;
        for (int i = 1; i <= ws1.getLastRowNum(); i++){
            Row row = SheetTurnos.getRow(i);
            //Cell cellA = row.getCell(0);
            if (row == null){
                
                for (int a = 0; a < numDocNDTurnos.size(); a ++){
                    Row row2 = SheetTurnos.createRow(a);
                    Cell cell = row2.createCell(0);
                    cell.setCellValue(numDocNDTurnos.get(a));
                    cell.setCellStyle(style);
                    Cell cell2 = row2.createCell(1);
                    cell2.setCellValue(nombreNDTurnos.get(a));
                    cell2.setCellStyle(style);

                    // Aplicar estilo centrado a la tercera columna
                    Cell cell3 = row2.createCell(3);
                    cell3.setCellValue(estadosNDTurnos.get(a));
                    cell3.setCellStyle(centeredStyle);
                }
            }
        }
        
        //System.out.println("indice en donde se ubicarán las expertas (más uno): " + indice);
        
       
        
        
        
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        //Ejecutar la iteración por segunda vez para agregar todas las expertas nuevas 
        // Recorrer lista de turnos y buscar coincidencias en la lista de asistencias
        rowIterator2 = ws2.iterator(); // Reiniciar iterador para capturar filas de nuevo
        rowIterator2.next(); // Saltar encabezado
        
        //Ejecutar la iteración por segunda ves para filtrar todas las expertas de interés
        
        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next(); 
            Cell numero = row.getCell(6);
            Cell nombre = row.getCell(7);
            // Verificar si la celda existe y si su tipo es numérico
            if (numero != null && numero.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                String nombreTurno = nombre.getStringCellValue();
                Double numeroTurno = numero.getNumericCellValue();
                
                if (numDocAsistencia.contains(numeroTurno)) {
                    numerosCoincidencias2.add(numeroTurno);

                    // Obtener la fecha principal de la columna A
                    Cell fechaPrincipalCell = row.getCell(0);
                    if (fechaPrincipalCell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(fechaPrincipalCell) || fechaPrincipalCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        Date fecha = fechaPrincipalCell.getDateCellValue();
                        fechasPrincipales2.add(fecha);
                    }
                    // Obtener la fecha inicial de la columna 
                    Cell fechaInicialCell = row.getCell(2);
                    if (fechaInicialCell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(fechaInicialCell) || fechaInicialCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        Date fecha = fechaInicialCell.getDateCellValue();
                        fechasIniciales2.add(fecha);
                    }

                    // Obtener la fecha final de la columna
                    Cell fechaFinalCell = row.getCell(4);
                    if (fechaFinalCell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(fechaFinalCell) || fechaFinalCell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        Date fecha = fechaFinalCell.getDateCellValue();
                        fechasFinales2.add(fecha);
                    }
                }
            }
        }
        
        // Mostrar resultados en la terminal
        for (int i = 0; i < numerosCoincidencias.size(); i++) {
            //System.out.println(numerosCoincidencias.get(i) + ": " + fechasPrincipales.get(i) + " // " + fechasIniciales.get(i) + " // " + fechasFinales.get(i));
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
            dataRow.createCell(0).setCellValue(numerosCoincidencias2.get(i));
            dataRow.createCell(1).setCellValue(fechasPrincipales2.get(i));
            dataRow.createCell(2).setCellValue(fechasIniciales2.get(i));
            dataRow.createCell(3).setCellValue(fechasFinales2.get(i));
        }
        
        // Nueva sección para encontrar la fecha objetivo
        Row row = ws1.getRow(1); // fila de interés es la 2 (índice 1)
        int colIndex = 7; // Columna H (índice 7)
        Cell cell;
        
        LocalDate fechaObjetivo;

        // Determinar el día de la semana de la fechaReferenciaPrueba
        DayOfWeek diaSemana = fechaReferenciaPrueba.getDayOfWeek();
        
        if (diaSemana == DayOfWeek.MONDAY) {
            // Si es lunes, restar dos días
            fechaObjetivo = fechaReferenciaPrueba.minusDays(2);
        } else if (diaSemana == DayOfWeek.TUESDAY ||
                   diaSemana == DayOfWeek.WEDNESDAY ||
                   diaSemana == DayOfWeek.THURSDAY ||
                   diaSemana == DayOfWeek.FRIDAY ||
                   diaSemana == DayOfWeek.SATURDAY) {
            // Si es martes, miércoles, jueves, viernes o sábado, restar un día
            fechaObjetivo = fechaReferenciaPrueba.minusDays(1);
        } else if (diaSemana == DayOfWeek.SUNDAY) {
            // Si es domingo, imprimir mensaje de error y salir del método
            System.out.println("No es posible ejecutar este proceso para los domingos.");
            return;
        } else {
            // Caso por defecto si por alguna razón no coincide con ningún día
            System.out.println("Error: Día de la semana no reconocido.");
            return;
        }

        //LocalDate fechaReferenciaPrueba = LocalDate.of(2024, 10, 03);
        //LocalDate fechaObjetivo = fechaReferenciaPrueba.minusDays(1); // Fecha objetivo (día anterior)
        String fechaObjetivoStr = fechaObjetivo.format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        Date fechaObjetivoEncontrada = null;        
        List<Object> coordenadas = new ArrayList<>();
        
        while (colIndex <= row.getLastCellNum()) {
            cell = row.getCell(colIndex);
            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                Date fechaCelda = cell.getDateCellValue();
                System.out.println(fechaCelda); 
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
                    
                    Cell celdaCoordenadas = (Cell) newSheet.getRow(0).createCell(4);
                    if (coordenadas.size() == 2) { // Verifica que la lista tenga el tamaño correcto
                        String coordenadasString = coordenadas.get(0) + "" + coordenadas.get(1);
                        System.out.println(coordenadasString);
                        celdaCoordenadas.setCellValue(coordenadasString);
                        System.out.println(celdaCoordenadas);
                    } else {
                        System.out.println("Error: La lista de coordenadas no tiene el tamaño correcto.");
                    }
                    System.out.print(fechaObjetivoEncontrada + " " + ((Object)fechaObjetivoEncontrada).getClass().getSimpleName());
                    //System.out.println(fechaReferenciaPrueba);
                    break; // Detener la iteración
                }
            }
            colIndex++;
        }
        System.out.println(coordenadas);
        
        
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
        
        String coordenadasString = coordenadas.get(0) + "" + coordenadas.get(1);
        Object coordenadasStringLen = length(coordenadasString);
        System.out.println(coordenadasStringLen);
        // Convertir la letra de la columna en un índice numérico
        
        int columnaIndex = 0;
        
        if (coordenadasStringLen.equals(2)) {
            columnaIndex = coordenadasString.charAt(0) - 'A';
            System.out.println("caso 1: " + columnaIndex);
        } else {
            // sumar los índices de las dos letras
            //int letra1 = coordenadasString.charAt(0) - 'A';
            int letra2 = coordenadasString.charAt(1) - 'A';
            columnaIndex = 26 + letra2;
            System.out.println("caso 2: " + columnaIndex);
        }
        
        for (int i = 2; i <= ws1.getLastRowNum(); i++){
            Row row2 = ws1.getRow(i);
            
            // Si la fila es nula, continuar con la siguiente
            if (row2 == null) {
                break;
            }
            // Leer la celda de la columna A (número de documento)
            Cell cellA = row2.getCell(0);
            if (cellA == null || cellA.getCellType() == Cell.CELL_TYPE_BLANK) {
                // Saltarse la celda si se encuentra una vacía
                continue;
            }
            
            double numeroDocumento = cellA.getNumericCellValue();
            
            // Crear la celda en la columna correspondiente si no existe
            Cell cellResultado = row2.getCell(columnaIndex);
            if (cellResultado == null) {
                cellResultado = row2.createCell(columnaIndex);
            }
            
            if (numDocNDAsistencia.contains(numeroDocumento)){
                cellResultado.setCellValue("#N/D");
                cellResultado.setCellStyle(centeredStyle);
            }
        }
        
        
        // Cargar el controlador JDBC para SQL Server (ajusta según tu versión)
        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");

        // URL de conexión, usuario y contraseña
        String url = "jdbc:sqlserver://192.168.1.3;databaseName=CASALIMPIA";
        String user = "fenix";
        String password = "Beck5388100NI";
        supernumerarios (url, user, password, wb1);
        
        Sheet ws3 = wb1.getSheet("Supernumerarios TCVA");
        
        // Obtener las columnas de interés como iteradores
        Iterator<Row> rowIterator3 = ws3.iterator();
        rowIterator3.next();
        // Llenar las listas con los números de documento 
        while (rowIterator3.hasNext()) {
            Row row2 = rowIterator3.next();
            String numeroDoc = row2.getCell(0).getStringCellValue();
            String horarioSupernumerarios = row2.getCell(3).getStringCellValue();
            numDocSupernumerarios.add(numeroDoc);
            horario.add(horarioSupernumerarios);
        }
        //System.out.println("Conjunto de números de supernumerarios:" + numDocSupernumerarios);
        
        // Convertir los Strings a Doubles
        List<Double> numerosDocSupernumerariosDouble = new ArrayList<>();
        List<Double> horarioDouble = new ArrayList<>();
        for (String numeroDoc : numDocSupernumerarios) {
            try {
                double numeroDocDouble = Double.parseDouble(numeroDoc);
                numerosDocSupernumerariosDouble.add(numeroDocDouble);
            } catch (NumberFormatException e) {
                System.err.println("Error al convertir '" + numeroDoc + "' a Double: " + e.getMessage());
            }
        }
        for (String horarios : horario) {
            try {
                double numeroHorarioDouble = Double.parseDouble(horarios);
                horarioDouble.add(numeroHorarioDouble);
            } catch (NumberFormatException e) {
                System.err.println("Error al convertir '" + horarios + "' a Double: " + e.getMessage());
            }
        }
        
        // Imprimir los números de supernumerarios como Doubles
        System.out.println("Conjunto de números de supernumerarios como Doubles:" + numerosDocSupernumerariosDouble);
        System.out.println("Conjunto de números de supernumerarios como Doubles:" + horarioDouble);
        
        
        
        int indice2 = 0;
        // Llenar las listas con los números de documento 
        // Iterar sobre las filas de ws1 comenzando desde la fila 3
        for (int i = 2; i <= ws1.getLastRowNum(); i++) {
            
            Row row2 = ws1.getRow(i);
            
            //Double numeroAsistencia = row2.getCell(0).getNumericCellValue();
            // Leer la celda de la columna A (número de documento)
            Cell cellA = row2.getCell(0);
            Double numeroDoc = cellA.getNumericCellValue();
            Cell cellE = row2.getCell(4);
            Cell cellF = row2.getCell(5);
            
            if (cellA == null || cellA.getCellType() == Cell.CELL_TYPE_BLANK) {
                indice2 ++;
                continue;
            }
            
            if (cellE == null || cellF == null) {
                cellE = row2.createCell(4);
                cellF = row2.createCell(5);
            }
            
            if (indice2 == 1){
                if (numerosDocSupernumerariosDouble.contains(numeroDoc)){
                    cellE.setCellValue("A");
                    cellE.setCellStyle(centeredStyle);
                    cellF.setCellValue(horarioDouble.get(i));
                    cellF.setCellStyle(centeredStyle);
                }
            }
        }
        System.out.println(indice2);
        /*
        FileOutputStream outputStream = new FileOutputStream("O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/ResultsAsistencias.xlsx");
        wb1.write(outputStream);
        outputStream.close();
        wb1.close();
        FileOutputStream outputStream2 = new FileOutputStream("O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/ResultsTurnos.xlsx");
        wb2.write(outputStream2);
        outputStream.close();
        wb2.close();
        //System.out.println("Archivo Excel creado exitosamente: resultados.xlsx");
        */
    }
    
    public static void main(String[] args) throws Exception {
        String inputFilePath1 = "O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/Asistencia Noviembre-2024-Turnos Horizen (2).xlsx";
        String inputFilePath2 = "O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/Reporte de 09 a 15 de noviembre.xlsx";
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
            //System.out.println("Archivo procesado exitosamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
