/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.turnoshorizen.procesamiento_hojas;
import com.casalimpia_app.turnoshorizen.procesamiento_hojas.tue_wed_thu_fri;
import static com.casalimpia_app.turnoshorizen.procesamiento_hojas.tue_wed_thu_fri.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
//import java.io.FileOutputStream;
import java.io.IOException;
//|import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
//import java.util.Iterator;
import java.util.List;
import static jdk.nashorn.internal.objects.NativeString.length;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
//import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author jcavilaa
 */
public class writeData {
    public static void validacionTurnos(Workbook wb1, Workbook wb2) throws IOException{
        Sheet ws1 = wb1.getSheet("TURNOS MES DE NOVIEMBRE");
        Sheet wsResultados = wb2.getSheet("Resultados");
        Sheet wsFilasCoincidentes = wb2.getSheet("FilasCoincidentes");
        Cell celdaCoordenadas = wsResultados.getRow(0).getCell(4);
        String columnaCoordenada = celdaCoordenadas.getStringCellValue().substring(0, 1);
        System.out.println("Valor de la columna tomada: " + columnaCoordenada);
        System.out.println("Valor de la celda coordenadas: " + celdaCoordenadas);
        
        // Convertir la letra de la columna en un índice numérico
        
        
        Object coordenadasStringLen = length(celdaCoordenadas.getStringCellValue());
        System.out.println(coordenadasStringLen);
        int columnaIndex = 0;
        
        if (coordenadasStringLen.equals(2)) {
            columnaIndex = celdaCoordenadas.getStringCellValue().charAt(0) - 'A';
            System.out.println("caso 1: " + columnaIndex);
        } else {
            // sumar los índices de las dos letras
            //int letra1 = coordenadasString.charAt(0) - 'A';
            int letra2 = celdaCoordenadas.getStringCellValue().charAt(1) - 'A';
            columnaIndex = 26 + letra2;
            System.out.println("caso 2: " + columnaIndex);
        }
        
        
        // Tomar el valor de la celda objetivo (fecha)
        Cell celdaCoordenadasObjetivo = ws1.getRow(1).getCell(columnaIndex);
        Date fechaObjetivo = celdaCoordenadasObjetivo.getDateCellValue();
        System.out.println("Fecha de la celda objetivo obtenida: " + fechaObjetivo);
        
        // Sacar los días de la semana de la fecha objetivo 
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(fechaObjetivo);
        int diaSemana = calendar.get(Calendar.DAY_OF_WEEK);

        
        // Crear listas para almacenar los números de documento
        List<Double> noActivaTurno = new ArrayList<>();
        List<Double> activaTurno = new ArrayList<>();

        if (diaSemana == Calendar.MONDAY ||
            diaSemana == Calendar.TUESDAY ||
            diaSemana == Calendar.WEDNESDAY ||
            diaSemana == Calendar.THURSDAY ||
            diaSemana == Calendar.FRIDAY) {
            
            escribirDatos(wsFilasCoincidentes, ws1, activaTurno, noActivaTurno, columnaIndex, fechaObjetivo, diaSemana);
            FileOutputStream outputStream = new FileOutputStream("O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/Results2.xlsx");
            wb1.write(outputStream);
            outputStream.close();
        } else if (diaSemana == Calendar.SATURDAY){
            //---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            int newColumnaIndex = columnaIndex + 1;
            // Iteración para setear las expertas a "Descanso" (SOLO SI LA FECHA OBJETIVO PERTENECE A UN DOMINGO)
            for (int i = 2; i <= ws1.getLastRowNum(); i++){
                Row row = ws1.getRow(i);

                if (row == null){
                    break;
                }

                // Obtener las cédulas para detener el ciclo 
                Cell cellA = row.getCell(0);

                if (cellA == null || cellA.getCellType() == Cell.CELL_TYPE_BLANK){
                    break;
                }

                //Celda en donde se escriben los datos
                Cell cellResultado = row.getCell(newColumnaIndex);
                //Si esta celda no existe, se crea una nueva
                if (cellResultado == null || cellResultado.getCellType() == Cell.CELL_TYPE_BLANK) {
                    cellResultado = row.createCell(newColumnaIndex);
                    cellResultado.setCellValue("Descanso");
                }
            } 
            /*
            // Crear un estilo con la fuente deseada
            Font fuenteCalibri = wb1.createFont();
            fuenteCalibri.setFontName("Calibri");
            fuenteCalibri.setFontHeightInPoints((short) 11);

            CellStyle estiloCalibri = wb1.createCellStyle();
            estiloCalibri.setFont(fuenteCalibri);

            // Aplicar el estilo a todas las celdas de la hoja
            for (Row fila : ws1 ) {
                for (Cell celda : fila) {
                    celda.setCellStyle(estiloCalibri);
                }
            }
            */
            //---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            escribirDatos(wsFilasCoincidentes, ws1, activaTurno, noActivaTurno, columnaIndex, fechaObjetivo, diaSemana);
            FileOutputStream outputStream = new FileOutputStream("O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/Results2.xlsx");
            wb1.write(outputStream);
            outputStream.close();
            wb1.close();
            
        } else {
            System.out.println("Error, el proceso solamente se ejecuta de lunes a sábado");
        }
    }
    
    public static void main(String[] args) throws Exception {
        String inputFilePath1 = "O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/ResultsAsistencias.xlsx";
        String inputFilePath2 = "O:/proyecto/Activacion-de-turnos/TurnosHorizen/src/main/java/com/casalimpia_app/turnoshorizen/ResultsTurnos.xlsx";
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

        validacionTurnos(wb1, wb2);
        /*
        wb1.write(new FileOutputStream(outputFilePath1));
        wb1.close();
        wb2.write(new FileOutputStream(outputFilePath2));
        wb2.close();
        */
        //System.out.println("Archivo procesado exitosamente.");
    }
    
}
