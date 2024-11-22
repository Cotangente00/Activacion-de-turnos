/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.turnoshorizen.model;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author jcavilaa
 */
public class consultasSIC {
    public static void supernumerarios (String url, String user, String password, Workbook wb) throws SQLException, Exception{
        try ( // Obtener una conexión
                    Connection connection = DriverManager.getConnection(url, user, password)) {
                System.out.println("Conexión establecida");
                try ( // Crear un statement para ejecutar consultas
                        Statement statement = connection.createStatement()) {
                    try (ResultSet visorSupernumerarios = statement.executeQuery("SELECT * FROM [CASALIMPIA].[pymesHogar].[visorReporteSupernumerarios] vs\n" +
                                                                                 "WHERE Coord = 'TCVA'")) {
                        Sheet ws = wb.createSheet("Supernumerarios TCVA");
                        
                        int rowNum = 1;
                        while (visorSupernumerarios.next()) {
                            Row row = ws.createRow(rowNum++);
                            row.createCell(0).setCellValue(visorSupernumerarios.getString("cedula"));
                            row.createCell(1).setCellValue(visorSupernumerarios.getString("nombre"));
                            row.createCell(2).setCellValue(visorSupernumerarios.getString("apellido"));
                            row.createCell(3).setCellValue(visorSupernumerarios.getString("estado"));
                            row.createCell(3).setCellValue(visorSupernumerarios.getString("Horario"));
                        }
                        
                        
                        
                    }
                }
        }
    }
}   