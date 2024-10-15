/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.turnoshorizen.procesamiento_hojas;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class writeData {
    public static void validacionTurnos(Workbook wb1, Workbook wb2){
        Sheet ws1 = wb1.getSheet("TURNOS MES DE OCTUBRE");
        Sheet ws2 = wb2.getSheetAt(2);
        
        
    }
}
