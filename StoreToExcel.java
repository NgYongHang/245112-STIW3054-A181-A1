/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.assignment1;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author User
 */
public class StoreToExcel {

    private static final String CONSTANT_EXCEL = "C:\\Users\\User\\Desktop\\Trivia.xls";
    private static final HSSFWorkbook myWorkBook = new HSSFWorkbook();
    private static final HSSFSheet mySheet = myWorkBook.createSheet();

    /**
     *
     * @param element1
     * @param element2
     * @param rowNum
     * @throws java.io.IOException
     */
    public static void excelLog(String element1, String element2, int rowNum) throws IOException {

        HSSFRow myRow;
        HSSFCell myCell;
        String excelData[][] = new String[1][2];
        excelData[0][0] = element1;
        excelData[0][1] = element2;

        myRow = mySheet.createRow(rowNum);
        for (int cellNum = 0; cellNum < 2; cellNum++) {
            myCell = myRow.createCell(cellNum);
            myCell.setCellValue(excelData[0][cellNum]);

            if (rowNum == 0) {
                Row row = mySheet.getRow(rowNum);
                CellStyle style = myWorkBook.createCellStyle();
                Font font = myWorkBook.createFont();
                font.setFontHeightInPoints((short) 12);
                font.setFontName(HSSFFont.FONT_ARIAL);
                font.setBold(true);
                font.setItalic(false);
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setFont(font);
                row.getCell(0).setCellStyle(style);

            }
        }

        for (int i = 0; i < 25; i++) {
            mySheet.autoSizeColumn(i);
        }

        try (FileOutputStream write = new FileOutputStream(CONSTANT_EXCEL)) {
            myWorkBook.write(write);
        }
           
    }

}
