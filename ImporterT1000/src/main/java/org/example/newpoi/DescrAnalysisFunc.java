package org.example.newpoi;

import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Optional;

public class DescrAnalysisFunc {

    //Получение номера колонки
    public static int getColNameIndex(Row row, String columnName) {
        int i = 0;
        for (Cell cell : row) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return i;
            }
            i++;
        }
        throw new RuntimeException("ColmName index not found");
    }
    //Получение значений ячейки
    public static String getCellValue(Sheet sheet, int rowNumber, int cellNumber) {
        Cell cell = getCellOrNull(sheet, rowNumber, cellNumber);
        String result = "";
        if(cell == null) {
            result = "";
        }else if(cell.getCellType() == CellType.BLANK){
            result = "";
        }else if(cell.getCellType() == CellType.STRING){
            result = cell.getStringCellValue();
        }else if(cell.getCellType() == CellType.NUMERIC) {
            result = String.valueOf((int) cell.getNumericCellValue());
        }else {
            result = "";
        }
        return result;
    }

    public static Cell getCellOrNull(Sheet sheet, int rowNumber, int cellNumber) {
        return Optional.ofNullable(sheet.getRow(rowNumber))
                .map(row -> row.getCell(cellNumber))
                .orElse(null);
    }
    //Запись книги
    public static void writeWorkbook(Workbook workbook, String fileName){
        try{
            FileOutputStream output = new FileOutputStream(fileName);
            workbook.write(output);
            output.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    //Получение индекса колонки если она есть если нету то создает колонку и возвращает индекс
    public static int getParameterNameIndex(Sheet sheetDescrParams, Row row, String ParameterName) {
        int x = 0;
        int index = 0;
        String isParametrNameInSheet = "";
        for (Cell cell : row) {
            //если есть колонка
            if (cell.getStringCellValue().equalsIgnoreCase(ParameterName)) {
                System.out.println("ParametrName "+ ParameterName + " Найдено  индекс = "+ x);
                isParametrNameInSheet = "Yes";
                index = x;
            }
            x++;
        }
        System.out.println(" x = "+x);

        if(isParametrNameInSheet.equalsIgnoreCase("Yes")){
            return index;
        }else{
            //получение индекса последней колонки
            sheetDescrParams.getRow(0).createCell((x)).setCellValue(ParameterName);
            System.out.println("ParametrName "+ ParameterName + " НЕ найдено  индекс = "+ (x));
            index = x;
            return index;
        }

    }
}

