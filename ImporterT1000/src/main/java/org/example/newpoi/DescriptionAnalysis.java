package org.example.newpoi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import javafx.scene.control.TextArea;

public class DescriptionAnalysis {
    public static void main(String[] args) throws IOException {
    }
    public static void DescriptAnalis_Func(String AbsolutePath_DescAnals_Descr,
                                           String AbsolutePath_DescAnals_DB,
                                           String tField_DescAnals_Descr,
                                           TextArea logMassege) throws IOException{

        String fileNameXlsx = tField_DescAnals_Descr;
        String fileSavePath = AbsolutePath_DescAnals_Descr.replace(fileNameXlsx, "");
        String fileName = fileNameXlsx.replace(".xlsx", "");
        String newSaveFile = fileSavePath + fileName + "_NEW.xlsx";

        System.out.println("Загрузка книги...");
        logMassege.appendText("Загрузка книги... \n");

        //открытие книг
        Workbook wb_Descr  = WorkbookFactory.create(new FileInputStream(AbsolutePath_DescAnals_Descr));
        Workbook wb_DB = WorkbookFactory.create(new FileInputStream(AbsolutePath_DescAnals_DB));

        //получение листов Data Base
        Sheet sheetDB_MMY_G_KeyWords = wb_DB.getSheet("MMY_G_KeyWords");
        Sheet sheetDB_MMY_G = wb_DB.getSheet("MMY_G");
        Sheet sheetDB_Params = wb_DB.getSheet("Parameters");

        //получение листов Description
        wb_Descr.createSheet("MMY_G");
        wb_Descr.createSheet("Parameters");

        Sheet sheetDescr = wb_Descr.getSheetAt(0);
        Sheet sheetDescrMMY_G = wb_Descr.getSheet("MMY_G");
        Sheet sheetDescrParams = wb_Descr.getSheet("Parameters");

        //Получение колонок с Description
        int Descr_BrandSku = DescrAnalysisFunc.getColNameIndex(sheetDescr.getRow(0), "BrandSku");
        int Descr_Brand = DescrAnalysisFunc.getColNameIndex(sheetDescr.getRow(0), "Brand");
        int Descr_Sku = DescrAnalysisFunc.getColNameIndex(sheetDescr.getRow(0), "Sku");
        int Descr_Descr = DescrAnalysisFunc.getColNameIndex(sheetDescr.getRow(0), "Description");

        int descr_MMY_G_Val = Descr_Descr + 1;
        int descr_MMY_G_EngFamily = Descr_Descr + 2;
        int descr_MMY_G_EngVersion = Descr_Descr + 3;
        int descr_MMY_G_Name = Descr_Descr + 4;
        int descr_MMY_G_Descr = Descr_Descr + 5;

        //запись колонок в лист sheetDescrMMY_G
        sheetDescrMMY_G.createRow(0).createCell(Descr_BrandSku).setCellValue("BrandSku");
        sheetDescrMMY_G.getRow(0).createCell(Descr_Brand).setCellValue("Brand");
        sheetDescrMMY_G.getRow(0).createCell(Descr_Sku).setCellValue("Sku");
        sheetDescrMMY_G.getRow(0).createCell(Descr_Descr).setCellValue("Description");
        sheetDescrMMY_G.getRow(0).createCell(descr_MMY_G_Val).setCellValue("MMY_G Value");
        sheetDescrMMY_G.getRow(0).createCell(descr_MMY_G_EngFamily).setCellValue("Engine Family");
        sheetDescrMMY_G.getRow(0).createCell(descr_MMY_G_EngVersion).setCellValue("Engine Version");
        sheetDescrMMY_G.getRow(0).createCell(descr_MMY_G_Name).setCellValue("MMY Group Name");
        sheetDescrMMY_G.getRow(0).createCell(descr_MMY_G_Descr).setCellValue("MMY Description");

        //запись колонок в лист sheetDescrParams
        sheetDescrParams.createRow(0).createCell(Descr_BrandSku).setCellValue("BrandSku");
        sheetDescrParams.getRow(0).createCell(Descr_Brand).setCellValue("Brand");
        sheetDescrParams.getRow(0).createCell(Descr_Sku).setCellValue("Sku");
        sheetDescrParams.getRow(0).createCell(Descr_Descr).setCellValue("Description");

        //Получение колонок с Data Base
        int db_MMY_G_KeyWord = DescrAnalysisFunc.getColNameIndex(sheetDB_MMY_G_KeyWords.getRow(0), "MMY_G_KeyWord");
        int db_MMY_G_KeyWord_Value = DescrAnalysisFunc.getColNameIndex(sheetDB_MMY_G_KeyWords.getRow(0), "MMY_G Value");

        int db_MMY_G_Value = DescrAnalysisFunc.getColNameIndex(sheetDB_MMY_G.getRow(0), "MMY_G Value");
        int db_MMY_G_EngFamily = DescrAnalysisFunc.getColNameIndex(sheetDB_MMY_G.getRow(0), "Engine Family");
        int db_MMY_G_EngVersion = DescrAnalysisFunc.getColNameIndex(sheetDB_MMY_G.getRow(0), "Engine Version");
        int db_MMY_G_Name = DescrAnalysisFunc.getColNameIndex(sheetDB_MMY_G.getRow(0), "MMY Group Name");
        int db_MMY_G_Descr = DescrAnalysisFunc.getColNameIndex(sheetDB_MMY_G.getRow(0), "MMY Description");

        int db_Params_SearchParam = DescrAnalysisFunc.getColNameIndex(sheetDB_Params.getRow(0), "SearchParameter");
        int db_Params_CorrectParam = DescrAnalysisFunc.getColNameIndex(sheetDB_Params.getRow(0), "CorrectParameter");
        int db_Params_ParameterName = DescrAnalysisFunc.getColNameIndex(sheetDB_Params.getRow(0), "ParameterName");

        //lastRow
        int lastRowDescr_Descr = sheetDescr.getLastRowNum()+1;
        int lastRowDB_MMY_G_KeyWord = sheetDB_MMY_G_KeyWords.getLastRowNum()+1;
        int lastRowDB_MMY_G = sheetDB_MMY_G.getLastRowNum()+1;
        int lastRowDB_Params = sheetDB_Params.getLastRowNum()+1;


        for(int i = 1; i < lastRowDescr_Descr; i++){
            System.out.println("   ---      Обработка строки " + i + " из " + (lastRowDescr_Descr - 1) + "     ---");
            logMassege.appendText(" ---  Обработка строки " + i + " из " + (lastRowDescr_Descr - 1) + " \n");
            //значение с Description текущей строки
            String DescrVal = DescrAnalysisFunc.getCellValue(sheetDescr, i, Descr_Descr);
            String descr_BrandSkuVal = DescrAnalysisFunc.getCellValue(sheetDescr, i, Descr_BrandSku);
            String descr_BrandVal = DescrAnalysisFunc.getCellValue(sheetDescr, i, Descr_Brand);
            String descr_SkuVal = DescrAnalysisFunc.getCellValue(sheetDescr, i, Descr_Sku);
            //System.out.println("Description = "+ DescrVal);

            //Перенос на лист Parameters значения колонок BrandSku Brand Sku Description
            sheetDescrParams.createRow(i).createCell(Descr_BrandSku).setCellValue(descr_BrandSkuVal);
            sheetDescrParams.getRow(i).createCell(Descr_Brand).setCellValue(descr_BrandVal);
            sheetDescrParams.getRow(i).createCell(Descr_Sku).setCellValue(descr_SkuVal);
            sheetDescrParams.getRow(i).createCell(Descr_Descr).setCellValue(DescrVal);

            //поиск мму групп
            for (int a = 1; a < lastRowDB_MMY_G_KeyWord; a++){
                String db_MMY_G_KeyWordVal = DescrAnalysisFunc.getCellValue(sheetDB_MMY_G_KeyWords, a, db_MMY_G_KeyWord);
                //System.out.println("MMY_G_KeyWordVal = "+ db_MMY_G_KeyWordVal);
                if(DescrVal.toLowerCase().contains(db_MMY_G_KeyWordVal.toLowerCase())){

                    //System.out.println("Description2 = "+ DescrVal);
                    //System.out.println(" найдено ММУ Группу : " + db_MMY_G_KeyWordVal);
                    String MMY_G_KeyWord_Val = DescrAnalysisFunc.getCellValue(sheetDB_MMY_G_KeyWords, a, db_MMY_G_KeyWord_Value);
                    //System.out.println("MMY_G_KeyWord_Val = "+ MMY_G_KeyWord_Val);

                    //поиск развернутой мму группЫ
                    for (int b = 1; b < lastRowDB_MMY_G; b++){
                        String MMY_G_Val = DescrAnalysisFunc.getCellValue(sheetDB_MMY_G, b, db_MMY_G_Value);
                        if (MMY_G_Val.equalsIgnoreCase(MMY_G_KeyWord_Val)){
                            //System.out.println(" найдено мму группу : "+ MMY_G_Val+" в строке " + b);
                            //Получение значений с колонок развернутого ММУ
                            String db_MMY_G_EngFamilyVal = DescrAnalysisFunc.getCellValue(sheetDB_MMY_G, b, db_MMY_G_EngFamily);
                            String db_MMY_G_EngVersionVal = DescrAnalysisFunc.getCellValue(sheetDB_MMY_G, b, db_MMY_G_EngVersion);
                            String db_MMY_G_NameVal = DescrAnalysisFunc.getCellValue(sheetDB_MMY_G, b, db_MMY_G_Name);
                            String db_MMY_G_DescrVal = DescrAnalysisFunc.getCellValue(sheetDB_MMY_G, b, db_MMY_G_Descr);

                            //Получение номер строки куда нужно произвести запись
                            int lastRowDescrMMY_G = sheetDescrMMY_G.getLastRowNum()+1;

                            //Запись
                            sheetDescrMMY_G.createRow(lastRowDescrMMY_G).createCell(Descr_BrandSku).setCellValue(descr_BrandSkuVal);
                            sheetDescrMMY_G.getRow(lastRowDescrMMY_G).createCell(Descr_Brand).setCellValue(descr_BrandVal);
                            sheetDescrMMY_G.getRow(lastRowDescrMMY_G).createCell(Descr_Sku).setCellValue(descr_SkuVal);
                            sheetDescrMMY_G.getRow(lastRowDescrMMY_G).createCell(Descr_Descr).setCellValue(DescrVal);

                            sheetDescrMMY_G.getRow(lastRowDescrMMY_G).createCell(descr_MMY_G_Val).setCellValue(MMY_G_Val);
                            sheetDescrMMY_G.getRow(lastRowDescrMMY_G).createCell(descr_MMY_G_EngFamily).setCellValue(db_MMY_G_EngFamilyVal);
                            sheetDescrMMY_G.getRow(lastRowDescrMMY_G).createCell(descr_MMY_G_EngVersion).setCellValue(db_MMY_G_EngVersionVal);
                            sheetDescrMMY_G.getRow(lastRowDescrMMY_G).createCell(descr_MMY_G_Name).setCellValue(db_MMY_G_NameVal);
                            sheetDescrMMY_G.getRow(lastRowDescrMMY_G).createCell(descr_MMY_G_Descr).setCellValue(db_MMY_G_DescrVal);

                        }
                    }

                    break;

                }
            }

            for (int d = 1; d < lastRowDB_Params; d++){

                String searchParameterVal = DescrAnalysisFunc.getCellValue(sheetDB_Params, d, db_Params_SearchParam);
                //System.out.println("SearchParameterVal = "+ SearchParameterVal);
                if(DescrVal.toLowerCase().contains(searchParameterVal.toLowerCase())){
                    //System.out.println("SearchParameter Найдено : "+searchParameterVal);
                    String correctParameterVal = DescrAnalysisFunc.getCellValue(sheetDB_Params, d, db_Params_CorrectParam);
                    //System.out.println("correctParameterVal Найдено : "+correctParameterVal);
                    String parameterNameVal = DescrAnalysisFunc.getCellValue(sheetDB_Params, d, db_Params_ParameterName);

                    //Получение индекса колонки parameterNameVal
                    int parameterNameIndex = DescrAnalysisFunc.getParameterNameIndex(sheetDescrParams, sheetDescrParams.getRow(0), parameterNameVal);
                   // System.out.println("parameterNameIndex = " + parameterNameIndex);
                    //System.out.println("i index  = " + i);

                    //System.out.println("correctParameterVal Найдено : "+correctParameterVal);
                    //Запись
                    sheetDescrParams.getRow(i).createCell(parameterNameIndex).setCellValue(correctParameterVal);

                }
            }
        }

        DescrAnalysisFunc.writeWorkbook(wb_Descr,newSaveFile);
        logMassege.appendText(" ---  Done!  ---\n");
        logMassege.appendText("           ---  Hasta la vista, baby! \n\n");
        System.out.println(" ---  Done!  ---");
        System.out.println("           ---  Hasta la vista, baby!");

    }
}
