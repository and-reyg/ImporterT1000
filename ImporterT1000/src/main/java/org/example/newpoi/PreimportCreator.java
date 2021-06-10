package org.example.newpoi;

import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Optional;
import javafx.scene.control.TextArea;

public class PreimportCreator {
    public static void main(String[] args) {
    }
    public static void PreimportCreator_Func(String AbsolutePath_Preimp_Descript,
                                             String AbsolutePath_Preimp_DB,
                                             String text_tField_Preimp_Descript,
                                             TextArea logMassege) throws IOException {

        //Перенос колонок brandKey, pKey, specification на лист virtList
        String fileNameXlsx = text_tField_Preimp_Descript;
        String fileSavePath = AbsolutePath_Preimp_Descript.replace(fileNameXlsx, "");
        String fileName = fileNameXlsx.replace(".xlsx", "");
        String newSaveFile = fileSavePath + fileName + "_NEW.xlsx";

        System.out.println("Загрузка книги...");
        logMassege.appendText("Загрузка книги... \n");

        Workbook wb_Preimport = WorkbookFactory.create(new FileInputStream(AbsolutePath_Preimp_Descript));
        Workbook wb_DB = WorkbookFactory.create(new FileInputStream(AbsolutePath_Preimp_DB));

        Sheet sheetKeyWords = wb_DB.getSheet("KeyWords");
        Sheet sheetMyBrands = wb_DB.getSheet("MyBrands");
        Sheet sheetSeries = wb_DB.getSheet("Series");
        Sheet sheetPreimptParams = wb_DB.getSheet("PreimportParams");
        Sheet sheetTemplateComp = wb_DB.getSheet("TemplateComponent");
        Sheet sheetPreimport = wb_Preimport.getSheetAt(0);

        int pDescription = PreimportFunc.getColNameIndex(sheetPreimport.getRow(0), "PRODUCT DESCRIPTION");
        int preimportBrand = PreimportFunc.getColNameIndex(sheetPreimport.getRow(0), "BRAND");

        int myBrands = PreimportFunc.getColNameIndex(sheetMyBrands.getRow(0), "Brand");
        int preimpParamBrand = PreimportFunc.getColNameIndex(sheetPreimptParams.getRow(0), "Brand");

        int lastRowPreimport = sheetPreimport.getLastRowNum() + 1;
        int lastRowMyBrands = sheetMyBrands.getLastRowNum() + 1;
        int lastRowKeyW = sheetKeyWords.getLastRowNum() + 1;
        int lastRowSeries = sheetSeries.getLastRowNum() + 1;
        int lastRowPreParams = sheetPreimptParams.getLastRowNum() + 1;
        int lastRowTemplateComp = sheetTemplateComp.getLastRowNum() + 1;

        int iMyBrands = pDescription + 1;
        int iMatch = pDescription + 2;
        int iSubtype = pDescription + 3;
        int iSeries = pDescription + 4;
        int iProductName = pDescription + 5;
        int iPartTerminologyName = pDescription + 6;
        int iTemplate = pDescription + 7;
        int iPurpose = pDescription + 8;
        int iShortDescr = pDescription + 9;
        int iDetailedDescr = pDescription + 10;

        sheetPreimport.getRow(0).createCell(iMyBrands).setCellValue("MyBrands");
        sheetPreimport.getRow(0).createCell(iMatch).setCellValue("Match");
        sheetPreimport.getRow(0).createCell(iSubtype).setCellValue("Subtype");
        sheetPreimport.getRow(0).createCell(iSeries).setCellValue("Series");
        sheetPreimport.getRow(0).createCell(iProductName).setCellValue("Product Name");
        sheetPreimport.getRow(0).createCell(iPartTerminologyName).setCellValue("PartTerminologyName");
        sheetPreimport.getRow(0).createCell(iTemplate).setCellValue("Template");
        sheetPreimport.getRow(0).createCell(iPurpose).setCellValue("For_Purpose");
        sheetPreimport.getRow(0).createCell(iShortDescr).setCellValue("Short_Descr");
        sheetPreimport.getRow(0).createCell(iDetailedDescr).setCellValue("Detailed_Descr");

        for (int i = 1; i < lastRowPreimport; i++) {
            System.out.println("       ---    Обработка  " + i + " строки из " + (lastRowPreimport - 1) + "  --- ");
            logMassege.appendText(" ---    Обработка  " + i + " строки из " + (lastRowPreimport - 1) + " \n");
            //Получение названия бренда с преимпорта
            String preimportBrandVal = PreimportFunc.getCellValue(sheetPreimport, i, preimportBrand);

            //Значение Product Description
            String productDescription = PreimportFunc.getCellValue(sheetPreimport, i, pDescription);

            for (int a = 1; a < lastRowMyBrands; a++) {
                //Получение названия бренда с MyBrands
                String myBrandVal = PreimportFunc.getCellValue(sheetMyBrands, a, myBrands);

                if (preimportBrandVal.equalsIgnoreCase(myBrandVal)) {
                    sheetPreimport.getRow(i).createCell(iMyBrands).setCellValue("My Brand");

                    //Поиск в дескрипшин ключевых слов
                    for (int b = 1; b < lastRowKeyW; b++) {
                        String KeyWordVal = PreimportFunc.getCellValue(sheetKeyWords, b, 0);

                        if (productDescription.toLowerCase().contains(KeyWordVal.toLowerCase())) {

                            //Если ключевое слово Not My
                            String keywordSubtype = PreimportFunc.getCellValue(sheetKeyWords, b, 2);
                            if (keywordSubtype.equalsIgnoreCase("Not My")) {
                                sheetPreimport.getRow(i).createCell(iSubtype).setCellValue("Not My");
                                break;
                            }

                            //Передача правильного значеия из ключевых слов
                            String ProdName = PreimportFunc.getCellValue(sheetKeyWords, b, 1);
                            //Поиск в product description   Series
                            String Series = PreimportFunc.FindSeriesFunc(sheetSeries, productDescription, lastRowSeries);
                            //Поиск в product description TemplateComponent и запись найденного в массив
                            ArrayList<String> templtComponentArr = new ArrayList<>(PreimportFunc.FindTemplateComponentFunc(sheetTemplateComp, productDescription, lastRowTemplateComp));
                            //Поиск в каких строках листа PreimportParams есть проверяемый бренд
                            ArrayList<Integer> currentBrandinPreimptParamRowArr = new ArrayList<>(PreimportFunc.CurrentBrandinPreimptParamRows(sheetPreimptParams, preimportBrandVal, lastRowPreParams, preimpParamBrand));

                            //генерация темплейта
                            ArrayList<String> arrTemplateGenerator = new ArrayList<>(PreimportFunc.TemplateGenerator(preimportBrandVal, ProdName, templtComponentArr));

                            //Проверка на совпадение 100%
                            int match100Val = PreimportFunc.Match100Func(arrTemplateGenerator,
                                    sheetPreimptParams,
                                    sheetPreimport,
                                    pDescription,
                                    i,
                                    ProdName,
                                    Series,
                                    currentBrandinPreimptParamRowArr);
                            if (match100Val == 100) {
                                break;
                            }

                            //Проверка на совпадение 80%
                            //генерация темплейта регулярки
                            ArrayList<String> arrTemplateGeneratorReg = new ArrayList<>(PreimportFunc.TemplateGeneratorReg(preimportBrandVal, ProdName, templtComponentArr));

                            int match70Val = PreimportFunc.Match70Func(arrTemplateGeneratorReg,
                                    sheetPreimptParams,
                                    sheetPreimport,
                                    pDescription,
                                    i,
                                    ProdName,
                                    Series,
                                    currentBrandinPreimptParamRowArr);
                            if (match70Val == 70) {
                                break;
                            }
                            //Проверка на совпадение 50%
                            int match50Val = PreimportFunc.Match50Func(currentBrandinPreimptParamRowArr,
                                    sheetPreimptParams,
                                    sheetPreimport,
                                    pDescription,
                                    i,
                                    ProdName,
                                    Series);
                            if (match50Val == 50) {
                                break;
                            }

                            //совпадение 10% иначе KEYWORD не найдено
                            //Запись
                            sheetPreimport.getRow(i).createCell(iMatch).setCellValue("10%");
                            sheetPreimport.getRow(i).createCell(iSubtype).setCellValue(keywordSubtype);
                            sheetPreimport.getRow(i).createCell(iSeries).setCellValue(Series);
                            sheetPreimport.getRow(i).createCell(iProductName).setCellValue(ProdName);


                        }
                    }
                    break;
                } else {
                    sheetPreimport.getRow(i).createCell(iMyBrands).setCellValue("Not My");
                }
            }

        }

        //запись
        PreimportFunc.writeWorkbook(wb_Preimport, newSaveFile);
        System.out.println("     ---  Done!  ---   ");
        System.out.println("           ---  Hasta la vista, baby! \n\n");
        logMassege.appendText("     ---  Done!  ---   \n");
        logMassege.appendText("           ---  Hasta la vista, baby! \n\n");
    }
}

