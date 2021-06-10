package org.example.newpoi;

import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Optional;

public class PreimportFunc {


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

    //Поиск Series
    public static String FindSeriesFunc(Sheet sheetSeries, String productDescription, int lastRowSeries){
        String Series = "";
        for(int c = 1; c < lastRowSeries; c++){
            String SeriesVal = PreimportFunc.getCellValue(sheetSeries, c, 0);
            if (productDescription.toLowerCase().contains(SeriesVal.toLowerCase())) {
                Series = PreimportFunc.getCellValue(sheetSeries, c, 1);
                break;
            }else{
                Series = "";
            }
        }
        //System.out.println("Series = "+ Series);
        return Series;
    }
    //Поиск в product description TemplateComponent и запись найденного в массив
    public static ArrayList<String> FindTemplateComponentFunc(Sheet sheetTemplateComp, String productDescription, int lastRowTemplateComp){
        ArrayList<String> templtComponentArr = new ArrayList<>();
        for(int y = 1; y < lastRowTemplateComp; y++){
            String templComponentVal = PreimportFunc.getCellValue(sheetTemplateComp,y,0);
            if (productDescription.toLowerCase().contains(templComponentVal.toLowerCase())){
                String templComponentValcorrect = PreimportFunc.getCellValue(sheetTemplateComp, y,1);
                templtComponentArr.add(templComponentValcorrect);
            }
        }
        return templtComponentArr;
    }
    //Поиск в каких строках листа PreimportParams есть проверяемый бренд
    public static ArrayList<Integer> CurrentBrandinPreimptParamRows (Sheet sheetPreimptParams, String preimportBrandVal, int lastRowPreParams, int preimpParamBrand){
        ArrayList<Integer> currentBrandinPreimptParamRowArr = new ArrayList<>();
        for(int y = 1; y < lastRowPreParams; y++){
            String preimpParamBrandVal = getCellValue(sheetPreimptParams, y, preimpParamBrand);
            if(preimpParamBrandVal.equalsIgnoreCase(preimportBrandVal)){
                currentBrandinPreimptParamRowArr.add(y);
            }
        }

        return currentBrandinPreimptParamRowArr;
    }
    //Генерация темплейта для точного совпадения
    public static ArrayList<String> TemplateGenerator (String preimportBrandVal, String ProdName, ArrayList<String> templtComponentArr){
        ArrayList<String> arrTemplateGenerator = new ArrayList<>();
        String brandL = preimportBrandVal.toLowerCase().replace("®", "");

        //Создание продукт нейм с окончанием S и без S
        int ProdNameLength = ProdName.length();
        String ProdNameLastS = ProdName.substring(ProdNameLength-1);

        String ProdNameL = ProdName.toLowerCase();
        String ProdNameSL = ProdName.toLowerCase();

        if(ProdNameLastS.equalsIgnoreCase("s")){
            ProdNameL = ProdNameL.substring(0, (ProdNameLength-1));
        }else {
            ProdNameSL = ProdName.toLowerCase() + "s";
        }

        //Создание продукт нейм без Set и Kit
        String ProdName_wo_SetKit = ProdName.toLowerCase();
        if(ProdName_wo_SetKit.toLowerCase().contains(" set")){
            ProdName_wo_SetKit = ProdName_wo_SetKit.toLowerCase().replace(" set","");
        }else if(ProdName_wo_SetKit.toLowerCase().contains(" kit")){
            ProdName_wo_SetKit = ProdName_wo_SetKit.toLowerCase().replace(" kit","");
        }

        int templtComponentArrSize = templtComponentArr.size();

        //генерация темплейтов
        String compnt1 = "";
        String compnt2 = "";
        String compnt3 = "";

        if(templtComponentArrSize == 3){
            compnt1 = templtComponentArr.get(0).toLowerCase();
            compnt2 = templtComponentArr.get(1).toLowerCase();
            compnt3 = templtComponentArr.get(2).toLowerCase();

            //большой круг
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ compnt3 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ compnt3 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ compnt3 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt3 +" "+ compnt2 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt3 +" "+ compnt2 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt3 +" "+ compnt2 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt1 +" "+ compnt2 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt1 +" "+ compnt2 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt1 +" "+ compnt2 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt2 +" "+ compnt1 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt2 +" "+ compnt1 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt2 +" "+ compnt1 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt3 +" "+ compnt1 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt3 +" "+ compnt1 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt3 +" "+ compnt1 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ compnt3 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ compnt3 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ compnt3 +" "+ ProdName_wo_SetKit);

            //2 елемента круг 1 и 2
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ ProdName_wo_SetKit);

            //2 елемент круг есть 1 и 3
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt3 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt3 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt3 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt1 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt1 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt1 +" "+ ProdName_wo_SetKit);

            //2 елемент круг есть 2 и 3
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt3 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt3 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt3 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt2 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt2 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ compnt2 +" "+ ProdName_wo_SetKit);

            //1 елемент круг 1 / 2 / 3
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt3 +" "+ ProdName_wo_SetKit);

            // без компонентов
            arrTemplateGenerator.add(brandL +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ ProdName_wo_SetKit);


        }else if(templtComponentArrSize == 2){
            compnt1 = templtComponentArr.get(0).toLowerCase();
            compnt2 = templtComponentArr.get(1).toLowerCase();

            //2 елемента круг 1 и 2
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ compnt2 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ compnt1 +" "+ ProdName_wo_SetKit);

            //1 елемент круг 1 / 2
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdName_wo_SetKit);

            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt2 +" "+ ProdName_wo_SetKit);

            // без компонентов
            arrTemplateGenerator.add(brandL +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ ProdName_wo_SetKit);

        }else if(templtComponentArrSize == 1){
            compnt1 = templtComponentArr.get(0).toLowerCase();

            //1 елемент круг 1
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ compnt1 +" "+ ProdName_wo_SetKit);

            // без компонентов
            arrTemplateGenerator.add(brandL +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ ProdName_wo_SetKit);

        }else{
            // без компонентов
            arrTemplateGenerator.add(brandL +" "+ ProdNameSL);
            arrTemplateGenerator.add(brandL +" "+ ProdNameL);
            arrTemplateGenerator.add(brandL +" "+ ProdName_wo_SetKit);
        }

        return arrTemplateGenerator;
    }

    //Генерация темплейта для точного совпадения REGULAR
    public static ArrayList<String> TemplateGeneratorReg (String preimportBrandVal, String ProdName, ArrayList<String> templtComponentArr){
        ArrayList<String> arrTemplGeneratorRegex = new ArrayList<>();
        String brandL = preimportBrandVal.toLowerCase().replace("®", "");

        //Создание продукт нейм с окончанием S и без S
        int ProdNameLength = ProdName.length();

        String ProdNameLastS = ProdName.substring(ProdNameLength-1);

        String ProdNameL = ProdName.toLowerCase();

        if(ProdNameLastS.equalsIgnoreCase("s")){
            ProdNameL = ProdNameL.substring(0, (ProdNameLength-1));
        }
        //Создание продукт нейм без Set и Kit
        String ProdName_wo_SetKit = ProdNameL;
        if(ProdName_wo_SetKit.toLowerCase().contains(" set")){
            ProdName_wo_SetKit = ProdName_wo_SetKit.toLowerCase().replace(" set","");
        }else if(ProdName_wo_SetKit.toLowerCase().contains(" kit")){
            ProdName_wo_SetKit = ProdName_wo_SetKit.toLowerCase().replace(" kit","");
        }
        int templtComponentArrSize = templtComponentArr.size();

        //генерация темплейтов
        String compnt1 = "";
        String compnt2 = "";
        String compnt3 = "";

        if(templtComponentArrSize == 3){
            compnt1 = templtComponentArr.get(0).toLowerCase();
            compnt2 = templtComponentArr.get(1).toLowerCase();
            compnt3 = templtComponentArr.get(2).toLowerCase();

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt1 +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt1 +"(.*)"+ ProdName_wo_SetKit +"(.*)");

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt2 +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt2 +"(.*)"+ ProdName_wo_SetKit +"(.*)");

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt3 +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt3 +"(.*)"+ ProdName_wo_SetKit +"(.*)");

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ ProdName_wo_SetKit +"(.*)");

        } else if(templtComponentArrSize == 2){
            compnt1 = templtComponentArr.get(0).toLowerCase();
            compnt2 = templtComponentArr.get(1).toLowerCase();

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt1 +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt1 +"(.*)"+ ProdName_wo_SetKit +"(.*)");

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt2 +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt2 +"(.*)"+ ProdName_wo_SetKit +"(.*)");

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ ProdName_wo_SetKit +"(.*)");
        } else if(templtComponentArrSize == 1){
            compnt1 = templtComponentArr.get(0).toLowerCase();

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt1 +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ compnt1 +"(.*)"+ ProdName_wo_SetKit +"(.*)");

            arrTemplGeneratorRegex.add(brandL +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ ProdName_wo_SetKit +"(.*)");

        } else {
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ ProdNameL);
            arrTemplGeneratorRegex.add(brandL +"(.*)"+ ProdName_wo_SetKit +"(.*)");
        }

        return arrTemplGeneratorRegex;
    }

    //Проверка на совпадение 100%
    public static Integer Match100Func(ArrayList<String> arrTemplateGenerator,
                                       Sheet sheetPreimptParams,
                                       Sheet sheetPreimport,
                                       int pDescription,
                                       int i,
                                       String ProdName,
                                       String Series,
                                       ArrayList<Integer> currentBrandinPreimptParamRowArr){
        //Индексы
        int templateIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Template");
        int partTerminologyIndex = getColNameIndex(sheetPreimptParams.getRow(0),"PartTerminologyName");
        int subtypeIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Subtype");
        int purposeIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Purpose");
        int shortDescrIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Short_Descr");
        int detailDescrIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Detailed_Descr");

        //Индексы листа для записи
        int iMatch = pDescription + 2;
        int iSubtype = pDescription + 3;
        int iSeries = pDescription + 4;
        int iProductName = pDescription + 5;
        int iPartTerminologyName = pDescription + 6;
        int iTemplate = pDescription + 7;
        int iPurpose = pDescription + 8;
        int iShortDescr = pDescription + 9;
        int iDetailedDescr = pDescription + 10;

        int returnValue = 0;
        String breakFromMatch100 = "";

        for(int x = 0; x < arrTemplateGenerator.size(); x++){
            String templtGeneratorVal = arrTemplateGenerator.get(x);
            for(int g = 0; g < currentBrandinPreimptParamRowArr.size(); g++){
                //rowNumber номер строки в которой есть текущий бренд на листе с преипорт параметрами
                int rowNumber = currentBrandinPreimptParamRowArr.get(g);
                String searchTemplate = getCellValue(sheetPreimptParams, rowNumber, templateIndex);
                if (templtGeneratorVal.equalsIgnoreCase(searchTemplate)) {
                    String PartTerminology = getCellValue(sheetPreimptParams, rowNumber, partTerminologyIndex);
                    String Subtype = getCellValue(sheetPreimptParams, rowNumber, subtypeIndex);
                    String Purpose = getCellValue(sheetPreimptParams, rowNumber, purposeIndex);
                    String shortDescr = getCellValue(sheetPreimptParams, rowNumber, shortDescrIndex);
                    String detailDescr = getCellValue(sheetPreimptParams, rowNumber, detailDescrIndex);

                    //Запись
                    sheetPreimport.getRow(i).createCell(iMatch).setCellValue("100%");
                    sheetPreimport.getRow(i).createCell(iSubtype).setCellValue(Subtype);
                    sheetPreimport.getRow(i).createCell(iSeries).setCellValue(Series);
                    sheetPreimport.getRow(i).createCell(iProductName).setCellValue(ProdName);
                    sheetPreimport.getRow(i).createCell(iPartTerminologyName).setCellValue(PartTerminology);
                    sheetPreimport.getRow(i).createCell(iTemplate).setCellValue(searchTemplate);
                    sheetPreimport.getRow(i).createCell(iPurpose).setCellValue(Purpose);
                    sheetPreimport.getRow(i).createCell(iShortDescr).setCellValue(shortDescr);
                    sheetPreimport.getRow(i).createCell(iDetailedDescr).setCellValue(detailDescr);

                    returnValue = 100;

                    breakFromMatch100 = "break";
                    break;
                }
            }
            if (breakFromMatch100.equalsIgnoreCase("break")) {

                break;
            }
        }

        return returnValue;
    }
    //Проверка на совпадение 70%
    public static Integer Match70Func(ArrayList<String> arrTemplateGeneratorReg,
                                      Sheet sheetPreimptParams,
                                      Sheet sheetPreimport,
                                      int pDescription,
                                      int i,
                                      String ProdName,
                                      String Series,
                                      ArrayList<Integer> currentBrandinPreimptParamRowArr){
        //Индексы
        int templateIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Template");
        int partTerminologyIndex = getColNameIndex(sheetPreimptParams.getRow(0),"PartTerminologyName");
        int subtypeIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Subtype");
        int purposeIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Purpose");
        int shortDescrIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Short_Descr");
        int detailDescrIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Detailed_Descr");

        //Индексы листа для записи
        int iMatch = pDescription + 2;
        int iSubtype = pDescription + 3;
        int iSeries = pDescription + 4;
        int iProductName = pDescription + 5;
        int iPartTerminologyName = pDescription + 6;
        int iTemplate = pDescription + 7;
        int iPurpose = pDescription + 8;
        int iShortDescr = pDescription + 9;
        int iDetailedDescr = pDescription + 10;

        int returnValue = 0;
        String breakFromMatch70 = "";

        for(int x = 0; x < arrTemplateGeneratorReg.size(); x++){
            String templtGeneratorVal = arrTemplateGeneratorReg.get(x);

            for(int g = 0; g < currentBrandinPreimptParamRowArr.size(); g++){
                //rowNumber номер строки в которой есть текущий бренд на листе с преипорт параметрами
                int rowNumber = currentBrandinPreimptParamRowArr.get(g);
                String searchTemplate = getCellValue(sheetPreimptParams, rowNumber, templateIndex);
                if (searchTemplate.toLowerCase().matches(templtGeneratorVal)) {
                    String PartTerminology = getCellValue(sheetPreimptParams, rowNumber, partTerminologyIndex);
                    String Subtype = getCellValue(sheetPreimptParams, rowNumber, subtypeIndex);
                    String Purpose = getCellValue(sheetPreimptParams, rowNumber, purposeIndex);
                    String shortDescr = getCellValue(sheetPreimptParams, rowNumber, shortDescrIndex);
                    String detailDescr = getCellValue(sheetPreimptParams, rowNumber, detailDescrIndex);

                    //Запись
                    sheetPreimport.getRow(i).createCell(iMatch).setCellValue("70%");
                    sheetPreimport.getRow(i).createCell(iSubtype).setCellValue(Subtype);
                    sheetPreimport.getRow(i).createCell(iSeries).setCellValue(Series);
                    sheetPreimport.getRow(i).createCell(iProductName).setCellValue(ProdName);
                    sheetPreimport.getRow(i).createCell(iPartTerminologyName).setCellValue(PartTerminology);
                    sheetPreimport.getRow(i).createCell(iTemplate).setCellValue(searchTemplate);
                    sheetPreimport.getRow(i).createCell(iPurpose).setCellValue(Purpose);
                    sheetPreimport.getRow(i).createCell(iShortDescr).setCellValue(shortDescr);
                    sheetPreimport.getRow(i).createCell(iDetailedDescr).setCellValue(detailDescr);

                    returnValue = 70;

                    breakFromMatch70 = "break";
                    break;
                }
            }
            if (breakFromMatch70.equalsIgnoreCase("break")) {
                break;
            }
        }

        return returnValue;
    }

    //Проверка на совпадение 50%
    public static Integer Match50Func(ArrayList<Integer> currentBrandinPreimptParamRowArr,
                                      Sheet sheetPreimptParams,
                                      Sheet sheetPreimport,
                                      int pDescription,
                                      int i,
                                      String ProdName,
                                      String Series){

        //Создание продукт нейм с окончанием S и без S
        int ProdNameLength = ProdName.length();
        String ProdNameLastS = ProdName.substring(ProdNameLength-1);
        String ProdNameL = ProdName.toLowerCase();
        if(ProdNameLastS.equalsIgnoreCase("s")){
            ProdNameL = ProdNameL.substring(0, (ProdNameLength-1));
        }
        //Создание продукт нейм без Set и Kit
        if(ProdNameL.toLowerCase().contains(" set")){
            ProdNameL = ProdNameL.toLowerCase().replace(" set","");
        }else if(ProdNameL.toLowerCase().contains(" kit")){
            ProdNameL = ProdNameL.toLowerCase().replace(" kit","");
        }

        //Индексы
        int partTerminologyIndex = getColNameIndex(sheetPreimptParams.getRow(0),"PartTerminologyName");
        int subtypeIndex = getColNameIndex(sheetPreimptParams.getRow(0),"Subtype");

        //Индексы листа для записи
        int iMatch = pDescription + 2;
        int iSubtype = pDescription + 3;
        int iSeries = pDescription + 4;
        int iProductName = pDescription + 5;
        int iPartTerminologyName = pDescription + 6;

        int returnValue = 0;
        String breakFromMatch50 = "";

        String generateParterminolog = "(.*)" + ProdNameL + "(.*)";

        for(int g = 0; g < currentBrandinPreimptParamRowArr.size(); g++){
            //rowNumber номер строки в которой есть текущий бренд на листе с преипорт параметрами
            int rowNumber = currentBrandinPreimptParamRowArr.get(g);
            String searchParterminolg = getCellValue(sheetPreimptParams, rowNumber, partTerminologyIndex);
            if (searchParterminolg.toLowerCase().matches(generateParterminolog)){
                String Subtype = getCellValue(sheetPreimptParams, rowNumber, subtypeIndex);

                //Запись
                sheetPreimport.getRow(i).createCell(iMatch).setCellValue("50%");
                sheetPreimport.getRow(i).createCell(iSubtype).setCellValue(Subtype);
                sheetPreimport.getRow(i).createCell(iSeries).setCellValue(Series);
                sheetPreimport.getRow(i).createCell(iProductName).setCellValue(ProdName);
                sheetPreimport.getRow(i).createCell(iPartTerminologyName).setCellValue(searchParterminolg);

                returnValue = 50;

                break;

            }
        }

        return returnValue;
    }


}

