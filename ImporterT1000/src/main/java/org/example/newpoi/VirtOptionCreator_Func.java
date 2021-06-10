package org.example.newpoi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import javafx.scene.control.TextArea;

public class VirtOptionCreator_Func {

    public static void VirtOptionCreator(String AbsolutePath_VirtOpt_PD,
                                          String AbsolutePath_VirtOpt_Infoimage,
                                          String text_tField_VirtOpt_PD,
                                          String text_tField_VirtColmNumber,
                                          TextArea logMassege) throws IOException {

        //Получение уникальных названий ключей, присвоение им индексов и запись в новую книгу
        //Перенос колонок brandKey, pKey, specification на лист virtList
        String fileNameXlsx = text_tField_VirtOpt_PD;
        String fileSavePath = AbsolutePath_VirtOpt_PD.replace(fileNameXlsx, "");
        String fileName = fileNameXlsx.replace(".xlsx", "");
        String newSaveFile = fileSavePath + fileName + "_NEW.xlsx";

        VirtAddSheetVirtList.AddSheetVirtList(AbsolutePath_VirtOpt_PD, fileSavePath, fileName, logMassege);
        //System.out.println("Разбор спецификаций...");
        logMassege.appendText("Разбор спецификаций...\n");



        //-----------------------      РАБОТА С НОВОЙ КНИГОЙ    ------------------

        //Переменная для установки количества колонок виртуальнх опций
        int VirtColmQt = Integer.parseInt(text_tField_VirtColmNumber);
        //int VirtColmQt = text_tField_VirtColmNumber;

        //открытие книги
        String NewWorkFile = fileSavePath + fileName + "_NEW.xlsx";

        Workbook wb = WorkbookFactory.create(new FileInputStream(NewWorkFile));

        //Открытие файла InfoImage
        Workbook wbInfoimage = WorkbookFactory.create(new FileInputStream(AbsolutePath_VirtOpt_Infoimage));

        Sheet sheet1 = wb.getSheetAt(0);
        Sheet virtList = wb.getSheet("virtList");
        Sheet SpecColmSort = wb.getSheet("SpecColmSort");
        Sheet sheetIcons = wbInfoimage.getSheetAt(0);

        int lastRow = sheet1.getLastRowNum()+1;
        int lastRowIcon = sheetIcons.getLastRowNum()+1;

        //Номера колонок
        int Sheet1_pKeyIndex = VirtFunc.getColNameIndex(sheet1.getRow(0),"pKey");
        int virtList_pKeyIndex = VirtFunc.getColNameIndex(virtList.getRow(0),"pKey");
        int virtList_specificationIndex = VirtFunc.getColNameIndex(virtList.getRow(0), "For_Specification");


        //Получение в массив уникальные название параметров со спецификаций
        HashSet<String> colmNamesSet = VirtFunc.UniqueColmNamesFromSpecification(virtList, lastRow, virtList_specificationIndex);
        ArrayList<String> colmNames = new ArrayList<>(colmNamesSet);
        //System.out.println(colmNames);

        //сортировка колонок в правильный порядок
        //перенос необходимых колонок в Конец
        colmNames = VirtFunc.sortColmMoveToEnd(colmNames, SpecColmSort);
        //перенос необходимых колонок в Начало
        colmNames = VirtFunc.sortColmMoveToStart(colmNames, SpecColmSort);
        //System.out.println(colmNames);
        //добавление колонок для поиска вирутальных опций
        colmNames.add(0,"virtualStart");
        colmNames.add("virtualEnd");
        //System.out.println(colmNames);

        //создание колонки ErrorMessage для записи в нее найденых ошибок при а=парсе спецификаций
        virtList.getRow(0).createCell(virtList_specificationIndex + 1).setCellValue("ErrorMessage");
        //запись на лист virtList название колонок виртуальных опций
        VirtFunc.writeVirtColmName(virtList, colmNames, virtList_specificationIndex);
        VirtFunc.writeWorkbook(wb, NewWorkFile);

        //Запись значений со спецификаций в отсортированные колонки
        VirtFunc.AddColmValuesForVirtOption(virtList, lastRow, virtList_specificationIndex, logMassege);
        VirtFunc.writeWorkbook(wb, NewWorkFile);

        //Поиск диапазона колонок с виртуальными опциями
        //номер колонки virtualStart и virtualEnd
        int virtualStart = VirtFunc.getColNameIndex(virtList.getRow(0), "virtualStart");
        int virtualEnd = VirtFunc.getColNameIndex(virtList.getRow(0), "virtualEnd");

        //индексы колонок для записи виртуалок
        int virtNameIndex = VirtFunc.getColNameIndex(sheet1.getRow(0),"VirtName_1");
        int virtValueIndex = VirtFunc.getColNameIndex(sheet1.getRow(0),"VirtValue_1");
        int icon1Index = VirtFunc.getColNameIndex(sheet1.getRow(0),"icon_1");

        //поиск количества колонок с вируталками
        int virtualQtCol = virtualEnd - virtualStart - 1;
        System.out.println("Количество колонок с вируталками : " + virtualQtCol);

        //Записываем индексы колонок с виртуалками в первый массив ДВУМЕРНого МАСИВА (arr2D)
        int[][] arr2D = VirtFunc.FuncArr2D(virtualStart, virtualEnd, virtualQtCol);
        //System.out.println("arr2D = " + Arrays.deepToString(arr2D));

        //Cоздание масиваСет для записи в него pKey и Перевод Сет в лист
        List<String> pKeySetList = new ArrayList<String>(VirtFunc.PKeySet(virtList, lastRow, virtList_pKeyIndex));
        //System.out.println("pKeySetList : " + pKeySetList);
        System.out.println("Количество перентов : "+pKeySetList.size());
        logMassege.appendText("Количество перентов : "+pKeySetList.size() + "\n");


/*
          Проход циклом по pKeySetList (обработка каждого ключа перента,
              к примеру первый ключ p1, ему соответсвтуют 3 строки
                  - получение номера этих трех строк (Row) / запись их в масив pRowIndex
                  - получение количество заполненых ячеек в колонке / запись их во второй индекс двумерного масива arr2D
              номера строк ключа p1 находятся в масиве pRowIndex, дальше необходимо записать в эти строки вирт параметры
                  - проход циклом pRowIndex (обработка 3х строк перента) и запись в них вирт параметров
              И так с каждым ключем)
        */

        System.out.println("Запись виртуалок...");
        logMassege.appendText("Запись виртуалок..." + "\n");

        //масив для записи номеров строк одного перента в одной итерации pKeySetList
        ArrayList<Integer> pRowIndex = new ArrayList<>();
        int pCount = 0;
        for(int i = 0; i<pKeySetList.size(); i++) {
            pCount++;
            System.out.println("Обработка перента : p" + (i+1)+ " из "+pKeySetList.size());
            logMassege.appendText("Обработка перента : p" + (i+1)+ " из "+pKeySetList.size() + "\n");
            for (int rowNum = 1; rowNum < lastRow; rowNum++) {
                Row row = virtList.getRow(rowNum);
                String colmPKeyValue = VirtFunc.getCellValue(virtList, rowNum, virtList_pKeyIndex);
                int rowIndex = row.getCell(virtList_pKeyIndex).getRowIndex();
                //System.out.println("rowIndex = " + rowIndex + " Row " + (rowIndex + 1) + " cellVal  = " + colmPKeyValue);

                if (colmPKeyValue == pKeySetList.get(i)) {
                    //System.out.println(" >> Key " + pKeySetList.get(i) + " FIND in " + "rowIndex " + rowIndex);
                    //Запись адреса текущей строки текущего ключа перента
                    pRowIndex.add(rowIndex);
                    //получение значений колонок параметров текущей строки текущего перента, подсчет количества и запись в arr2D
                    arr2D = VirtFunc.FuncArr2DColmValQt(virtList, arr2D, rowIndex);
                }
            }
            //System.out.println("Масив с адресами строк перента " + pRowIndex);
            //System.out.println("arr2D = " + Arrays.deepToString(arr2D));
            //System.out.println("Проверка функции getBigerColQtIndex = \n = " + VirtFunc.GetBigerColQtIndex(arr2D));
            ArrayList<Integer> bigerColQtIndex = VirtFunc.GetBigerColQtIndex(arr2D);
            //System.out.println("bigerColQtIndex = " + bigerColQtIndex);

            //проход циклом pRowIndex (обработка строк текущего перента) и запись в лист виртуальных опций
            VirtFunc.WriteVirtOptionCurPKey(VirtColmQt, virtList, sheet1, sheetIcons, pRowIndex, bigerColQtIndex, virtNameIndex, virtValueIndex, icon1Index, lastRowIcon);

            //очистить массивы и Araylist
            for(int c = 0; c < arr2D.length; c++){
                arr2D[c][1] = 0;
            }
            pRowIndex.clear();
        }
        VirtFunc.writeWorkbook(wb, NewWorkFile);
        System.out.println("     ---  Готово!  ---   ");
        logMassege.appendText("     ---  Done!  ---   \n");
        logMassege.appendText("           ---  Hasta la vista, baby! \n\n");



    }

}