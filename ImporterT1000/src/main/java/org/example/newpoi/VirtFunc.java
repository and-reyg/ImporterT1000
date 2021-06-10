package org.example.newpoi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import javafx.scene.control.TextArea;

public class VirtFunc {
    public static void main(String[] args) throws IOException {

    }
    //функция для получения места нахождения колонки
    public static int getColNameIndex(Row row, String columnName) {
        int i = 0;
        for (Cell cell : row) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return i;
            }
            i++;
        }
        throw new RuntimeException("ColNameIndex index not found");
    }
    //Запись
    public static void writeWorkbook(Workbook workbook, String fileName){
        try{
            FileOutputStream output = new FileOutputStream(fileName);
            workbook.write(output);
            output.close();
            //workbook.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    //Получение уникальных названий ключей
    public static Set<String> getBrandKeySet(Sheet sheet, int columnIndex) {
        Set<String> brandKeySet = new HashSet<>();

        Iterator<Row> rowIterator = sheet.rowIterator();
        //skip first raw
        if (rowIterator.hasNext()) rowIterator.next();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            brandKeySet.add(row.getCell(columnIndex).getStringCellValue());
        }
        return brandKeySet;
    }
    //формирование индексов для ключей
    public static Map<String, Integer> generateIndexedForPKey(Set<String> brandKeySet) {
        int i = 1;
        Map<String, Integer> indexes = new HashMap<>();
        for (String brandKey : brandKeySet) {
            indexes.put(brandKey, i++);
        }
        return indexes;
    }
    //Запись индексов
    public static void writePKey(Sheet sheet, int brandKeyIndex, int pKeyIndex, Map<String, Integer> indexes) {
        Iterator<Row> rowIterator = sheet.rowIterator();
        //skip first raw
        if (rowIterator.hasNext()) {
            Row firstRow = rowIterator.next();
            //firstRow.createCell(pKeyIndex).setCellValue("pKey");
        }
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            int indexKey = indexes.getOrDefault(row.getCell(brandKeyIndex).getStringCellValue(), -1);
            row.createCell(pKeyIndex).setCellValue("p"+indexKey);
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
    //Перенос колонок brandKey, pKey, specification на лист virtList
    public static void writeVirtList(Sheet sheet1Origin, Sheet virtListOrigin, int brandKeyIndexOrigin, int pKeyIndexOrigin, int specificationIndexOrigin){
        int rowNum = 0;
        for(Row row : sheet1Origin){
            String brandKeyVal = VirtFunc.getCellValue(sheet1Origin, rowNum, brandKeyIndexOrigin);
            virtListOrigin.createRow(rowNum).createCell(0).setCellValue(brandKeyVal);
            String pKeyVal = VirtFunc.getCellValue(sheet1Origin, rowNum, pKeyIndexOrigin);
            virtListOrigin.getRow(rowNum).createCell(1).setCellValue(pKeyVal);
            String specificationVal = VirtFunc.getCellValue(sheet1Origin, rowNum, specificationIndexOrigin);
            virtListOrigin.getRow(rowNum).createCell(2).setCellValue(specificationVal);
            rowNum++;
        }
    }
    //Получение в массив уникальные название параметров со спецификаций
    public static HashSet<String> UniqueColmNamesFromSpecification(Sheet virtList, int lastRow, int virtList_specificationIndex) {
        HashSet<String> colmNames = new HashSet<>();
        for (int i = 1; i < lastRow; i++) {
            //получение значения с ячейки спецификации
            String spec = VirtFunc.getCellValue(virtList, i, virtList_specificationIndex);

            if (spec.toLowerCase().contains(":os:")) {
                while (spec.toLowerCase().contains(":os:")) {
                    //номер вхождения первых двоеточий
                    int index2dot = spec.indexOf(":");
                    //название первой колонки
                    String colmName = spec.substring(0, index2dot);
                    //запись в массив название колонки
                    colmNames.add(colmName);
                    //номер вхождения первой скобы
                    int indexOS = spec.indexOf("[");
                    //Удаление первой [:os:] и все что перед ней
                    spec = spec.substring(indexOS + 6);
                    //осталась строка без первого параметра и так будет обрезатся до тех пор пока есть :os:
                }
                if(spec.toLowerCase().contains(":")) {
                    int index2dot = spec.indexOf(":");
                    String colmName = spec.substring(0, index2dot);
                    colmNames.add(colmName);
                }

            } else if (spec.toLowerCase().contains(": ")) {
                //эта часть отработает если нет :os: но есть двоеточие с пробелом
                int index2dot = spec.indexOf(":");
                String colmName = spec.substring(0, index2dot);
                colmNames.add(colmName);
            }

        }
        return colmNames;
    }
    //сортировка колонок в правильный порядок ПЕРЕНОС В КОНЕЦ
    public static ArrayList<String> sortColmMoveToEnd (ArrayList<String> colmNames, Sheet SpecColmSort){
        //перечень параметров которые содержатся в имени колонки которую необходимо перенести в КОНЕЦ
        ArrayList<String> findInColmNameForMoveToEnd = new ArrayList<>();
        int i = 0;
        for(Row row: SpecColmSort){
            if(i != 0){
                //String parameterName = row.getCell(1).getStringCellValue();
                String parameterName = getCellValue(SpecColmSort, i, 1);
                if(parameterName != ""){
                    findInColmNameForMoveToEnd.add(parameterName);
                }

            }
            i++;
        }
        //System.out.println(findInColmNameForMoveToEnd);

        //Ищет каждое имя в колонке которую нужно перенести
        for (String searchValueForMoveToEnd : findInColmNameForMoveToEnd){
            //в массив moveColms добавится сисок колонок которые необходимо переместить
            ArrayList<String> moveColms = new ArrayList<>();
            Iterator<String> colmNamesIterator = colmNames.iterator();//создаем итератор
            while(colmNamesIterator.hasNext()) {//до тех пор, пока в списке есть элементы
                String colmName = colmNamesIterator.next();//получаем следующий элемент
                if(colmName.toLowerCase().contains(searchValueForMoveToEnd)){
                    moveColms.add(colmName);
                    //удаление колонок с массива colmNames
                    colmNamesIterator.remove();
                }
            }
            //добавление в конец массива colmNames колонки которые отбирали ранее для перемещения
            for (String moveColm : moveColms){
                colmNames.add(moveColm);
            }
        }
        return colmNames;
    }
    //сортировка колонок в правильный порядок ПЕРЕНОС В НАЧАЛО
    public static ArrayList<String> sortColmMoveToStart (ArrayList<String> colmNames, Sheet SpecColmSort){
        //перечень параметров которые содержатся в имени колонки которую необходимо перенести в начало
        ArrayList<String> findInColmNameForMoveToStart = new ArrayList<>();
        int i = 0;
        for(Row row: SpecColmSort){
            if(i != 0){
                //String parameterName = row.getCell(1).getStringCellValue();
                String parameterName = getCellValue(SpecColmSort, i, 0);
                if(parameterName != ""){
                    findInColmNameForMoveToStart.add(parameterName);
                }

            }
            i++;
        }
       // System.out.println(findInColmNameForMoveToStart);
        //Ищет каждое имя в колонке которую нужно перенести
        for (String searchValueForMoveToStart : findInColmNameForMoveToStart){
            //в массив moveColms добавится сисок колонок которые необходимо переместить
            ArrayList<String> moveColms = new ArrayList<>();
            Iterator<String> colmNamesIterator = colmNames.iterator();//создаем итератор
            while(colmNamesIterator.hasNext()) {//до тех пор, пока в списке есть элементы
                String colmName = colmNamesIterator.next();//получаем следующий элемент
                if(colmName.toLowerCase().contains(searchValueForMoveToStart)){
                    moveColms.add(colmName);
                    //удаление колонок с массива colmNames
                    colmNamesIterator.remove();
                }
            }

            //переменая для получение в нее колонки которая должна быть первой
            String mainValue = "";
            Iterator<String> moveColmsIterator = moveColms.iterator();//создаем итератор
            while(moveColmsIterator.hasNext()) {//до тех пор, пока в списке есть элементы
                String colmName = moveColmsIterator.next();//получаем следующий элемент
                if(colmName.equalsIgnoreCase(searchValueForMoveToStart)){
                    mainValue = colmName;
                    //удаление колонок с массива moveColms
                    moveColmsIterator.remove();
                }
            }
            //добавление mainValue в конец moveColms
            moveColms.add(mainValue);

            //добавление в начало массива colmNames колонки которые отбирали ранее для перемещения
            for (String moveColm : moveColms){
                colmNames.add(0, moveColm);
            }
        }
        return colmNames;
    }
    //запись на лист virtList название колонок виртуальных опций
    public static void writeVirtColmName(Sheet virtList, ArrayList<String> colmNames, int virtList_specificationIndex){
        int i = 2;
        for (String colmName : colmNames){
            virtList.getRow(0).createCell(virtList_specificationIndex + i).setCellValue(colmName);
            i++;
        }

    }
    //Запись значений со спецификаций в отсортированные колонки
    public static void AddColmValuesForVirtOption(Sheet virtList, int lastRow, int virtList_specificationIndex, TextArea logMassege){
        int errorMessageIndex = VirtFunc.getColNameIndex(virtList.getRow(0), "ErrorMessage");
        //запись значений виртуальных опций
        for (int i = 1; i < lastRow; i++){
            System.out.println("Обработка "+i + " строки из "+ (lastRow-1));
            logMassege.appendText("Обработка "+i + " строки из "+ (lastRow-1) + "\n");
            //получение значения с ячейки спецификации
            String spec = VirtFunc.getCellValue(virtList, i, virtList_specificationIndex);
            if (spec.toLowerCase().contains(":os:")) {
                while (spec.toLowerCase().contains(":os:")) {
                    //номер вхождения первых двоеточий
                    int index2dot = spec.indexOf(":");
                    //название первой колонки
                    String colmName = spec.substring(0, index2dot);
                    //номер вхождения первой скобы
                    int indexOS = spec.indexOf("[");
                    //значение первой колонки
                    try {
                        String colmValue = spec.substring(index2dot+2, indexOS);
                    }catch(StringIndexOutOfBoundsException e) {
                        virtList.getRow(i).createCell(errorMessageIndex).setCellValue("ERROR отсутствует \": \"  = "+ spec);
                        break;
                    }
                    String colmValue = spec.substring(index2dot+2, indexOS);
                    //поиск индекса colmName (колонка куда записать данные)
                    int colmIndex = VirtFunc.getColNameIndex(virtList.getRow(0), colmName);

                    //запись colmValue
                    virtList.getRow(i).createCell(colmIndex).setCellValue(colmValue);

                    //Удаление первой [:os:] и все что перед ней
                    spec = spec.substring(indexOS + 6);
                    //осталась строка без первого параметра и так будет обрезатся до тех пор пока есть :os:

                }
                if(spec.toLowerCase().contains(": ")) {
                    int index2dot = spec.indexOf(":");
                    String colmName = spec.substring(0, index2dot);
                    String colmValue = spec.substring(index2dot+2);
                    //поиск индекса colmName (колонка куда записать данные)
                    int colmIndex = VirtFunc.getColNameIndex(virtList.getRow(0), colmName);

                    //запись colmValue
                    virtList.getRow(i).createCell(colmIndex).setCellValue(colmValue);
                }else if(spec.length() > 0){
                    virtList.getRow(i).createCell(errorMessageIndex).setCellValue("ERROR отсутствует \": \"  = "+ spec);
                }
            }else if (spec.toLowerCase().contains(": ")) {
                //эта часть отработает если нет :os: но есть двоеточие с пробелом
                int index2dot = spec.indexOf(":");
                String colmName = spec.substring(0, index2dot);
                String colmValue = spec.substring(index2dot+2);
                //поиск индекса colmName (колонка куда записать данные)
                int colmIndex = VirtFunc.getColNameIndex(virtList.getRow(0), colmName);

                //запись colmValue
                virtList.getRow(i).createCell(colmIndex).setCellValue(colmValue);
            }else if(spec.length() > 0){
                virtList.getRow(i).createCell(errorMessageIndex).setCellValue("ERROR отсутствует \": \"  = "+ spec);
            }
        }
    }
    //Двумерный массив для записи индексов виртуальных колонок в первый массив
    public static int[][] FuncArr2D (int virtualStart, int virtualEnd, int virtualQtCol){
        //Запись индексов колонок с виртуалками в ДВУМЕРНЫЙ МАСИВ (arr2D)
        int[][] arr2D = new int[virtualQtCol][2];
        int z = 0;
        for(int i = virtualStart+1; i<virtualEnd; i++ ){
            arr2D[z][0] = i;
            //System.out.println(i);
            z++;
        }
        //System.out.println("Запсиали в масив arr2D");
        return arr2D;
    }
    //создание сет для записи в нее pKey
    public static HashSet<String> PKeySet(Sheet virtList, int lastRow, int virtList_pKeyIndex){
        //создаем сет для записи в нее pKey
        HashSet<String> pKeySet = new HashSet<String>();
        //записываем в сет уникальные ключи с колонки KEY
        for (int rowNum = 1; rowNum < lastRow; rowNum++) {
            Row row = virtList.getRow(rowNum);

            //String cellValue = getCellValue(virtList, rowNum, virtList_pKeyIndex);
            //конвертация типа String в тип int
            //int pKey = Integer.parseInt(cellValue.trim());
            String pKey = getCellValue(virtList, rowNum, virtList_pKeyIndex);
            pKeySet.add(pKey);
        }
        //System.out.println("Уникальные значения колонки Key: "+ pKeySet);
        return pKeySet;
    }
    //получение значений колонок параметров текущей строки текущего перента, подсчет количества и запись в arr2D
    public static int[][] FuncArr2DColmValQt(Sheet virtList, int[][] arr2D, int rowIndex){
        for(int x = 0; x<arr2D.length; x++) {
            //System.out.println("Значение масссива arr2D: " + arr2D[x][0]);
            //создаем cell для проверки на пустоту, и счета заполенных ячеек
            Cell c = virtList.getRow(rowIndex).getCell(arr2D[x][0]);
            //System.out.println("cell = " + c);
            //запись в arr2D / если ячейка заполенная то записывает во второй индекс двумерного масива 1, если нет то 0
            if (c == null || c.getCellType() == CellType.BLANK || VirtFunc.getCellValue(virtList, rowIndex, arr2D[x][0]) == "") {
                arr2D[x][1] += 0;
                //System.out.println(" +0 ");
            }else{
                arr2D[x][1] += 1;
                //System.out.println(" +1 ");
            }
        }
        return arr2D;
    }
    //функция для поиска наиболее заполненых колонок, и сортировка от большего к меньшему
    public static ArrayList<Integer> GetBigerColQtIndex(int[][] arr2D){
        //Получаем уникальные значения со второго ряда двумерного массива (значения количества заполненых ячееек)
        HashSet<Integer> setCelCount = new HashSet<Integer>();
        for(int i = 0; i < arr2D.length; i++){
            setCelCount.add(arr2D[i][1]);
        }
        //получаем  список уникальных значений со второго ряда двумерного массива от большего к меньшему
        List<Integer> listColmCount = new ArrayList<Integer>(setCelCount);
        Collections.sort(listColmCount);
        Collections.reverse(listColmCount);

        //ищем номер колнки в которой самое большое значение, и записіваем первой в массив и т.д.
        ArrayList<Integer> sortColmNumber = new ArrayList<>();
        for (int celCount : listColmCount) {
            //System.out.println(celCount);
            for(int i = 0; i < arr2D.length; i++) {
                //System.out.println(" индекс i = " + i + " / colmNumb = " + arr2D[i][0] + " qt = " + arr2D[i][1]);
                if(arr2D[i][1] == celCount && celCount != 0){
                    //System.out.println("colmNumb = " + arr2D[i][0] + " qt = " + arr2D[i][1] + " == " + celCount);
                    sortColmNumber.add(arr2D[i][0]);
                }
            }
        }
        return sortColmNumber;
    }
    //обработка строк текущего перента и запись в лист виртуальных опций
    public static void WriteVirtOptionCurPKey(int VirtColmQt, Sheet virtList, Sheet sheet1, Sheet sheetIcons, ArrayList<Integer> pRowIndex, ArrayList<Integer> bigerColQtIndex, int virtNameIndex, int virtValueIndex, int icon1Index, int lastRowIcon){
        for (int i_pRowIndex : pRowIndex) {
            int nIndex = 0;
            int vIndex = 0;
            int iIndex = 0;
            int z = 0;
            Row rowVirtList = virtList.getRow(i_pRowIndex);
            Row rowSheet1 = sheet1.getRow(i_pRowIndex);

            for(int i_bColQtIndex : bigerColQtIndex) {
                if(z == VirtColmQt){
                    break;
                }
                String virtColName = virtList.getRow(0).getCell(i_bColQtIndex).getStringCellValue();
                //System.out.println("Вывод названия колонок  " + virtColName);
                Cell cellVirtValue = rowVirtList.getCell(i_bColQtIndex);
                if (cellVirtValue == null || cellVirtValue.getCellType() == CellType.BLANK || VirtFunc.getCellValue(virtList, i_pRowIndex, i_bColQtIndex) == "") {
                    //System.out.println(" row " + i_pRowIndex + "  колонка " + virtColName + " ПУСТАЯ!");
                    //проверка на пустоту, если пусто ничего не делать
                }else{
                    String virtValue = VirtFunc.getCellValue(virtList, i_pRowIndex, i_bColQtIndex);
                    //System.out.println("Вывод виртопции  " + virtValue);
                    rowSheet1.createCell(virtNameIndex+nIndex).setCellValue(virtColName + ":");
                    rowSheet1.createCell(virtValueIndex+vIndex).setCellValue(virtValue);

                    //запись ICON
                    if(virtColName.toLowerCase().contains("finish") || virtColName.toLowerCase().contains("color")) {
                        //System.out.println(" Нашли цвет ");
                        for(int x = 1; x<lastRowIcon; x++) {
                            String colorName = VirtFunc.getCellValue(sheetIcons, x, 0);
                            //System.out.println("colorName = "+ colorName);
                            String icon = VirtFunc.getCellValue(sheetIcons, x, 1);

                            if(virtValue.equalsIgnoreCase(colorName)){
                                rowSheet1.createCell(icon1Index+iIndex).setCellValue(icon);
                                //System.out.println("icon1Index+iIndex = " + (icon1Index+iIndex));
                                break;
                            }
                        }
                    }
                }

                nIndex += 2;
                vIndex += 2;
                iIndex++;
                z++;
            }
        }
    }
    public static void WriteTXTFile(String filename, String text) throws IOException{
        FileWriter fileWriter = new FileWriter(filename);
        //FileWriter fileWriter = new FileWriter(filename, true); // если нужно записать в конец файла
        BufferedWriter bufferedWriter = new BufferedWriter(fileWriter);
        bufferedWriter.write(text);
        bufferedWriter.close();
        fileWriter.close();
    }
    public static Integer IterationQtForWriteFile(List<String> pKeySetList, int DelitNa){
        //Количество интераций для открытия и записи в файл
        int pKeyQt = pKeySetList.size();
        //System.out.println("pKeyQt "+pKeyQt);
        int ostatokDelenia = pKeyQt % DelitNa;
        int iterationQt = 0;
        if(ostatokDelenia > 0){
            iterationQt = (pKeyQt / DelitNa)+1;
        }else{
            iterationQt = (pKeyQt / DelitNa);
        }
        return iterationQt;
    }
    //получение позиции перента c текстового файла "ParentPosInPKeySetList.txt"
    public static Integer ParentPositionInPKeySetList(String Path) throws IOException {
        //получение позиции перента
        String parentPosString = "";
        for (String line : Files.readAllLines(Paths.get(Path))) {
            parentPosString = line;
        }
        int parentPos = Integer.parseInt (parentPosString);
        return parentPos;
    }

    public static String getPath(JFileChooser chooser){
        File file = chooser.getSelectedFile();
        String fileName = chooser.getSelectedFile().getName();
        String fileAbsolutePath = chooser.getSelectedFile().getAbsolutePath();
        String filePath = fileAbsolutePath.replace(fileName, "");
       // System.out.println("fileName = "+ fileName);
       // System.out.println("fileAbsolutePath = "+fileAbsolutePath);
        //System.out.println("filePath ="+filePath);
        return filePath;
    }

    //Очистка всех ячеек виртопций для новой записи
    public static void CleanVirtColm(Sheet sheet){

        int lastRow = sheet.getLastRowNum()+1;
        //Индекс колонок
        int VirtName_1_Index = VirtFunc.getColNameIndex(sheet.getRow(0),"VirtName_1");
        int icon_1_Index = VirtFunc.getColNameIndex(sheet.getRow(0),"icon_1");

        //Поиск индекса последней иконки
        Row row = sheet.getRow(0);
        int a = 0;
        int iconNumber = 1;
        Cell cell = sheet.getRow(0).getCell(icon_1_Index + a);
        String cellValue = row.getCell(icon_1_Index + a).getStringCellValue();
        String FindValue = "icon_" + iconNumber;
        int newIconIndex = icon_1_Index + a;
        int lastIconIndex = 0;

        while(cellValue.equalsIgnoreCase(FindValue)){
            cellValue = row.getCell(icon_1_Index + a).getStringCellValue();
            FindValue = "icon_" + iconNumber;
            newIconIndex = icon_1_Index + a;

            if(!cellValue.equalsIgnoreCase(FindValue)){
               // System.out.println("cellValue != FindValue");
               // System.out.println("Last icon = " + ("icon_"+ (iconNumber - 1)));
                lastIconIndex = newIconIndex - 1;
                //System.out.println("lastIconIndex = " + lastIconIndex);
            }
            a++;
            iconNumber++;
        }
        for(int b = 1; b < lastRow; b++){
            for(int c = VirtName_1_Index; c <= lastIconIndex; c++){
                sheet.getRow(b).createCell(c).setBlank();
            }

        }




    }
}


