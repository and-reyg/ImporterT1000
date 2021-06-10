package org.example.newpoi;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import javafx.scene.control.TextArea;

public class VirtAddSheetVirtList {
    public static void main(String[] args) throws IOException {

    }
    public static void AddSheetVirtList(String path, String fileSavePath, String fileName, TextArea logMassege) throws IOException{
        System.out.println("Загрузка книги...");
        logMassege.appendText("Загрузка книги... \n");

        String newSaveFile = fileSavePath + fileName + "_NEW.xlsx";

        Workbook wbOrigin = WorkbookFactory.create(new FileInputStream(path));
        Sheet sheet1Origin = wbOrigin.getSheetAt(0);
        Sheet virtListOrigin = wbOrigin.createSheet("VirtList");

        int lastRowOrigin = sheet1Origin.getLastRowNum()+1;


        int brandKeyIndexOrigin = VirtFunc.getColNameIndex(sheet1Origin.getRow(0), "BrandKey");
        System.out.println(brandKeyIndexOrigin);
        int pKeyIndexOrigin = VirtFunc.getColNameIndex(sheet1Origin.getRow(0),"pKey");
        int specificationIndexOrigin = VirtFunc.getColNameIndex(sheet1Origin.getRow(0), "For_Specification");
        System.out.println("Простановка уникальных id пернетам...");
        logMassege.appendText("Простановка уникальных id пернетам...\n");
        System.out.println("Создание нового листа...");
        logMassege.appendText("Создание нового листа...\n");
        //Получение уникальных названий ключей, присвоение им индексов и запись в новую книгу
        Set<String> brandKeySet = VirtFunc.getBrandKeySet(sheet1Origin, brandKeyIndexOrigin);
        Map<String, Integer> indexes = VirtFunc.generateIndexedForPKey(brandKeySet);
        VirtFunc.writePKey(sheet1Origin, brandKeyIndexOrigin, pKeyIndexOrigin, indexes);

        //Перенос колонок brandKey, pKey, specification на лист virtList
        VirtFunc.writeVirtList(sheet1Origin, virtListOrigin, brandKeyIndexOrigin, pKeyIndexOrigin, specificationIndexOrigin);
        VirtFunc.CleanVirtColm(sheet1Origin);

        VirtFunc.writeWorkbook(wbOrigin, newSaveFile);
        wbOrigin.close();
        System.out.println("ID проставлены / Данные скопированы \n");
        logMassege.appendText("ID проставлены / Данные скопированы \n");

    }
}

