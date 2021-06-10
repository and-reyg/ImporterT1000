package org.example.ui;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ResourceBundle;
import java.util.prefs.Preferences;

import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import org.example.newpoi.DescriptionAnalysis;
import org.example.newpoi.PreimportCreator;
import org.example.newpoi.TerminatorSaing;
import org.example.newpoi.VirtOptionCreator_Func;

import javax.swing.*;
import javax.swing.plaf.InternalFrameUI;


public class Controller {

    private static final String LAST_USED_FOLDER = "";
    @FXML
    private ResourceBundle resources;

    @FXML
    private TextArea terminatorSay;

    @FXML
    private Button info;

    @FXML
    private TextField tField_VirtOpt_PD;

    @FXML
    private TextField tField_VirtOpt_Infoimage;

    @FXML
    private TextField tField_VirtColmNumber;

    @FXML
    private Button Add_VirtOpt_PD;

    @FXML
    private Button Add_VirtOpt_Infoimage;

    @FXML
    private Button Run_VirtOpt;

    @FXML
    private TextField tField_Preimp_Descript;

    @FXML
    private TextField tField_Preimp_DB;

    @FXML
    private Button Add_Preimp_Descr;

    @FXML
    private Button Add_Preimp_DB;

    @FXML
    private Button Run_Preimport;

    @FXML
    private TextField tField_DescAnals_Descr;

    @FXML
    private TextField tField_DescAnals_DB;

    @FXML
    private Button Add_DescAnals_Descr;

    @FXML
    private Button Add_DescAnals_DB;

    @FXML
    private Button Run_DescriptionAnalis;

    @FXML
    public TextArea logMassege;

    String AbsolutePath_VirtOpt_PD;
    String AbsolutePath_VirtOpt_Infoimage;
    String AbsolutePath_Preimp_Descript;
    String AbsolutePath_Preimp_DB;
    String AbsolutePath_DescAnals_Descr;
    String AbsolutePath_DescAnals_DB;

    @FXML
    void initialize() {
        logMassege.appendText("\n");

        info.setOnAction(e -> {
            if(info != null){
                try {
                    Desktop.getDesktop().browse(new URL("https://drive.google.com/drive/folders/1qrig0Dpp6EPdo2dK8UNu6RCG0XIr5fYB?usp=sharing").toURI());
                } catch (IOException ioException) {
                    ioException.printStackTrace();
                } catch (URISyntaxException uriSyntaxException) {
                    uriSyntaxException.printStackTrace();
                }
                terminatorSay.setText(TerminatorSaing.SpeechButtonClick());
            }
        });

        Add_VirtOpt_PD.setOnAction(e -> {
            if(Add_VirtOpt_PD != null){
                /*System.out.println("нажата кнокпа <<Add_VirtOpt_PD>>");
                logMassege.appendText("нажата кнокпа <<Add_VirtOpt_PD>> \n");*/

                AbsolutePath_VirtOpt_PD = FilePath(tField_VirtOpt_PD);
                logMassege.appendText("Получен файл PD:  " + AbsolutePath_VirtOpt_PD  + "\n");
                terminatorSay.setText(TerminatorSaing.SpeechButtonClick());
            }
        });

        Add_VirtOpt_Infoimage.setOnAction(e -> {
            if(Add_VirtOpt_Infoimage != null){
                /*System.out.println("нажата кнокпа <<Add_VirtOpt_Infoimage>>");
                logMassege.appendText("нажата кнокпа <<Add_VirtOpt_Infoimage>> \n");*/

                AbsolutePath_VirtOpt_Infoimage = FilePath(tField_VirtOpt_Infoimage);
                logMassege.appendText("Получен файл Infoimage:  " + AbsolutePath_VirtOpt_Infoimage  + "\n");
                terminatorSay.setText(TerminatorSaing.SpeechButtonClick());
            }
        });
       Run_VirtOpt.setOnAction(e -> {
           if(Run_VirtOpt != null){
               /*System.out.println("нажата кнокпа <<Run_VirtOpt>> \n");
               logMassege.setText("нажата кнокпа <<Run_VirtOpt>> \n");*/

               String text_tField_VirtOpt_PD = tField_VirtOpt_PD.getText();
               String text_tField_VirtOpt_Infoimage = tField_VirtOpt_Infoimage.getText();
               String text_tField_VirtColmNumber = tField_VirtColmNumber.getText();

               logMassege.setText("Получен файл PD:  " + AbsolutePath_VirtOpt_PD  + "\n");
               logMassege.appendText("Получен файл Infoimage:  " + AbsolutePath_VirtOpt_Infoimage  + "\n");
               logMassege.appendText("Будет создано колонок вииртопций:  " + text_tField_VirtColmNumber + " \n");
               /*//абсолютный путь проверка
               logMassege.appendText("AbsolutePath_VirtOpt_PD = " + AbsolutePath_VirtOpt_PD  + "\n");
               logMassege.appendText("AbsolutePath_VirtOpt_Infoimage = " + AbsolutePath_VirtOpt_Infoimage  + "\n");*/

               logMassege.appendText("Запускаю обработку Виртуальных опций \n");


               try {
                   VirtOptionCreator_Func.VirtOptionCreator(AbsolutePath_VirtOpt_PD,
                                                            AbsolutePath_VirtOpt_Infoimage,
                                                            text_tField_VirtOpt_PD,
                                                            text_tField_VirtColmNumber,
                                                            logMassege);
               } catch (IOException ioException) {
                   String a = ioException.toString();
                   logMassege.appendText("\n\n   ERROR:  " + a);
                   if(a.equalsIgnoreCase("java.io.IOException: GC overhead limit exceeded") ||
                           a.equalsIgnoreCase("java.io.IOException: Java heap space")){
                       logMassege.appendText("\n    Слишком большой файл!");
                   }
                   ioException.printStackTrace();
               } catch (NullPointerException e1){
                   logMassege.appendText("\n   ERROR: Получено нулевое значение! \n " +
                           "Возможные варианты решения: \n" +
                           "   * Удалить пустые строки; \n" +
                           "   * Проверить основные колонки на пустые значения \n");
                   e1.printStackTrace();
               }

               terminatorSay.setText(TerminatorSaing.SpeechButtonClick());
           }
       });

       //---------------  Второй Модуль, Поиск Новых СКУ -----------------
        Add_Preimp_Descr.setOnAction(e -> {
            if(Add_Preimp_Descr != null){
                /*System.out.println("нажата кнокпа <<Add_Preimp_Descr>>");
                logMassege.appendText("нажата кнокпа <<Add_Preimp_Descr>> \n");*/

                AbsolutePath_Preimp_Descript = FilePath(tField_Preimp_Descript);
                logMassege.appendText("Получен файл Preimport Description:  " + AbsolutePath_Preimp_Descript  + "\n");
                terminatorSay.setText(TerminatorSaing.SpeechButtonClick());
            }
        });
        Add_Preimp_DB.setOnAction(e -> {
            if(Add_Preimp_DB != null){
                /*System.out.println("нажата кнокпа <<Add_Preimp_DB>>");
                logMassege.appendText("нажата кнокпа <<Add_Preimp_DB>> \n");*/

                AbsolutePath_Preimp_DB = FilePath(tField_Preimp_DB);
                logMassege.appendText("Получен файл Preimport_DataBase:  " + AbsolutePath_Preimp_DB  + "\n");
                terminatorSay.setText(TerminatorSaing.SpeechButtonClick());
            }
        });
        Run_Preimport.setOnAction(e -> {
            if(Run_Preimport != null){
                /*System.out.println("нажата кнокпа <<Run_Preimport>>");
                logMassege.appendText("нажата кнокпа <<Run_Preimport>> \n");*/

                String text_tField_Preimp_Descript = tField_Preimp_Descript.getText();
                String text_tField_Preimp_DB = tField_Preimp_DB.getText();

                logMassege.setText("Получен файл Preimport Description:  " + AbsolutePath_Preimp_Descript  + "\n");
                logMassege.appendText("Получен файл Preimport_DataBase:  " + AbsolutePath_Preimp_DB  + "\n");
                /*//абсолютный путь проверка
                logMassege.appendText("AbsolutePath_Preimp_Descript = " + AbsolutePath_Preimp_Descript  + "\n");
                logMassege.appendText("AbsolutePath_Preimp_DB = " + AbsolutePath_Preimp_DB  + "\n");*/

                logMassege.appendText("Запускаю обработку Преимпорта \n");

                try {
                    PreimportCreator.PreimportCreator_Func(AbsolutePath_Preimp_Descript,
                                                           AbsolutePath_Preimp_DB,
                                                           text_tField_Preimp_Descript,
                                                           logMassege);
                } catch (IOException ioException) {
                    String a = ioException.toString();
                    logMassege.appendText("\n\n   ERROR:  " + a);
                    if(a.equalsIgnoreCase("java.io.IOException: GC overhead limit exceeded") ||
                            a.equalsIgnoreCase("java.io.IOException: Java heap space")){
                        logMassege.appendText("\n    Слишком большой файл!");
                    }
                    ioException.printStackTrace();
                } catch (NullPointerException e1){
                    logMassege.appendText("\n   ERROR: Получено нулевое значение! \n " +
                            "Возможные варианты решения: \n" +
                            "   * Удалить пустые строки; \n" +
                            "   * Проверить основные колонки на пустые значения \n");
                    e1.printStackTrace();
                }
                terminatorSay.setText(TerminatorSaing.SpeechButtonClick());

            }
        });

        //---------------  Третий Модуль, Разбор Дерьмового Description -----------------
        Add_DescAnals_Descr.setOnAction(e -> {
            if(Add_DescAnals_Descr != null){
                /*System.out.println("нажата кнокпа <<Add_DescAnals_Descr>>");
                logMassege.appendText("нажата кнокпа <<Add_DescAnals_Descr>> \n");*/

                AbsolutePath_DescAnals_Descr = FilePath(tField_DescAnals_Descr);
                logMassege.appendText("Получен файл Product Description:  " + AbsolutePath_DescAnals_Descr  + "\n");
                terminatorSay.setText(TerminatorSaing.SpeechButtonClick());
            }
        });
        Add_DescAnals_DB.setOnAction(e -> {
            if(Add_DescAnals_DB != null){
                /*System.out.println("нажата кнокпа <<Add_DescAnals_DB>>");
                logMassege.appendText("нажата кнокпа <<Add_DescAnals_DB>> \n");*/

                AbsolutePath_DescAnals_DB = FilePath(tField_DescAnals_DB);
                logMassege.appendText("Получен файл DescrAnalysis_DB:  " + AbsolutePath_DescAnals_Descr  + "\n");
                terminatorSay.setText(TerminatorSaing.SpeechButtonClick());
            }
        });
        Run_DescriptionAnalis.setOnAction(e -> {
           if(Run_DescriptionAnalis != null){
               /*System.out.println("нажата кнокпа <<Run_DescriptionAnalis>>");
               logMassege.setText("нажата кнокпа <<Run_DescriptionAnalis>> \n");*/

               String text_tField_DescAnals_Descr = tField_DescAnals_Descr.getText();
               String text_tField_DescAnals_DB = tField_DescAnals_DB.getText();

               logMassege.setText("Получен файл Product Description:  " + AbsolutePath_DescAnals_Descr  + "\n");
               logMassege.appendText("Получен файл DescrAnalysis_DB:  " + AbsolutePath_DescAnals_Descr  + "\n");
               /*//абсолютный путь проверка
               logMassege.appendText("AbsolutePath_DescAnals_Descr = " + AbsolutePath_DescAnals_Descr  + "\n");
               logMassege.appendText("AbsolutePath_DescAnals_DB = " + AbsolutePath_DescAnals_DB  + "\n");*/

               logMassege.appendText("Запускаю обработку Description \n");

               try {
                   DescriptionAnalysis.DescriptAnalis_Func(AbsolutePath_DescAnals_Descr,
                                                           AbsolutePath_DescAnals_DB,
                                                           text_tField_DescAnals_Descr,
                                                           logMassege);
               } catch (IOException ioException) {
                   String a = ioException.toString();
                   logMassege.appendText("\n\n   ERROR:  " + a);
                   if(a.equalsIgnoreCase("java.io.IOException: GC overhead limit exceeded") ||
                           a.equalsIgnoreCase("java.io.IOException: Java heap space")){
                       logMassege.appendText("\n    Слишком большой файл!");
                   }
                   ioException.printStackTrace();
               } catch (NullPointerException e1){
                   logMassege.appendText("\n   ERROR: Получено нулевое значение! \n " +
                           "Возможные варианты решения: \n" +
                           "   * Удалить пустые строки; \n" +
                           "   * Проверить основные колонки на пустые значения \n");
                   e1.printStackTrace();
               }
               terminatorSay.setText(TerminatorSaing.SpeechButtonClick());

           }
       });
    }
    private static boolean isNumber(String s) throws NumberFormatException {
        try {
            Integer.parseInt(s);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
    private static String FilePath(TextField textField){
        //Для открытия в окна выбора файла с последней папке
        Preferences prefs = Preferences.userRoot().node(Controller.class.getName());

        JFileChooser chooser = new JFileChooser(prefs.get(LAST_USED_FOLDER,
                new File(".").getAbsolutePath()));
        int returnVal = chooser.showDialog(null, "Открыть файл");

        String fileAbsolutePath = "";

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            prefs.put(LAST_USED_FOLDER, chooser.getSelectedFile().getParent());
            String fileName = chooser.getSelectedFile().getName();
            //Имя файла отобразиться в поле Интерфейса
            textField.setText(fileName);

            fileAbsolutePath = chooser.getSelectedFile().getAbsolutePath();

        }
        return fileAbsolutePath;

    }
    public static void openWebpage(String url) {
        try {
            new ProcessBuilder("x-www-browser", url).start();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private static void ShowException(){

    }
}

