package org.example.newpoi;
import javafx.scene.control.TextArea;

import static java.lang.Thread.sleep;

public class TerminatorSaing  {

    public static void Speech(TextArea terminatorSay){
        String[] ArrSpeach = new String[12];
        ArrSpeach[0] = "Where is Джон Коннор?";
        ArrSpeach[1] = "Когда уже ЗП зайдет!?";
        ArrSpeach[2] = "Как говорил мой дед: я твой дед";
        ArrSpeach[3] = "Что, работы много? Иди чаю попей, я все сделаю!";
        ArrSpeach[4] = "Ну и погодка...";
        ArrSpeach[5] = "Скоро там уже 18:00?";
    }
    public static String SpeechButtonClick(){

        String[] ArrSpeach = new String[11];
        ArrSpeach[0] = "Relax, take it easy.. ";
        ArrSpeach[1] = "Don't worry, я все сделаю!";
        ArrSpeach[2] = "Скоро там уже 18:00?";
        ArrSpeach[3] = "Когда уже ЗП зайдет!?";
        ArrSpeach[4] = "Where is Джон Коннор?";
        ArrSpeach[5] = "Что, работы много? Иди чаю попей, я все сделаю!";
        ArrSpeach[6] = "Ну и погодка...";
        ArrSpeach[7] = "Как говорил мой дед: я твой дед";
        ArrSpeach[8] = "Как говориться, делу время, а потехе час. ";
        ArrSpeach[9] = "Don't worry, я все сделаю!";
        ArrSpeach[10] = "Relax, take it easy.. ";

        int i = (int)(Math.random() * 11);
        String speech = ArrSpeach[i];

        return speech;
    }
}
