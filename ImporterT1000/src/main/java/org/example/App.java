package org.example;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

/**
 * ImporterT1000
 */
public class App extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{
        ClassLoader classLoader = getClass().getClassLoader();

        Parent root = FXMLLoader.load(classLoader.getResource("ui/Gui_Importer.fxml"));
        primaryStage.setTitle("Importer t1000");
        primaryStage.setScene(new Scene(root, 670.0, 705.0));
        primaryStage.setResizable(false);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
