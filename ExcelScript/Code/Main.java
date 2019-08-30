package sample;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import java.io.*;
import java.util.Scanner;

import javax.xml.soap.Text;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{
        primaryStage.setTitle("Excel GUI");

        //Instantiate Labels
        Label filePathLabel = new Label("Please enter the file path the file is located in: ");
        Label fileNameLabel = new Label("Please enter the file name: ");

        //Instantiate Text Fields
        TextField filePathField = new TextField();
        TextField fileNameField = new TextField();

        //Instantiate Button
        Button button = new Button("Run Program");

        //Instantiate VBox for display
        VBox vbox = new VBox();

        //Instantiate Grid Pane for display
        GridPane gridPane = new GridPane();

        //Place objects on GridPane
        gridPane.add(filePathLabel, 0, 0);
        gridPane.add(filePathField, 1, 0);
        gridPane.add(fileNameLabel, 0, 1);
        gridPane.add(fileNameField, 1, 1);

        //Place objects on VBox
        vbox.getChildren().addAll(gridPane, button);


        //Set Alignments on boxes
        vbox.setAlignment(Pos.CENTER);
        gridPane.setAlignment(Pos.CENTER);

        //Set up to be visualized
        Scene scene = new Scene(vbox);
        primaryStage.setScene(scene);

        button.setOnMouseClicked(event -> {
                    try {
                        Process process = Runtime.getRuntime().exec("python main.py");
                        Scanner scanner = new Scanner(process.getInputStream());
                        PrintWriter printWriter = new PrintWriter(process.getOutputStream());
                        printWriter.println(filePathField.getText());
                        printWriter.println(fileNameField.getText()+".xlsx");
                        printWriter.flush();
                        try {
                            Thread.sleep(10);
                        } catch (InterruptedException e) {
                        }
                        String str = scanner.nextLine();
                        if (str.contains("Success!")){
                            VBox successBox = new VBox();
                            Label successLabel = new Label("The Excel Sheet has been colored!");
                            successBox.getChildren().add(successLabel);
                            Scene successScene = new Scene(successBox);
                            successBox.setAlignment(Pos.CENTER);
                            primaryStage.setScene(successScene);
                        }
                        else{
                            VBox failureBox = new VBox();
                            Label failureLabel = new Label("Error");
                            failureBox.getChildren().add(failureLabel);
                            Scene failureScene = new Scene(failureBox);
                            failureBox.setAlignment(Pos.CENTER);
                            primaryStage.setScene(failureScene);
                        }
                    } catch (Exception e) {}
                });

        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}