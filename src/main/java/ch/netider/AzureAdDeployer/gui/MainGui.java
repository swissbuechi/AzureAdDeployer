package ch.netider.AzureAdDeployer.gui;

import ch.netider.AzureAdDeployer.config.AppConfig;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;

import java.io.IOException;

public class MainGui extends Application {

    @Override
    public void start(Stage primaryStage) throws IOException {
        buildMainView(primaryStage);
    }

    public void show() {
        launch();
    }

    private void buildMainView(Stage primaryStage) throws IOException {
        Parent root = FXMLLoader.load(getClass().getResource("/fxml/main_navigation.fxml"));
        primaryStage.setTitle(AppConfig.APP_NAME);
        primaryStage.setMinWidth(400);
        primaryStage.setMinHeight(300);
        primaryStage.getIcons().add(new Image(getClass().getClassLoader().getResource("images/logo.png").toString()));
        Scene scene = new Scene(root, 1600, 800);
        scene.getStylesheets().add("css/stylesheet.css");
        primaryStage.setScene(scene);
        primaryStage.show();
    }
}
