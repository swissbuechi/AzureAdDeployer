package ch.netider.AzureAdDeployer.gui.controller;

import ch.netider.AzureAdDeployer.config.AppConfig;
import javafx.fxml.FXML;
import javafx.scene.control.Label;


public class WelcomeController {

    @FXML
    private Label appName;

    @FXML
    private Label author;

    @FXML
    public void initialize() {
        appName.setText(AppConfig.APP_NAME + " " + AppConfig.VERSION);
        author.setText("Developed by " + AppConfig.AUTHOR);
    }

}