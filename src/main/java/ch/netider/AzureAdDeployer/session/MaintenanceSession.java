package ch.netider.AzureAdDeployer.session;

import com.github.tuupertunut.powershelllibjava.PowerShellExecutionException;

import java.io.IOException;

public class MaintenanceSession extends PoweShellSession {

    public MaintenanceSession() {
        super("MaintenanceSession");
    }

    public String execute(String... input) {
        super.open();
        try {
            setInput(input);
            setRawOutput(super.powerShell.executeCommands(input));
            setOutput(getRawOutput());
            if (getOutput() != null) {
                return getOutput();
            }
        } catch (PowerShellExecutionException | IOException | NullPointerException ex) {
            ex.printStackTrace();
            setError(getRawOutput());
        }
        return null;
    }

}
