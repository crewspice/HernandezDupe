package org.example;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.jacob.com.ComFailException;

public class Launcher {
    static {
        System.load(System.getProperty("user.dir") + "\\lib\\jacob-1.21-x64.dll");
    }

    public static void main(String[] args) {
        // Permission mode
        Variant QBPermissionMode = new Variant(1);
        // Mode for Multi user/Single User or both, this setting is both
        Variant QBaccessMode = new Variant(2);
        // Leave Empty to use the currently opened QB file
        String fileLocation = "Z:\\max high reach final.qbw";
        String XMLRequest = "<?qbxml version=\"16.0\"?>\n" +
                "<QBXML>\n" +
                "  <QBXMLMsgsRq onError=\"stopOnError\">\n" +
                "    <InvoiceQueryRq requestID=\"4\">\n" +
                "      <TxnDateRangeFilter>\n" +
                "        <FromTxnDate>2024-05-01</FromTxnDate>\n" +
                "        <ToTxnDate>2024-05-15</ToTxnDate>\n" +
                "      </TxnDateRangeFilter>\n" +
                "      <IncludeLineItems>true</IncludeLineItems>\n" +
                "    </InvoiceQueryRq>\n" +
                "  </QBXMLMsgsRq>\n" +
                "</QBXML>\n";

        String appID = ""; // or null, as indicated in the documentation
        String applicationName = "QB Sync Test";
        String connectionPreference = "localQBD"; // or whatever is appropriate

        Dispatch MySessionManager = new Dispatch("QBXMLRP2.RequestProcessor");

        try {
            Variant result = Dispatch.call(MySessionManager, "OpenConnection2", new Variant(appID), new Variant(applicationName), QBPermissionMode);

            // Check if OpenConnection2 was successful
            if (result == null) {
                System.out.println("OpenConnection2 succeeded");
            } else {
                System.out.println("OpenConnection2 failed: " + result.toString());
                return; // Exit or handle failure appropriately
            }

            // Logging for debugging purposes
            System.out.println("BeginSession parameters: fileLocation=" + fileLocation + ", QBaccessMode=" + QBaccessMode.toString());

            Variant ticket = Dispatch.call(MySessionManager, "BeginSession", new Variant(fileLocation), QBaccessMode);
            Variant apiResponse = Dispatch.call(MySessionManager, "ProcessRequest", ticket, new Variant(XMLRequest));
            System.out.println(apiResponse.toString());

            Dispatch.call(MySessionManager, "EndSession", ticket);
            Dispatch.call(MySessionManager, "CloseConnection");
        } catch (ComFailException e) {
            System.out.println("Exception message: " + e.getMessage());
            e.printStackTrace(); // Print the stack trace for debugging
        }
    }
}
