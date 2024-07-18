package org.example;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Launcher {
    public static void main(String[] args) {
        //Permission mode
        Variant QBPermissionMode = new Variant(1);
        //Mode for Multi user/Single User or both, this setting is both
        Variant QBaccessMode = new Variant(2);
        //Leave Empty to use the currently opened QB file
        //String fileLocation = "";
        String XMLRequest = "";
        String appID = "";//not needed unless you want to set AppID
        String applicationName = "QB Sync Test";
        Dispatch MySessionManager = new Dispatch("QMXMLRP2.RequestProcessor");
        Dispatch.call(MySessionManager, "OpenConnection2", appID, applicationName, QBPermissionMode, QBaccessMode);
        Variant ticket = Dispatch.call(MySessionManager, "BeginSession", fileLocation, QBaccessMode);
        Variant apiResponse = Dispatch.call(MySessionManager, "ProcessRequest", ticket, XMLRequest);
        System.out.println(apiResponse.toString());
        Dispatch.call(MySessionManager, "EndSession", ticket);
        Dispatch.call(MySessionManager, "CloseConnection");

    }
}
