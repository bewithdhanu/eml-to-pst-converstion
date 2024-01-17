package com.example.demo;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class CreatePstFromEml {
    private static final Logger logger = LoggerFactory.getLogger(CreatePstFromEml.class);

    public static void main(String[] args) {
        try {
            // Replace "C:/path/to/your/new.pst" with the desired path for your PST file
            String pstFilePath = "/Users/dhanuk/Downloads/new.pst";

            // Create a new Outlook Application instance
            ActiveXComponent outlook = new ActiveXComponent("Outlook.Application");

            // Create a new MAPI namespace
            Dispatch namespace = Dispatch.call(outlook, "GetNamespace", "MAPI").toDispatch();

            // Create a new PST file
            Dispatch pst = Dispatch.call(namespace, "AddStore", pstFilePath).toDispatch();

            // Iterate over your EML files and add them to the PST file
            String[] emlFiles = {"/Users/dhanuk/Downloads/file1.eml", "/Users/dhanuk/Downloads/file2.eml",
                    "/Users/dhanuk/Downloads/file3.eml", "/Users/dhanuk/Downloads/file4.eml",
                    "/Users/dhanuk/Downloads/file5.eml"};
            for (String emlFile : emlFiles) {
                Dispatch mailItem = Dispatch.call(outlook, "CreateItem", 0).toDispatch();
                Dispatch.call(mailItem, "Import", emlFile, 1024); // 1024 is olRFC822 (EML) format
                Dispatch.call(pst, "AddStore", mailItem);
            }

            System.out.println("PST file created successfully.");

        } catch (Exception e) {
            logger.error("An error occurred: {}", e.getMessage(), e);
        }
    }
}

