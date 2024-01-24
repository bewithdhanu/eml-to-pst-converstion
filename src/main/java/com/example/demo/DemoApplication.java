package com.example.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) {
		public static void main(String[] args) {
        // Set Aspose license (if you have one)
        // License license = new License();
        // license.setLicense("path/to/Aspose.Email.Java.lic");

        // Specify the path to the directory containing EML files
        String emlDirectory = "path/to/eml/files";

        // Specify the output PST file
        String pstFilePath = "path/to/output/file.pst";

        // Get all EML files in the directory
        try {
            Directory.CreateDirectory(emlDirectory);
            Files.walk(Paths.get(emlDirectory))
                    .filter(Files::isRegularFile)
                    .filter(file -> file.toString().toLowerCase().endsWith(".eml"))
                    .forEach(emlFile -> {
                        try {
                            // Load each EML file into the MailMessageCollection
                            MailMessage eml = MailMessage.load(emlFile.toString());
                            mailMessages.add(eml);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    });
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Initialize the MailMessageCollection
        MailMessageCollection mailMessages = new MailMessageCollection();

        // Iterate through EML files and load them into the MailMessageCollection
        for (String emlFile : Arrays.asList("file1.eml", "file2.eml", "file3.eml")) {
            MailMessage eml = MailMessage.load(emlDirectory + "/" + emlFile);
            mailMessages.add(eml);
        }

        // Initialize the PersonalStorage
        PersonalStorage personalStorage = PersonalStorage.create(pstFilePath, FileFormatVersion.Unicode);

        // Create a new folder in the PST file
        FolderInfo inboxFolder = personalStorage.getRootFolder().addSubFolder("Inbox");

        // Add EML messages to the PST folder
        inboxFolder.addMessages(mailMessages);

        // Save changes to the PST file
        personalStorage.dispose();
    }
	}

}
