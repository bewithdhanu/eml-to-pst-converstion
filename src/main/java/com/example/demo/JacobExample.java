package com.example.demo;

import java.io.File;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class JacobExample {
  public static void main(String[] args) {
    ActiveXComponent outlook = new ActiveXComponent("Outlook.Application");
    Variant stores = outlook.getProperty("Session").getProperty("Stores");
    Variant store = Dispatch.call(stores.getDispatch(), "AddStore", "C:\\archive.pst", 3);
    Variant folders = store.getDispatch().getProperty("Folders");
    Variant folder = Dispatch.call(folders.getDispatch(), "Add", "Inbox");
    Variant items = folder.getDispatch().getProperty("Items");

    File emlFolder = new File("C:\\eml_files");
    File[] emlFiles = emlFolder.listFiles();

    for (File emlFile : emlFiles) {
      Variant item = Dispatch.call(items.getDispatch(), "Add", 0);
      Dispatch.put(item.getDispatch(), "Subject", emlFile.getName());
      Dispatch.put(item.getDispatch(), "Body", "This is a test email.");
      Dispatch.put(item.getDispatch(), "Save", true);
    }
  }
}
