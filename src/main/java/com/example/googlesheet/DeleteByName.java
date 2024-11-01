package com.example.googlesheet;

import java.io.IOException;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;

public class DeleteByName {
    private static final Drive driveService = GoogleConfig.getDriveService();

    public static void deleteSpreadsheetByName(String name) throws IOException {
        String query = "name='" + name + "' and mimeType='application/vnd.google-apps.spreadsheet'";
        FileList result = driveService.files().list().setQ(query).execute();
        if (result.getFiles().isEmpty()) {
            System.out.println("No spreadsheet found with name: " + name);
        } else {
            for (File res: result.getFiles()) {
                String fileId = res.getId();
                driveService.files().delete(fileId).execute();
            }
            // Delete the file by ID
            System.out.println("Spreadsheets with name " + name + " deleted.");
        }
    }

    public static void main(String[] args) throws IOException {
        String spreadsheetName = "Hello World";
        deleteSpreadsheetByName(spreadsheetName);
    }
}

