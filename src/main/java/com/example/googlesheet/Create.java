package com.example.googlesheet;

import java.io.IOException;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.Permission;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.SpreadsheetProperties;

public class Create {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();
    private static final Drive driveService = GoogleConfig.getDriveService();

    public static String createSpreadsheet(String title) throws IOException {
        Spreadsheet spreadsheet = new Spreadsheet()
                .setProperties(new SpreadsheetProperties()
                        .setTitle(title));
        spreadsheet = sheetsService.spreadsheets().create(spreadsheet)
                .setFields("spreadsheetId")
                .execute();

        // Get the new spreadsheet id
        String fileId = spreadsheet.getSpreadsheetId();

        Permission userPermission = new Permission()
                .setType("user")
                .setRole("writer") // Cambia de "owner" a "writer" por ahora
                .setEmailAddress("rolondavid95@gmail.com");  // Coloca aquí tu email personal

        // Intenta sin `setTransferOwnership(true)` para ver si solo con "writer" funciona
        driveService.permissions().create(fileId, userPermission).execute();

        // Prints the new spreadsheet id
        System.out.println("Spreadsheet ID: " + fileId);
        return fileId;
    }

    public static void main(String... args) throws IOException {
        createSpreadsheet("Test Spreadsheet Demo2");
    }
}
