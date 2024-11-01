package com.example.googlesheet;

import java.io.IOException;

import com.google.api.services.drive.Drive;

public class DeleteSpreadsheet {
    private static final Drive driveService = GoogleConfig.getDriveService();
    public static void deleteSpreadsheet(String spreadsheetId) throws IOException {
        try {
            driveService.files().delete(spreadsheetId).execute();
            System.out.println("Spreadsheet with ID " + spreadsheetId + " has been deleted.");
        } catch (IOException e) {
            System.out.println("Spreadsheet with ID " + spreadsheetId + " does not exist.");
        }
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "18yyYA0gNEpt1Sn1xObNUWDH0tYNThKuKVE_bS464Xmg";
        deleteSpreadsheet(spreadsheetId);
    }
}
