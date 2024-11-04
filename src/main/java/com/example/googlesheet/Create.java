package com.example.googlesheet;

import java.io.IOException;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.SpreadsheetProperties;

public class Create {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static String createSpreadsheet(String title) throws IOException {
        Spreadsheet spreadsheet = new Spreadsheet()
                .setProperties(new SpreadsheetProperties()
                        .setTitle(title));
        spreadsheet = sheetsService.spreadsheets().create(spreadsheet)
                .setFields("spreadsheetId")
                .execute();

        // Get the new spreadsheet id
        String fileId = spreadsheet.getSpreadsheetId();

        // Prints the new spreadsheet id
        System.out.println("Spreadsheet ID: " + fileId);
        return fileId;
    }

    public static void main(String... args) throws IOException {
        createSpreadsheet("Test Spreadsheet Demo2");
    }
}
