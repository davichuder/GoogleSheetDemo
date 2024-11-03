package com.example.googlesheet;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.Spreadsheet;

public class EditFilename {

    private static final Drive driveService = GoogleConfig.getDriveService();
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void main(String[] args) throws Exception {

        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String originalName = "Test Spreadsheet Demo";

        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();

        if (!spreadsheet.getProperties().getTitle().equals(originalName)) {
            File updatedFile = new File();
            updatedFile.setName(originalName);
            driveService.files().update(spreadsheetId, updatedFile).execute();
            System.out.println("Nombre de la hoja de cálculo revertido a: " + originalName);
        }
    }
}
