package com.example.googlesheet;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.UpdateSheetPropertiesRequest;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class RenameSheet {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void main(String[] args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        String newSheetName = "Lista de Invitados";
        
        int sheetId = getSheetId(spreadsheetId, sheetName);

        // Crear la solicitud para cambiar el nombre de la hoja
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setUpdateSheetProperties(new UpdateSheetPropertiesRequest()
                .setProperties(new SheetProperties()
                        .setSheetId(sheetId)
                        .setTitle(newSheetName))
                .setFields("title")));

        // Ejecutar la solicitud
        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();

        System.out.println("Nombre de la hoja cambiado a: " + newSheetName);
    }

    public static int getSheetId(String spreadsheetId, String sheetName) throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        for (Sheet sheet : spreadsheet.getSheets()) {
            if (sheet.getProperties().getTitle().equals(sheetName)) {
                return sheet.getProperties().getSheetId();
            }
        }
        throw new IllegalArgumentException("Sheet with name " + sheetName + " not found");
    }
}

