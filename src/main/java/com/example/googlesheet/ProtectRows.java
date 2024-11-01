package com.example.googlesheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.AddProtectedRangeRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.Editors;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.ProtectedRange;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;

public class ProtectRows {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void protectRow(String spreadsheetId, String sheetName, int row, List<String> editors) throws IOException {
        int rowIndex = row - 1;
        int sheetId = getSheetId(spreadsheetId, sheetName);
        verifyStartRowIndex(spreadsheetId, sheetId, rowIndex);

        List<Request> requests = new ArrayList<>();
        ProtectedRange protectedRange = new ProtectedRange()
                .setRange(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(rowIndex)
                        .setEndRowIndex(rowIndex + 1))
                .setDescription("Protected Row")
                .setEditors(new Editors().setUsers(editors));

        requests.add(new Request().setAddProtectedRange(new AddProtectedRangeRequest().setProtectedRange(protectedRange)));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Protected row at index " + rowIndex);
    }

    private static int getSheetId(String spreadsheetId, String sheetName) throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        for (Sheet sheet : spreadsheet.getSheets()) {
            if (sheet.getProperties().getTitle().equals(sheetName)) {
                return sheet.getProperties().getSheetId();
            }
        }
        throw new IllegalArgumentException("Sheet with name " + sheetName + " not found");
    }

    private static void verifyStartRowIndex(String spreadsheetId, int sheetId, int startRowIndex)
            throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        Sheet sheet = spreadsheet.getSheets().get(sheetId);
        if (startRowIndex > sheet.getProperties().getGridProperties().getRowCount()) {
            throw new IllegalArgumentException("Start row index is out of range.");
        }
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        int row = 2;
        // Lista vacio unicamente el propietario puede editar
        List<String> editors = List.of();
        protectRow(spreadsheetId, sheetName, row, editors);

        // Lista con los editores que pueden editar
        List<String> editors2 = List.of("rolondavid95@gmail.com");
        protectRow(spreadsheetId, sheetName, row, editors2);
    }
}

