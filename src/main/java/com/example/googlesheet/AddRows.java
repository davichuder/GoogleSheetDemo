package com.example.googlesheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.DimensionRange;
import com.google.api.services.sheets.v4.model.InsertDimensionRequest;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;

public class AddRows {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void addRow(String spreadsheetId, String sheetName, int row) throws IOException {
        // Siempre mueve las otas filas para abajo
        int sheetId = getSheetId(spreadsheetId, sheetName);
        verifyStartRowIndex(spreadsheetId, sheetId, row);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setInsertDimension(new InsertDimensionRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("ROWS")
                        .setStartIndex(row)
                        .setEndIndex(row + 1))
                .setInheritFromBefore(true)));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Inserted row at index " + row);
    }

    private static void verifyStartRowIndex(String spreadsheetId, int sheetId, int startRowIndex)
            throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        Sheet sheet = spreadsheet.getSheets().get(sheetId);
        if (startRowIndex > sheet.getProperties().getGridProperties().getRowCount()) {
            throw new IllegalArgumentException("Start row index is out of range.");
        }
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

    public static void addRows(String spreadsheetId, String sheetName, int startRow, int rowCount) throws IOException {
        int sheetId = getSheetId(spreadsheetId, sheetName);
        verifyStartRowIndex(spreadsheetId, sheetId, startRow);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setInsertDimension(new InsertDimensionRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("ROWS")
                        .setStartIndex(startRow)
                        .setEndIndex(startRow + rowCount))
                .setInheritFromBefore(true)));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Inserted rows from index " + startRow + " to " + (startRow + rowCount - 1));
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        int row = 2;
        addRow(spreadsheetId, sheetName, row);

        int startRow = 4;
        addRows(spreadsheetId, sheetName, startRow, 3);
    }
}
