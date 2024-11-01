package com.example.googlesheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.DeleteDimensionRequest;
import com.google.api.services.sheets.v4.model.DimensionRange;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;

public class DeleteRows {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void deleteRow(String spreadsheetId, String sheetName, int rowIndex)
            throws IOException {
        rowIndex--;
        int sheetId = getSheetId(spreadsheetId, sheetName);
        verifyStartRowIndex(spreadsheetId, sheetId, rowIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setDeleteDimension(new DeleteDimensionRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("ROWS")
                        .setStartIndex(rowIndex)
                        .setEndIndex(rowIndex + 1))));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Deleted row at index " + rowIndex);
    }

    private static void verifyStartRowIndex(String spreadsheetId, int sheetId, int rowIndex) throws IOException {
        if (rowIndex < 0) {
        throw new IllegalArgumentException("Start row index is out of range");
        }
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        Sheet sheet = spreadsheet.getSheets().get(sheetId);
        if (rowIndex > sheet.getProperties().getGridProperties().getRowCount()) {
            throw new IllegalArgumentException("Start row index is out of range");
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

    public static void deleteRowsFrom(String spreadsheetId, String sheetName, int startRowIndex) throws IOException {
        startRowIndex--;
        int sheetId = getSheetId(spreadsheetId, sheetName);
        verifyStartRowIndex(spreadsheetId, sheetId, startRowIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setDeleteDimension(new DeleteDimensionRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("ROWS")
                        .setStartIndex(startRowIndex))));
    
        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Deleted rows from index " + startRowIndex + " to the end.");
    }
    

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        int rowIndex = 1;
        deleteRow(spreadsheetId, sheetName, rowIndex);

        int rowIndex2 = 10;
        deleteRowsFrom(spreadsheetId, sheetName, rowIndex2);
    }
}