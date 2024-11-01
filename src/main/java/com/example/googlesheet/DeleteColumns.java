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

public class DeleteColumns {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void deleteRightColumnsFrom(String spreadsheetId, String sheetName, String column)
            throws IOException {
        int sheetId = getSheetId(spreadsheetId, sheetName);
        int startColumnIndex = getStartColumnIndex(spreadsheetId, sheetId, column);
        verifyStartColumnIndex(spreadsheetId, sheetId, startColumnIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setDeleteDimension(new DeleteDimensionRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("COLUMNS")
                        .setStartIndex(startColumnIndex))));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Deleted columns from index " + startColumnIndex + " to the end.");
    }

    private static void verifyStartColumnIndex(String spreadsheetId, int sheetId, int startColumnIndex)
            throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        Sheet sheet = spreadsheet.getSheets().get(sheetId);
        if (startColumnIndex > sheet.getProperties().getGridProperties().getColumnCount()) {
            throw new IllegalArgumentException("Start column index is out of range.");
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

    private static int getStartColumnIndex(String spreadsheetId, int sheetId, String column)
            throws IOException {
        int result = 0;
        for (int i = 0; i < column.length(); i++) {
            result *= 26;
            result += column.charAt(i) - 'A' + 1;
        }
        return result - 1;
    }

    public static void deleteColumn(String spreadsheetId, String sheetName, String column)
            throws IOException {
        int sheetId = getSheetId(spreadsheetId, sheetName);
        int startColumnIndex = getStartColumnIndex(spreadsheetId, sheetId, column) - 1;
        verifyStartColumnIndex(spreadsheetId, sheetId, startColumnIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setDeleteDimension(new DeleteDimensionRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("COLUMNS")
                        .setStartIndex(startColumnIndex)
                        .setEndIndex(startColumnIndex + 1))));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Deleted column at index " + startColumnIndex);
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";

        String column = "C";
        deleteColumn(spreadsheetId, sheetName, column);

        String column2 = "O";
        deleteRightColumnsFrom(spreadsheetId, sheetName, column2);
    }
}
