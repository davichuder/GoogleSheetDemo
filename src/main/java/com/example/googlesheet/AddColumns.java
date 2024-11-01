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

public class AddColumns {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void addColumn(String spreadsheetId, String sheetName, String column) throws IOException {
        // Siempre se corre para la derecha el resto de las columnas
        int sheetId = getSheetId(spreadsheetId, sheetName);
        int columnIndex = getStartColumnIndex(spreadsheetId, sheetId, column);
        verifyStartColumnIndex(spreadsheetId, sheetId, columnIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setInsertDimension(new InsertDimensionRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("COLUMNS")
                        .setStartIndex(columnIndex)
                        .setEndIndex(columnIndex + 1))
                .setInheritFromBefore(true)));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Inserted column at index " + columnIndex);
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

    public static void addColumns(String spreadsheetId, String sheetName, String startColumn, int columnCount) throws IOException {
        int sheetId = getSheetId(spreadsheetId, sheetName);
        int startColumnIndex = getStartColumnIndex(spreadsheetId, sheetId, startColumn);
        verifyStartColumnIndex(spreadsheetId, sheetId, startColumnIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setInsertDimension(new InsertDimensionRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("COLUMNS")
                        .setStartIndex(startColumnIndex)
                        .setEndIndex(startColumnIndex + columnCount))
                .setInheritFromBefore(true)));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Inserted columns from index " + startColumnIndex + " to " + (startColumnIndex + columnCount - 1));
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        String column = "C";
        addColumn(spreadsheetId, sheetName, column);

        String column2 = "E";
        addColumns(spreadsheetId, sheetName, column2, 3);
    }
}
