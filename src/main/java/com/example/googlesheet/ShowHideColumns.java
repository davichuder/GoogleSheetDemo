package com.example.googlesheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.DimensionProperties;
import com.google.api.services.sheets.v4.model.DimensionRange;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.UpdateDimensionPropertiesRequest;

public class ShowHideColumns {

    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void hideColumn(String spreadsheetId, String sheetName, String column) throws IOException {
        int sheetId = getSheetId(spreadsheetId, sheetName);
        int columnIndex = getStartColumnIndex(spreadsheetId, sheetId, column);
        verifyStartColumnIndex(spreadsheetId, sheetId, columnIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("COLUMNS")
                        .setStartIndex(columnIndex)
                        .setEndIndex(columnIndex + 1))
                .setProperties(new DimensionProperties().setHiddenByUser(true))
                .setFields("hiddenByUser")));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Hidden column at index " + columnIndex);
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

    public static void showColumn(String spreadsheetId, String sheetName, String column) throws IOException {
        int sheetId = getSheetId(spreadsheetId, sheetName);
        int columnIndex = getStartColumnIndex(spreadsheetId, sheetId, column);
        verifyStartColumnIndex(spreadsheetId, sheetId, columnIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("COLUMNS")
                        .setStartIndex(columnIndex)
                        .setEndIndex(columnIndex + 1))
                .setProperties(new DimensionProperties().setHiddenByUser(false))
                .setFields("hiddenByUser")));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Shown column at index " + columnIndex);
    }

    public static void main(String[] args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        String column = "C";
        hideColumn(spreadsheetId, sheetName, column);

        showColumn(spreadsheetId, sheetName, column);
    }
}
