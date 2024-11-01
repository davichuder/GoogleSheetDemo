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

public class ShowHideRows {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void hideRow(String spreadsheetId, String sheetName, int row) throws IOException {
        int rowIndex = row - 1;
        int sheetId = getSheetId(spreadsheetId, sheetName);
        verifyStartRowIndex(spreadsheetId, sheetId, rowIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("ROWS")
                        .setStartIndex(rowIndex)
                        .setEndIndex(rowIndex + 1))
                .setProperties(new DimensionProperties().setHiddenByUser(true))
                .setFields("hiddenByUser")));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Hidden row at index " + rowIndex);
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

    public static void showRow(String spreadsheetId, String sheetName, int row) throws IOException {
        int rowIndex = row - 1;
        int sheetId = getSheetId(spreadsheetId, sheetName);
        verifyStartRowIndex(spreadsheetId, sheetId, rowIndex);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setUpdateDimensionProperties(new UpdateDimensionPropertiesRequest()
                .setRange(new DimensionRange()
                        .setSheetId(sheetId)
                        .setDimension("ROWS")
                        .setStartIndex(rowIndex)
                        .setEndIndex(rowIndex + 1))
                .setProperties(new DimensionProperties().setHiddenByUser(false))
                .setFields("hiddenByUser")));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Shown row at index " + rowIndex);
    }

    public static void main(String[] args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        int row = 2;
        hideRow(spreadsheetId, sheetName, row);

        showRow(spreadsheetId, sheetName, row);
    }
}
