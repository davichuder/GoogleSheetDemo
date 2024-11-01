package com.example.googlesheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.AddProtectedRangeRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.DeleteProtectedRangeRequest;
import com.google.api.services.sheets.v4.model.Editors;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.ProtectedRange;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;

public class ProtectColumns {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void protectColumn(String spreadsheetId, String sheetName, String column, List<String> editors) throws IOException {
        int sheetId = getSheetId(spreadsheetId, sheetName);
        int columnIndex = getStartColumnIndex(sheetId, column);
        verifyStartColumnIndex(spreadsheetId, sheetId, columnIndex);

        removeExistingProtections(spreadsheetId, sheetId, columnIndex);

        List<Request> requests = new ArrayList<>();
        ProtectedRange protectedRange = new ProtectedRange()
                .setRange(new GridRange()
                        .setSheetId(sheetId)
                        .setStartColumnIndex(columnIndex)
                        .setEndColumnIndex(columnIndex + 1))
                .setDescription("Protected Column")
                .setEditors(new Editors().setUsers(editors));

        requests.add(new Request().setAddProtectedRange(new AddProtectedRangeRequest().setProtectedRange(protectedRange)));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Protected column at index " + columnIndex);
    }

    // Métodos auxiliares para obtener el ID de la hoja y el índice de la columna
    private static int getSheetId(String spreadsheetId, String sheetName) throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        for (Sheet sheet : spreadsheet.getSheets()) {
            if (sheet.getProperties().getTitle().equals(sheetName)) {
                return sheet.getProperties().getSheetId();
            }
        }
        throw new IllegalArgumentException("Sheet with name " + sheetName + " not found");
    }

    private static int getStartColumnIndex(int sheetId, String column)
            throws IOException {
        int result = 0;
        for (int i = 0; i < column.length(); i++) {
            result *= 26;
            result += column.charAt(i) - 'A' + 1;
        }
        return result - 1;
    }

    private static void verifyStartColumnIndex(String spreadsheetId, int sheetId, int startColumnIndex)
            throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        Sheet sheet = spreadsheet.getSheets().get(sheetId);
        if (startColumnIndex > sheet.getProperties().getGridProperties().getColumnCount()) {
            throw new IllegalArgumentException("Start column index is out of range.");
        }
    }

    private static void removeExistingProtections(String spreadsheetId, int sheetId, int columnIndex) throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        List<ProtectedRange> protectedRanges = spreadsheet.getSheets().stream()
                .filter(sheet -> sheet.getProperties().getSheetId() == sheetId)
                .findFirst()
                .map(Sheet::getProtectedRanges)
                .orElse(new ArrayList<>());
    
        List<Request> deleteRequests = new ArrayList<>();
    
        for (ProtectedRange protectedRange : protectedRanges) {
            GridRange range = protectedRange.getRange();
            if (range != null && 
                range.getStartColumnIndex() != null && 
                range.getStartColumnIndex() == columnIndex) {
                
                deleteRequests.add(new Request()
                    .setDeleteProtectedRange(
                        new DeleteProtectedRangeRequest()
                            .setProtectedRangeId(protectedRange.getProtectedRangeId())
                    )
                );
            }
        }
    
        if (!deleteRequests.isEmpty()) {
            BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(deleteRequests);
            
            sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
            System.out.println("Removed " + deleteRequests.size() + " existing protections for column " + columnIndex);
        }
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        String column = "E";
        // Lista vacio unicamente el propietario puede editar
        List<String> editors = List.of();
        // protectColumn(spreadsheetId, sheetName, column, editors);

        // Lista con los editores que pueden editar
        List<String> editors2 = List.of("rolondavid95@gmail.com");
        protectColumn(spreadsheetId, sheetName, column, editors2);
    }
}

