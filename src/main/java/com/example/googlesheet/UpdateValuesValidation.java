package com.example.googlesheet;

import com.google.api.client.googleapis.json.GoogleJsonError;
import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;
import com.google.api.services.sheets.v4.SheetsScopes;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class UpdateValuesValidation {

    public static UpdateValuesResponse updateValues(String spreadsheetId,
            String range,
            String valueInputOption,
            List<List<Object>> values)
            throws IOException {

        GoogleCredentials credentials = GoogleCredentials.getApplicationDefault()
                .createScoped(Collections.singleton(SheetsScopes.SPREADSHEETS));
        HttpRequestInitializer requestInitializer = new HttpCredentialsAdapter(credentials);

        Sheets service = new Sheets.Builder(new NetHttpTransport(),
                GsonFactory.getDefaultInstance(),
                requestInitializer)
                .setApplicationName("Sheets samples")
                .build();

        UpdateValuesResponse result = null;
        try {
            ValueRange body = new ValueRange().setValues(values);
            result = service.spreadsheets().values().update(spreadsheetId, range, body)
                    .setValueInputOption(valueInputOption)
                    .execute();
            System.out.printf("%d cells updated.", result.getUpdatedCells());
        } catch (GoogleJsonResponseException e) {
            GoogleJsonError error = e.getDetails();
            if (error.getCode() == 404) {
                System.out.printf("Spreadsheet not found with id '%s'.\n", spreadsheetId);
            } else {
                throw e;
            }
        }
        return result;
    }

    public static void applyDataValidation(String spreadsheetId, String range, List<String> validationList) throws IOException {
        GoogleCredentials credentials = GoogleCredentials.getApplicationDefault()
                .createScoped(Collections.singleton(SheetsScopes.SPREADSHEETS));
        HttpRequestInitializer requestInitializer = new HttpCredentialsAdapter(credentials);

        Sheets service = new Sheets.Builder(new NetHttpTransport(),
                GsonFactory.getDefaultInstance(),
                requestInitializer)
                .setApplicationName("Sheets samples")
                .build();

        List<Request> requests = new ArrayList<>();

        // Convert validation list to ConditionValue list
        List<ConditionValue> conditionValues = new ArrayList<>();
        for (String value : validationList) {
            conditionValues.add(new ConditionValue().setUserEnteredValue(value));
        }

        // Define the data validation rule
        DataValidationRule validationRule = new DataValidationRule()
                .setCondition(new BooleanCondition()
                        .setType("ONE_OF_LIST")
                        .setValues(conditionValues))
                .setStrict(true)
                .setShowCustomUi(true);

        String sheetName = range.split("!")[0];
        String[] rangeParts = range.split("!")[1].split(":");

        int startColumn = rangeParts[0].replaceAll("\\d", "").charAt(0) - 'A';
        int startRow = Integer.parseInt(rangeParts[0].replaceAll("[^\\d]", "")) - 1;

        int endColumn = rangeParts[1].replaceAll("\\d", "").charAt(0) - 'A' + 1;
        int endRow = Integer.parseInt(rangeParts[1].replaceAll("[^\\d]", ""));


        int sheetId = getSheetId(service, spreadsheetId, sheetName);

        // Apply the data validation rule to the range
        requests.add(new Request().setSetDataValidation(new SetDataValidationRequest()
                .setRange(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(startRow)
                        .setEndRowIndex(endRow)
                        .setStartColumnIndex(startColumn)
                        .setEndColumnIndex(endColumn))
                .setRule(validationRule)));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        service.spreadsheets().batchUpdate(spreadsheetId, body).execute();
    }

    public static int getSheetId(Sheets service, String spreadsheetId, String sheetName) throws IOException {
        Spreadsheet spreadsheet = service.spreadsheets().get(spreadsheetId).execute();
        for (Sheet sheet : spreadsheet.getSheets()) {
            if (sheet.getProperties().getTitle().equals(sheetName)) {
                return sheet.getProperties().getSheetId();
            }
        }
        throw new IllegalArgumentException("Sheet with name " + sheetName + " not found");
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String range = "Sheet1!A1:O41";
        String valueInputOption = "USER_ENTERED";

        List<List<Object>> values = new ArrayList<>();
        values.add(List.of("# invitado", "Apellido", "Nombre", "ESTADO (ESCONDIDA)", "Apodo", "Email", "Celular", "Género", "Idioma", "Idioma (otro)", "Grupo", "Grupo (otro)", "Relac. con homenajeado/a", "Relac. con Usted", "Nota"));

        for (int i = 1; i <= 40; i++) {
            List<Object> row = new ArrayList<>();
            row.add(i);
            for (int j = 1; j < 15; j++) {
                row.add("");
            }
            values.add(row);
        }

        updateValues(spreadsheetId, range, valueInputOption, values);

        // Define the validation list
        List<String> validationList = List.of("Male", "Female", "Other");

        // Apply data validation
        String rangeValidation = "Sheet1!H2:H2";
        applyDataValidation(spreadsheetId, rangeValidation, validationList);
    }
}
