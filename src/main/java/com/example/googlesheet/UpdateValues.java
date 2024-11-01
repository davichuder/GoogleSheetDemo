package com.example.googlesheet;

import java.io.IOException;
import java.util.List;

import com.google.api.client.googleapis.json.GoogleJsonError;
import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

public class UpdateValues {
    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static UpdateValuesResponse updateValues(String spreadsheetId,
            String range,
            String valueInputOption,
            List<List<Object>> values)
            throws IOException {

        UpdateValuesResponse result = null;
        try {
            // Updates the values in the specified range.
            ValueRange body = new ValueRange()
                    .setValues(values);
            result = sheetsService.spreadsheets().values().update(spreadsheetId, range, body)
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

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String range = "Sheet1!A1:E1";
        String valueInputOption = "USER_ENTERED";
        List<List<Object>> values = List.of(List.of("Alpha", "Beta", "Gamma", "Delta", "Epsilon"));
        updateValues(spreadsheetId, range, valueInputOption, values);

        String range2 = "Sheet1!A2:A6";
        String valueInputOption2 = "USER_ENTERED";
        List<List<Object>> values2 = List.of(List.of("1"),
                List.of("2"), List.of("3"), List.of("4"), List.of("5"));
        updateValues(spreadsheetId, range2, valueInputOption2, values2);
    }
}