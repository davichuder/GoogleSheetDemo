package com.example.googlesheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.google.api.client.googleapis.json.GoogleJsonError;
import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

public class BatchUpdateValues {
  private static final Sheets sheetsService = GoogleConfig.getSheetsService();

  public static BatchUpdateValuesResponse batchUpdateValues(String spreadsheetId,
      String range,
      String valueInputOption,
      List<List<Object>> values)
      throws IOException {

    List<ValueRange> data = new ArrayList<>();
    data.add(new ValueRange()
        .setRange(range)
        .setValues(values));

    BatchUpdateValuesResponse result = null;
    try {
      // Updates the values in the specified range.
      BatchUpdateValuesRequest body = new BatchUpdateValuesRequest()
          .setValueInputOption(valueInputOption)
          .setData(data);
      result = sheetsService.spreadsheets().values().batchUpdate(spreadsheetId, body).execute();
      System.out.printf("%d cells updated.", result.getTotalUpdatedCells());
    } catch (GoogleJsonResponseException e) {
      // TODO(developer) - handle error appropriately
      GoogleJsonError error = e.getDetails();
      if (error.getCode() == 404) {
        System.out.printf("Spreadsheet not found with id '%s'.\n", spreadsheetId);
      } else {
        throw e;
      }
    }
    return result;
  }

  public static void main(String[] args) throws IOException {
    String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
    String range = "Sheet1!A1:E1";
    String valueInputOption = "USER_ENTERED";
    List<List<Object>> values = List.of(List.of("Alpha", "Beta", "Gamma", "Delta", "Epsilon"));
    batchUpdateValues(spreadsheetId, range, valueInputOption, values);
  }
}