package com.example.googlesheet;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import com.google.api.client.googleapis.json.GoogleJsonError;
import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;

/* Class to demonstrate the use of Spreadsheet Batch Update Values API */
public class BatchUpdateValues {
  /**
   * Set values in one or more ranges of spreadsheet.
   *
   * @param spreadsheetId    - Id of the spreadsheet.
   * @param range            - Range of cells of the spreadsheet.
   * @param valueInputOption - Determines how input data should be interpreted.
   * @param values           - list of rows of values to input.
   * @return spreadsheet with updated values
   * @throws IOException - if credentials file not found.
   */
  public static BatchUpdateValuesResponse batchUpdateValues(String spreadsheetId,
                                                            String range,
                                                            String valueInputOption,
                                                            List<List<Object>> values)
      throws IOException {
        /* Load pre-authorized user credentials from the environment.
           TODO(developer) - See https://developers.google.com/identity for
            guides on implementing OAuth2 for your application. */
    GoogleCredentials credentials = GoogleCredentials.getApplicationDefault()
        .createScoped(Collections.singleton(SheetsScopes.SPREADSHEETS));
    HttpRequestInitializer requestInitializer = new HttpCredentialsAdapter(
        credentials);

    // Create the sheets API client
    Sheets service = new Sheets.Builder(new NetHttpTransport(),
        GsonFactory.getDefaultInstance(),
        requestInitializer)
        .setApplicationName("Sheets samples")
        .build();

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
      result = service.spreadsheets().values().batchUpdate(spreadsheetId, body).execute();
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