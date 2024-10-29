package com.example.googlesheet;

import java.io.IOException;
import java.util.Collections;
import java.util.List;

import com.google.api.client.googleapis.json.GoogleJsonError;
import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;

/* Class to demonstrate the use of Spreadsheet Update Values API */
public class UpdateValues {
  /**
   * Sets values in a range of a spreadsheet.
   *
   * @param spreadsheetId    - Id of the spreadsheet.
   * @param range            - Range of cells of the spreadsheet.
   * @param valueInputOption - Determines how input data should be interpreted.
   * @param values           - List of rows of values to input.
   * @return spreadsheet with updated values
   * @throws IOException - if credentials file not found.
   */
  public static UpdateValuesResponse updateValues(String spreadsheetId,
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

    UpdateValuesResponse result = null;
    try {
      // Updates the values in the specified range.
      ValueRange body = new ValueRange()
          .setValues(values);
      result = service.spreadsheets().values().update(spreadsheetId, range, body)
          .setValueInputOption(valueInputOption)
          .execute();
      System.out.printf("%d cells updated.", result.getUpdatedCells());
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