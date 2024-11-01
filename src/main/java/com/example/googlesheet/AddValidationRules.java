package com.example.googlesheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.BooleanCondition;
import com.google.api.services.sheets.v4.model.ConditionValue;
import com.google.api.services.sheets.v4.model.DataValidationRule;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.SetDataValidationRequest;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;

public class AddValidationRules {

    private static final Sheets sheetsService = GoogleConfig.getSheetsService();

    public static void applyValidationRule(String spreadsheetId, int sheetId, String column, String ruleType) throws IOException {
        int startColumnIndex = getStartColumnIndex(column);
        final int firstRow = 1;

        GridRange range = new GridRange()
                .setSheetId(sheetId)
                .setStartColumnIndex(startColumnIndex)
                .setEndColumnIndex(startColumnIndex + 1)
                .setStartRowIndex(firstRow);

        DataValidationRule rule = new DataValidationRule();

        String regex;
        switch (ruleType) {
            case "EMAIL":
                rule.setCondition(new BooleanCondition()
                        .setType("TEXT_IS_EMAIL"));
                break;

            // FIX: Texto sin numeros
            case "TEXT":
                regex = "=REGEXMATCH(" + column + (firstRow + 1) + ", \"^[A-Za-zÀ-ÿ\\s']+$\")";
                rule.setCondition(new BooleanCondition()
                        .setType("CUSTOM_FORMULA")
                        .setValues(Collections.singletonList(
                                new ConditionValue().setUserEnteredValue(regex))));
                break;

            case "NUMBER":
                regex = "=REGEXMATCH(TO_TEXT(" + column + (firstRow + 1) + "); \"^\\d+$\")";
                rule.setCondition(new BooleanCondition()
                        .setType("CUSTOM_FORMULA")
                        .setValues(Collections.singletonList(
                                new ConditionValue().setUserEnteredValue(regex))));
                break;
            case "DUPLICATE":
                regex = "=COUNTIF(" + column + ":" + column + ", " + column + (firstRow + 1) + ")=1";
                rule.setCondition(new BooleanCondition()
                        .setType("CUSTOM_FORMULA")
                        .setValues(Collections.singletonList(
                                new ConditionValue().setUserEnteredValue(regex))));
                break;

            default:
                throw new IllegalArgumentException("Invalid rule type specified");
        }

        rule.setInputMessage("Please enter a valid " + ruleType.toLowerCase());

        Request validationRequest = new Request()
                .setSetDataValidation(new SetDataValidationRequest()
                        .setRange(range)
                        .setRule(rule));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest()
                .setRequests(Collections.singletonList(validationRequest));

        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Applied " + ruleType + " validation to column " + column);
    }

    private static int getStartColumnIndex(String column) {
        int result = 0;
        for (int i = 0; i < column.length(); i++) {
            result *= 26;
            result += column.charAt(i) - 'A' + 1;
        }
        return result - 1;
    }

    public static void applyDataValidation(String spreadsheetId, String range, List<String> validationList)
            throws IOException {

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

        int sheetId = getSheetId(spreadsheetId, sheetName);

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
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String sheetName = "Sheet1";
        int sheetId = getSheetId(spreadsheetId, sheetName);

        applyValidationRule(spreadsheetId, sheetId, "F", "EMAIL");
        applyValidationRule(spreadsheetId, sheetId, "C", "TEXT");
        applyValidationRule(spreadsheetId, sheetId, "A", "NUMBER");
        applyValidationRule(spreadsheetId, sheetId, "A", "DUPLICATE");

        List<String> validationList = List.of("Male", "Female", "Other");
        String rangeValidation = "Sheet1!H2:H41";
        applyDataValidation(spreadsheetId, rangeValidation, validationList);

        List<String> validationList2 = List.of("Familia cercana", "Amigos", "Trabajo");
        String rangeValidation2 = "Sheet1!K2:K41";
        applyDataValidation(spreadsheetId, rangeValidation2, validationList2);
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
}
