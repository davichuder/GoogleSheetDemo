package com.example.googlesheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.google.api.client.googleapis.json.GoogleJsonError;
import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.AddConditionalFormatRuleRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.BooleanCondition;
import com.google.api.services.sheets.v4.model.BooleanRule;
import com.google.api.services.sheets.v4.model.CellFormat;
import com.google.api.services.sheets.v4.model.Color;
import com.google.api.services.sheets.v4.model.ConditionValue;
import com.google.api.services.sheets.v4.model.ConditionalFormatRule;
import com.google.api.services.sheets.v4.model.DataValidationRule;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.SetDataValidationRequest;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

public class Format {

    private static Sheets sheetsService = GoogleConfig.getSheetsService();

    public static UpdateValuesResponse updateValues(String spreadsheetId,
            String range,
            String valueInputOption,
            List<List<Object>> values)
            throws IOException {

        UpdateValuesResponse result = null;
        try {
            ValueRange body = new ValueRange().setValues(values);
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

    public static int getSheetId(String spreadsheetId, String sheetName) throws IOException {
        Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
        for (Sheet sheet : spreadsheet.getSheets()) {
            if (sheet.getProperties().getTitle().equals(sheetName)) {
                return sheet.getProperties().getSheetId();
            }
        }
        throw new IllegalArgumentException("Sheet with name " + sheetName + " not found");
    }

    public static void applyConditionalFormatting(String spreadsheetId, String range) throws IOException {
        List<Request> requests = new ArrayList<>();

        // Obtener el nombre de la hoja y el rango de celdas a formatear
        String sheetName = range.split("!")[0];
        String[] rangeParts = range.split("!")[1].split(":");

        int startColumn = rangeParts[0].replaceAll("\\d", "").charAt(0) - 'A';
        int startRow = Integer.parseInt(rangeParts[0].replaceAll("[^\\d]", "")) - 1;
        int endColumn = rangeParts[1].replaceAll("\\d", "").charAt(0) - 'A' + 1;
        int endRow = Integer.parseInt(rangeParts[1].replaceAll("[^\\d]", ""));

        int sheetId = getSheetId(spreadsheetId, sheetName);

        // Colores para las reglas
        Color headerColor = new Color().setRed(0.6f).setGreen(0.8f).setBlue(1.0f); // Color para la cabecera
        Color oddRowColor = new Color().setRed(0.9f).setGreen(0.9f).setBlue(0.9f); // Color para filas impares
        Color evenRowColor = new Color().setRed(1.0f).setGreen(1.0f).setBlue(1.0f); // Color para filas pares

        // Regla para la cabecera
        ConditionalFormatRule headerRule = new ConditionalFormatRule()
                .setRanges(List.of(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(startRow)
                        .setEndRowIndex(startRow + 1)
                        .setStartColumnIndex(startColumn)
                        .setEndColumnIndex(endColumn)))
                .setBooleanRule(new BooleanRule()
                        .setCondition(new BooleanCondition()
                                .setType("CUSTOM_FORMULA")
                                .setValues(List.of(new ConditionValue()
                                        .setUserEnteredValue("=TRUE"))))
                        .setFormat(new CellFormat().setBackgroundColor(headerColor)));

        // Regla para filas impares
        ConditionalFormatRule oddRowRule = new ConditionalFormatRule()
                .setRanges(List.of(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(startRow + 1)
                        .setEndRowIndex(endRow)
                        .setStartColumnIndex(startColumn)
                        .setEndColumnIndex(endColumn)))
                .setBooleanRule(new BooleanRule()
                        .setCondition(new BooleanCondition()
                                .setType("CUSTOM_FORMULA")
                                .setValues(List.of(new ConditionValue()
                                        .setUserEnteredValue("=ISODD(ROW())"))))
                        .setFormat(new CellFormat().setBackgroundColor(oddRowColor)));

        // Regla para filas pares
        ConditionalFormatRule evenRowRule = new ConditionalFormatRule()
                .setRanges(List.of(new GridRange()
                        .setSheetId(sheetId)
                        .setStartRowIndex(startRow + 1)
                        .setEndRowIndex(endRow)
                        .setStartColumnIndex(startColumn)
                        .setEndColumnIndex(endColumn)))
                .setBooleanRule(new BooleanRule()
                        .setCondition(new BooleanCondition()
                                .setType("CUSTOM_FORMULA")
                                .setValues(List.of(new ConditionValue()
                                        .setUserEnteredValue(
                                                "=ISEVEN(ROW())"))))
                        .setFormat(new CellFormat().setBackgroundColor(evenRowColor)));

        // Agregar las reglas a la solicitud de actualización
        requests.add(new Request().setAddConditionalFormatRule(new AddConditionalFormatRuleRequest()
                .setRule(headerRule)
                .setIndex(0)));

        requests.add(new Request().setAddConditionalFormatRule(new AddConditionalFormatRuleRequest()
                .setRule(oddRowRule)
                .setIndex(0)));

        requests.add(new Request().setAddConditionalFormatRule(new AddConditionalFormatRuleRequest()
                .setRule(evenRowRule)
                .setIndex(0)));

        // Ejecutar la solicitud
        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.println("Conditional formatting applied with alternate row colors and header color.");
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        String range = "Sheet1!A1:O41";
        String valueInputOption = "USER_ENTERED";

        List<List<Object>> values = new ArrayList<>();
        values.add(List.of("# invitado", "Apellido", "Nombre", "ESTADO (ESCONDIDA)", "Apodo", "Email",
                "Celular", "Género", "Idioma", "Idioma (otro)", "Grupo", "Grupo (otro)",
                "Relac. con homenajeado/a", "Relac. con Usted", "Nota"));

        for (int i = 1; i <= 40; i++) {
            List<Object> row = new ArrayList<>();
            row.add(i);
            for (int j = 1; j < 15; j++) {
                row.add("");
            }
            values.add(row);
        }

        updateValues(spreadsheetId, range, valueInputOption, values);

        // Apply conditional formatting
        String rangeFormat = "Sheet1!A1:O41";
        applyConditionalFormatting(spreadsheetId, rangeFormat);
    }
}
