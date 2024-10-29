package com.example.googlesheet;

import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;
import java.io.IOException;
import java.util.Arrays;

public class DeleteSpreadsheet {

    public static void deleteSpreadsheet(String spreadsheetId) throws IOException {
        GoogleCredentials credentials = GoogleCredentials.getApplicationDefault()
                .createScoped(Arrays.asList(
                    SheetsScopes.SPREADSHEETS,
                    DriveScopes.DRIVE
                ));
        HttpCredentialsAdapter requestInitializer = new HttpCredentialsAdapter(credentials);

        Drive driveService = new Drive.Builder(new NetHttpTransport(), GsonFactory.getDefaultInstance(), requestInitializer)
                .setApplicationName("Drive API Snippet")
                .build();

        // Eliminar la hoja de cálculo
        driveService.files().delete(spreadsheetId).execute();
        System.out.println("Spreadsheet with ID " + spreadsheetId + " has been deleted.");
    }

    public static void main(String... args) throws IOException {
        // Reemplaza con el ID de la hoja de cálculo que quieres eliminar
        String spreadsheetId = "18yyYA0gNEpt1Sn1xObNUWDH0tYNThKuKVE_bS464Xmg";
        deleteSpreadsheet(spreadsheetId);
    }
}
