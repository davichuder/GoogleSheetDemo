package com.example.googlesheet;

import java.io.IOException;
import java.util.Collections;

import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;

public class DeleteByName {

    public static void deleteSpreadsheetByName(String name) throws IOException {
        GoogleCredentials credentials = GoogleCredentials.getApplicationDefault()
            .createScoped(Collections.singleton(DriveScopes.DRIVE));
        HttpRequestInitializer requestInitializer = new HttpCredentialsAdapter(credentials);

        Drive driveService = new Drive.Builder(new NetHttpTransport(),
            GsonFactory.getDefaultInstance(),
            requestInitializer)
            .setApplicationName("Delete Google Sheets")
            .build();

        // Search for the file by name
        String query = "name='" + name + "' and mimeType='application/vnd.google-apps.spreadsheet'";
        FileList result = driveService.files().list().setQ(query).execute();
        if (result.getFiles().isEmpty()) {
            System.out.println("No spreadsheet found with name: " + name);
        } else {
            for (File res: result.getFiles()) {
                String fileId = res.getId();
                driveService.files().delete(fileId).execute();
            }
            // Delete the file by ID
            System.out.println("Spreadsheets with name " + name + " deleted.");
        }
    }

    public static void main(String[] args) throws IOException {
        String spreadsheetName = "Hello World";
        deleteSpreadsheetByName(spreadsheetName);
    }
}

