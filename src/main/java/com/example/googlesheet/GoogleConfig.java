package com.example.googlesheet;

import java.io.IOException;
import java.util.Arrays;

import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;

public class GoogleConfig {
    private static HttpRequestInitializer getRequestInitializer() throws IOException {
        GoogleCredentials credentials = GoogleCredentials.getApplicationDefault()
                .createScoped(Arrays.asList(
                        SheetsScopes.SPREADSHEETS,
                        DriveScopes.DRIVE));
        return new HttpCredentialsAdapter(
                credentials);
    }

    public static Sheets getSheetsService() {
        HttpRequestInitializer requestInitializer;
        try {
            requestInitializer = getRequestInitializer();
        } catch (IOException e) {
            return null;
        }
        return new Sheets.Builder(new NetHttpTransport(), GsonFactory.getDefaultInstance(), requestInitializer)
                .setApplicationName("Sheets API Snippet")
                .build();
    }

    public static Drive getDriveService() {
        HttpRequestInitializer requestInitializer;
        try {
            requestInitializer = getRequestInitializer();
        } catch (IOException e) {
            return null;
        }
        return new Drive.Builder(new NetHttpTransport(), GsonFactory.getDefaultInstance(), requestInitializer)
                .setApplicationName("Drive API Snippet")
                .build();
    }
}
