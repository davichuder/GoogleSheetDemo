package com.example.googlesheet;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;
import com.google.api.client.http.HttpRequestInitializer;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;

import java.io.IOException;
import java.util.Collections;

public class ListFiles {

    public static void main(String... args) throws IOException {
        GoogleCredentials credentials = GoogleCredentials.getApplicationDefault()
                .createScoped(Collections.singleton(DriveScopes.DRIVE));
        HttpRequestInitializer requestInitializer = new HttpCredentialsAdapter(credentials);

        // Create the Drive API client
        Drive driveService = new Drive.Builder(new NetHttpTransport(), GsonFactory.getDefaultInstance(), requestInitializer)
                .setApplicationName("Drive API Snippet")
                .build();

        // List the files in the Drive
        FileList result = driveService.files().list()
                .setPageSize(10)
                .setFields("nextPageToken, files(id, name)")
                .execute();
        java.util.List<File> files = result.getFiles();
        if (files == null || files.isEmpty()) {
            System.out.println("No files found.");
        } else {
            System.out.println("Files:");
            for (File file : files) {
                System.out.printf("%s (%s)\n", file.getName(), file.getId());
            }
        }
    }
}
