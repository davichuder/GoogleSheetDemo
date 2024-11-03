package com.example.googlesheet;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.Permission;
import com.google.api.services.drive.model.File;

import java.io.IOException;

public class SpreadsheetSharing {

    private static final Drive driveService = GoogleConfig.getDriveService();

    public static String getShareableLink(String spreadsheetId, String role) throws IOException {
        Permission permission = new Permission();
        permission.setType("anyone");
        permission.setRole(role);

        driveService.permissions().create(spreadsheetId, permission).execute();

        File file = driveService.files().get(spreadsheetId)
                .setFields("webViewLink")
                .execute();

        return file.getWebViewLink();
    }

    public static void main(String[] args) {
        String spreadsheetId = "1fzuzZjLYraBbcM_YhBG50m_TyEuUxy8krZXXJhYU4H4";
        // Puede ser reader, commenter, o writer, hay otros pero creo vamos usar esos generalmente
        String role = "reader";
        try {
            String shareableLink = getShareableLink(spreadsheetId, role);
            System.out.println("Link para compartir: " + shareableLink);
        } catch (IOException e) {
            System.err.println("Error al obtener el enlace de acceso: " + e.getMessage());
        }
    }
}

