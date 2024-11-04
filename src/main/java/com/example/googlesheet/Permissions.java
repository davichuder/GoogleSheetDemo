package com.example.googlesheet;

import java.io.IOException;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.Permission;
import com.google.api.services.drive.model.PermissionList;

public class Permissions {

    private static final Drive driveService = GoogleConfig.getDriveService();

    public static void grantPermissions(String spreadsheetId, String editor) throws IOException {
        Permission userPermission = new Permission()
                .setType("user")
                .setRole("writer") // Cambia de "owner" a "writer" por ahora
                .setEmailAddress(editor);  // Coloca aquí tu email personal

        // Intenta sin `setTransferOwnership(true)` para ver si solo con "writer" funciona
        userPermission = driveService.permissions().create(spreadsheetId, userPermission).execute();

        // Prints the new spreadsheet id
        System.out.println("Spreadsheet ID: " + spreadsheetId + " with permission: " + userPermission.getId());
    }

    public static void revokePermission(String spreadsheetId, String userEmail) throws IOException {
        PermissionList permissions = driveService.permissions()
                .list(spreadsheetId)
                .setFields("permissions(id,emailAddress,role)")
                .execute();

        if (permissions.getPermissions().isEmpty()) {
            System.out.println("No permissions found for the spreadsheet.");
            return;
        }

        // Busca el permiso del usuario especificado
        for (Permission permission : permissions.getPermissions()) {
            System.out.println("Permission ID: " + permission.getId());
            if (permission.getEmailAddress() != null && permission.getEmailAddress().equals(userEmail)) {
                String permissionId = permission.getId();
                driveService.permissions().delete(spreadsheetId, permissionId).execute();
                System.out.println("Permiso revocado exitosamente para: " + userEmail);
                return;
            }
            Permission permissionAux = driveService.permissions().get(spreadsheetId, permission.getId()).execute();
            if (permissionAux.getEmailAddress() != null) {
                System.out.println(permission.getEmailAddress());
            }
        }
        System.out.println("No se encontró ningún permiso para el usuario: " + userEmail);
    }

    public static void main(String... args) throws IOException {
        String spreadsheetId = "1wUh4rUsLtChDHPwmbluw25NOmO3Yx5mOIM0jmM-J8PU";
        String editor = "rolondavid95@gmail.com";

        // grantPermissions(spreadsheetId, editor);
        revokePermission(spreadsheetId, editor);
    }
}
