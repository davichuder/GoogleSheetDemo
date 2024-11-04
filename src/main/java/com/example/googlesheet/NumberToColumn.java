package com.example.googlesheet;

public class NumberToColumn {
    public static void main(String[] args) {
        int[] initialPoint = { 27, 1 };
        System.out.println(convertToColumnName(initialPoint[0]) + initialPoint[1]);
    }

    public static String convertToColumnName(int number) {
        StringBuilder columnName = new StringBuilder();
        while (number > 0) {
            int remainder = (number - 1) % 26;
            columnName.insert(0, (char) (remainder + 'A'));
            number = (number - 1) / 26;
        }
        return columnName.toString();
    }
}
