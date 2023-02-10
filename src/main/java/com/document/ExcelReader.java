package com.document;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

import tech.tablesaw.api.Table;
import tech.tablesaw.io.xlsx.*;

public class ExcelReader {
    private static Table readFile(String path) {
        XlsxReader reader = new XlsxReader();
        XlsxReadOptions options = XlsxReadOptions.builder(path).build();
        return reader.read(options);
    }

    public static Table readFileFromTerminal() {
        Scanner s = new Scanner(System.in);
        Table result = Table.create();
        ArrayList<String> validColumns = new ArrayList<>(Arrays.asList("id", "name", "score1", "score2", "score3"));
        System.out.println("Please enter a path from excel (xlsx)");
        try {
            while (true) {
                String input = s.nextLine();
                try {
                    result = readFile(input);
                    List<String> columnNames = result.columnNames(); // returns all column names
                    for (String value : validColumns) {
                        if (!columnNames.contains(value)) {
                            throw new Exception("Invalid column " + value);
                        }
                    }

                    break;
                } catch (Exception nfe) {
                    System.err.println(nfe);
                    System.err.println("Error: " + nfe.getMessage());
                }

            }

        } catch (Exception e) {
            s.close();
        }
        return result;
    }

}
