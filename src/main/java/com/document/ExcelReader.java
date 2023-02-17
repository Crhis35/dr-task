package com.document;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import tech.tablesaw.api.DoubleColumn;
import tech.tablesaw.api.StringColumn;
import tech.tablesaw.api.Table;
import tech.tablesaw.io.xlsx.*;
import tech.tablesaw.selection.Selection;
import tech.tablesaw.io.csv.*;

public class ExcelReader {
    static String file = "";

    private static Table readFile(String path) throws IOException {
        XlsxReader reader = new XlsxReader();
        XlsxReadOptions options = XlsxReadOptions.builder(path).build();

        return reader.read(options);
    }

    public static void writeTable(Table t) throws IOException {
        String path = ExcelReader.file.split("[.]")[0] + System.nanoTime() + ".csv";
        CsvWriter writer = new CsvWriter();
        CsvWriteOptions options = CsvWriteOptions.builder(path).build();

        writer.write(t, options);
        System.out.print(Colors.PURPLE_BACKGROUND_BRIGHT + "File generated on path: " + path + Colors.RESET);

    }

    public static Table readFileFromTerminal() {
        Scanner s = new Scanner(System.in);
        Table result = Table.create();
        System.out.print(Colors.GREEN + "Please enter a path from excel (xlsx): " + Colors.RESET);
        Map<String, String> validColumns = Map.of(
                "id", "STRING,INTEGER",
                "name", "STRING",
                "score1", "INTEGER,DOUBLE",
                "score2", "INTEGER,DOUBLE",
                "score3", "INTEGER,DOUBLE");
        try {
            while (true) {
                String input = s.nextLine();
                try {
                    result = readFile(input);
                    List<String> columnNames = result.columnNames();

                    for (Map.Entry<String, String> entry : validColumns.entrySet()) {
                        String key = entry.getKey();
                        List<String> types = Arrays.asList(entry.getValue().split(","));
                        String currentType = result.column(key).type().toString();
                        if (!columnNames.contains(key)) {
                            throw new Exception("Invalid column " + key);
                        }
                        if (!types.contains(currentType)) {
                            throw new Exception(
                                    String.format("Error type on column: %s \nExpected types: %s \nReceived types: %s",
                                            key + Colors.RESET,
                                            Colors.GREEN_BACKGROUND + types + Colors.RESET, Colors.RED_BACKGROUND_BRIGHT
                                                    + currentType + Colors.RESET));
                        }
                    }
                    List<String> scoreCols = Arrays.asList("score1", "score2", "score3");
                    for (String col : scoreCols) {

                        StringColumn parser = result.column(col).asStringColumn();

                        for (String a : parser) {
                            if (Float.parseFloat(a) < 0 || Float.parseFloat(a) > 5) {
                                throw new Exception(
                                        String.format("Error value on column: %s \nReceived values: %s",
                                                col + Colors.RESET,
                                                Colors.RED + a + Colors.RESET));
                            }
                        }

                    }

                    ExcelReader.file = input;
                    break;
                } catch (Exception nfe) {
                    System.err.print(Colors.RED);
                    System.err.println(nfe);
                    System.err.println("Error: " + nfe.getMessage() + Colors.RESET);
                    System.out.print(Colors.BLUE + "Please enter a valid path from excel (xlsx): " + Colors.RESET);
                }

            }

        } catch (Exception e) {
            s.close();
        }
        return result;
    }

}
