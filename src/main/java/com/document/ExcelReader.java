package com.document;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import tech.tablesaw.api.DoubleColumn;
import tech.tablesaw.api.StringColumn;
import tech.tablesaw.api.Table;
import tech.tablesaw.io.xlsx.*;
import tech.tablesaw.io.csv.*;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import static tech.tablesaw.aggregate.AggregateFunctions.mean;

public class ExcelReader {
    static String file = "";
    static int tableSize = 0;
    public static Map<String, List<String>> errors = Map.of(
            "columns", new ArrayList<String>(),
            "numbers", new ArrayList<String>(),
            "fixed", new ArrayList<String>(),
            "types", new ArrayList<String>());

    static Map<String, String> validColumns = Map.of(
            "id", "STRING,INTEGER",
            "name", "STRING",
            "score1", "INTEGER,DOUBLE",
            "score2", "INTEGER,DOUBLE",
            "score3", "INTEGER,DOUBLE");

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

    public static String isNumeric(String str) {
        String current = str;

        if (str.contains(",")) {
            current = current.replace(",", ".");
            errors.get("fixed").add(String.format("Received % -> fixed as %", str, current));
        }

        if (current.matches("-?\\d+(\\.\\d+)?")) {
            return current;
        }
        return "false";
    }

    public static Table readFileFromTerminal2() {
        Scanner s = new Scanner(System.in);
        Table result = Table.create();
        System.out.print(Colors.GREEN + "Please enter a path from excel (xlsx): " + Colors.RESET);

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

    public static Table readFileFromTerminal() {
        Scanner s = new Scanner(System.in);
        Table result = Table.create();
        System.out.print(Colors.GREEN + "Please enter a path from excel (xlsx): " + Colors.RESET);

        try {
            while (true) {
                String input = s.nextLine();
                try {
                    result = readFile(input);
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
        Table t = Table.create();
        for (Map.Entry<String, String> entry : validColumns.entrySet()) {
            try {
                String key = entry.getKey();
                StringColumn currentCol = result.column(key).asStringColumn();
                tableSize = currentCol.size();
                List<String> types = Arrays.asList(entry.getValue().split(","));
                if (types.contains("STRING")) {
                    t.addColumns(currentCol.setName(key));
                } else {
                    DoubleColumn col = DoubleColumn.create(key);
                    currentCol.forEach(row -> {
                        String value = isNumeric(row);
                        if (value == "false") {
                            errors.get("types")
                                    .add(String.format("Get invalid value on %s, received -> %s. Set to 0", key, row));
                            col.append(0);
                        } else {
                            double score = Double.parseDouble(value);
                            if (score > 5 || score < 0) {
                                col.append(0);
                                errors.get("numbers")
                                        .add(String.format("Get invalid value on %s, received -> %s. Set to 0", key,
                                                row));
                            } else {
                                col.append(score);
                            }

                        }
                    });
                    t.addColumns(col);
                }
            } catch (Exception nfe) {
                errors.get("columns").add(nfe.getMessage());
                System.err.println(nfe);
            }
        }
        return t;
    }

    public static Object parseType(String type, String value) {
        switch (type) {
            case "DOUBLE":
                return Double.parseDouble(value);
            case "INTEGER":
                return Integer.parseInt(value);
            case "BOOLEAN":
                return Boolean.parseBoolean(value);
            default:
                return value;
        }

    }

    public static void writeSheet(Worksheet sheet, Map<String, String> map, Table t) {
        int counter = 1;

        for (Map.Entry<String, String> entry : map.entrySet()) {
            String key = entry.getKey();
            List<String> types = Arrays.asList(entry.getValue().split(","));
            StringColumn parser = t.column(key).asStringColumn();
            sheet.getCells().get(types.get(0) +
                    counter).putValue(key);
            int localCounter = counter + 1;

            for (String val : parser) {
                sheet.getCells().get(types.get(0) +
                        localCounter).putValue(parseType(types.get(1), val));
                localCounter += 1;
            }

            counter = 1;

        }
    }

    public static void writeLogs(Worksheet sheet) {
        int counter = 1;
        Map<String, String> idx = Map.of(
                "columns", "A",
                "numbers", "B",
                "fixed", "C",
                "types", "D");
        for (Map.Entry<String, List<String>> entry : errors.entrySet()) {
            String key = entry.getKey();
            List<String> types = entry.getValue();
            sheet.getCells().get(idx.get(key) +
                    counter).putValue(key);
            int localCounter = counter + 1;

            for (String val : types) {
                sheet.getCells().get(idx.get(key) +
                        localCounter).putValue(val);
                localCounter += 1;
            }

            counter = 1;

        }
    }

    public static void saveFile(Table t) throws Exception {

        Workbook workbook = new Workbook();
        workbook.getWorksheets().add("logs");
        workbook.getWorksheets().add("plots");
        Worksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.setName("students");

        Map<String, String> sheet1 = Map.of(
                "id", "A,STRING",
                "name", "B,STRING",
                "score1", "C,DOUBLE",
                "score2", "D,DOUBLE",
                "score3", "E,DOUBLE",
                "globalScore", "F,DOUBLE",
                "approved", "G,BOOLEAN");

        String path = ExcelReader.file.split("[.]")[0] + System.nanoTime() + ".xlsx";
        int chartIndex2 = worksheet.getCharts().add(ChartType.HISTOGRAM, 5, 7, 15, 12);
        // Access the instance of the newly added chart
        Chart chart2 = worksheet.getCharts().get(chartIndex2);

        // Set chart data source as the range "A1:C4"
        chart2.setChartDataRange("F2:F" + tableSize, true);
        writeSheet(worksheet, sheet1, t);

        Worksheet worksheet3 = workbook.getWorksheets().get("logs");
        writeLogs(worksheet3);

        Worksheet worksheet2 = workbook.getWorksheets().get("plots");
        Map<String, String> sheet2 = Map.of(
                "approved", "A,BOOLEAN",
                "count", "B,INTEGER");

        worksheet2.getCells().get("A4").putValue("Standard Deviation on globalScore");
        worksheet2.getCells().get("B4").putValue(t.doubleColumn("globalScore").standardDeviation());
        Table t2 = t.countBy(t.categoricalColumn("approved"));
        writeSheet(worksheet2, sheet2, t2);
        // ChartType.HISTOGRAM
        int chartIndex = worksheet2.getCharts().add(ChartType.PIE, 5, 7, 15, 12);
        // Access the instance of the newly added chart
        Chart chart = worksheet2.getCharts().get(chartIndex);

        // Set chart data source as the range "A1:C4"
        chart.setChartDataRange("A1:B3", true);

        workbook.save(path, SaveFormat.XLSX);
    }
}
