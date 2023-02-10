package com.document;

import java.util.List;

import tech.tablesaw.api.DoubleColumn;
import tech.tablesaw.api.Table;

/**
 * Hello world!
 *
 */
public class App {
    public static void main(String[] args) {
        Table t = ExcelReader.readFileFromTerminal();
        System.out.println(t.print());

    }
}
