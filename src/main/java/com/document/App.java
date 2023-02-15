package com.document;

import tech.tablesaw.api.BooleanColumn;
import tech.tablesaw.api.DoubleColumn;
import tech.tablesaw.api.Table;
import tech.tablesaw.plotly.Plot;
import tech.tablesaw.plotly.api.HorizontalBarPlot;
import tech.tablesaw.plotly.components.Figure;
import tech.tablesaw.plotly.components.Layout;
import tech.tablesaw.plotly.traces.PieTrace;
import static tech.tablesaw.aggregate.AggregateFunctions.mean;

/**
 * Hello world!
 *
 */
public class App {
    public static void main(String[] args) {
        Table t = ExcelReader.readFileFromTerminal();
        processData(t);
        t = t.sortDescendingOn("globalScore");
        print(t);
        plots(t);
        try {
            ExcelReader.writeTable(t);
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public static void print(Table t) {
        System.out.println(
                Colors.BLUE_BACKGROUND_BRIGHT + "Standard Deviation: "
                        + t.doubleColumn("globalScore").standardDeviation() + Colors.RESET);
    }

    public static void processData(Table t) {
        DoubleColumn scores = DoubleColumn.create("globalScore");
        BooleanColumn approvers = BooleanColumn.create("approved");

        t.forEach(row -> {
            double globalScore = row.getNumber("score1") * 0.2 +
                    row.getNumber("score2") * 0.3 +
                    row.getNumber("score3") * 0.5;
            scores.append(globalScore);
            approvers.append(globalScore >= 3);
        });
        t.addColumns(scores, approvers);

    }

    public static void plots(Table t) {
        Table t2 = t.countBy(t.categoricalColumn("approved"));
        PieTrace trace = PieTrace.builder(t2.categoricalColumn("approved"),
                t2.numberColumn("Count")).build();
        Layout layout = Layout.builder().title("Total students by approved").build();

        Plot.show(new Figure(layout, trace));
        Table means = t.summarize("globalScore", mean).by("approved");

        // Plot
        Plot.show(
                HorizontalBarPlot.create(
                        "Means by approved", // plot title
                        means, // table
                        "approved", // grouping column name
                        "mean [globalScore]")); // numeric column name
    }
}
