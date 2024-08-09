package com.shiyaam;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;

@Command(name = "", description = "Runs all the query in the excel")
public class MainCommand implements Runnable {
    private static final String JDBC_URL = "jdbc:postgresql://localhost:5432/excelqueryrunner";
    private static final String USERNAME = "postgres";
    private static final String PASSWORD = "postgres";

    private static final Logger LOGGER = LoggerFactory.getLogger(MainCommand.class);

    

    public static void main(String[] args) {
        new CommandLine(new MainCommand()).execute(args);
    }

    @Override
    public void run() {
        try (Connection connection = DriverManager.getConnection(JDBC_URL, USERNAME, PASSWORD)) {
            LOGGER.info("SQL is running");
            File file = new File("./test.xlsx"); // here is where your excel is hardcoded
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            var len = sheet.getLastRowNum();
            LOGGER.info("Last row num: {}", len);
            for (int i = 1; i <= len; i++) {
                Row row = sheet.getRow(i);
                String query = row.getCell(0).getStringCellValue();
                double count = row.getCell(2).getNumericCellValue();
                LOGGER.info("Row 1: Query: {}, Count: {}\n", query, count);
                Statement st = connection.createStatement();
                String modifiedQuery = query.replaceAll("count\\(\\*\\)", "*");
                ResultSet modifiedrs = st.executeQuery(modifiedQuery);
                var list = new ArrayList<String>();
                while (modifiedrs.next()) {
                    LOGGER.info("{}", modifiedrs.getString("order_id"));
                    list.add(modifiedrs.getString("order_id"));
                }
                StringBuilder sb = new StringBuilder();
                if (list.size() == 1) {
                    sb.append("\"" + list.get(0) + "\"");
                } else {
                    for (int k = 0; k < list.size() - 1; k++) {
                        sb.append("\"" + list.get(k) + "\" OR ");
                    }
                    sb.append("\"" + list.get((list.size() - 1)) + "\"");
                }
                String aggregatedQuery = sb.toString();
                LOGGER.info("Aggregated query: {}", aggregatedQuery);
                var aggregatedQueryCell = row.createCell(4);
                aggregatedQueryCell.setCellValue(aggregatedQuery);
                var aggregatedQueryCount = row.createCell(5);
                aggregatedQueryCount.setCellValue(list.size());
                FileOutputStream fos = new FileOutputStream(file);
                workbook.write(fos);
            }
            LOGGER.info("Got the sheets");
        } catch (SQLException e) {
            LOGGER.error("SQL State: {}\nError Message: {}", e.getSQLState(), e.getMessage());
        } catch (IOException e) {
            LOGGER.error("Exception while trying to get xlsx file: {}", e.getMessage());
        }
    }
}