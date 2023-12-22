package org.example;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.*;
import java.util.UUID;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToInsertQueries {

    public static void main(String[] args) {
        String excelFilePath = "src/main/resources/FileNew.xlsx";
        String outputFilePath = "src/main/resources/insert.txt";

        try {
            String[] columnNames = getColumnNamesFromDatabase("cms_anas_be", "gradi_giudizio");
            generateInsertQueriesFromExcel(excelFilePath, outputFilePath, columnNames);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String[] getColumnNamesFromDatabase(String schema, String tableName) throws Exception {
        String[] columnNames = null;
        try {
            String jdbcUrl = "jdbc:mysql://anasbe.test.dc1.ns:3306/anas";
            String username = "anas";
            String password = "4N4S_password";

            Connection connection = DriverManager.getConnection(jdbcUrl, username, password);

            String query = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS " +
                    "WHERE TABLE_SCHEMA = '" + schema + "' AND TABLE_NAME = '" + tableName + "'";

            Statement statement = connection.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);

            ResultSet resultSet = statement.executeQuery(query);

            int columnCount = 0;
            while (resultSet.next()) {
                columnCount++;
            }

            resultSet.beforeFirst();

            columnNames = new String[columnCount];
            int index = 0;
            while (resultSet.next()) {
                columnNames[index++] = resultSet.getString("COLUMN_NAME");
            }

            resultSet.close();
            statement.close();
            connection.close();
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }

        return columnNames;
    }

    private static int getColumnIndex(Row headerRow, String columnName) {
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalArgumentException("Column '" + columnName + "' not found in the header row.");
    }

    private static String generateRandomId() {
        return UUID.randomUUID().toString().replace("-", "").toUpperCase();
    }

    private static boolean containsColumn(String[] columnNames, String targetColumn) {
        for (String columnName : columnNames) {
            if (columnName.equalsIgnoreCase(targetColumn)) {
                return true;
            }
        }
        return false;
    }

    public static void generateInsertQueriesFromExcel(String excelFilePath, String outputFilePath, String[] columnNames)
            throws IOException, SQLException {
        FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
        XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
        BufferedWriter writer = new BufferedWriter(new FileWriter(outputFilePath));

        Connection connection = null;
        PreparedStatement selectStatement = null;

        try {
            String jdbcUrl = "jdbc:mysql://anasbe.test.dc1.ns:3306/anas";
            String username = "anas";
            String password = "4N4S_password";
            connection = DriverManager.getConnection(jdbcUrl, username, password);

            String selectQuery = "SELECT COUNT(*) AS count FROM cms_anas_be.gradi_giudizio WHERE TIPO = ? AND FASE = ? AND AUTORITA = ?";
            selectStatement = connection.prepareStatement(selectQuery);

            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);

            Row headerRow = sheet.getRow(0);
            int colCount = headerRow.getPhysicalNumberOfCells();

            int tipoIndex = getColumnIndex(headerRow, "TIPO");
            int faseIndex = getColumnIndex(headerRow, "FASE");
            int autoritaIndex = getColumnIndex(headerRow, "AUTORITA");

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                Cell tipoCell = row.getCell(tipoIndex);
                Cell faseCell = row.getCell(faseIndex);
                Cell autoritaCell = row.getCell(autoritaIndex);

                selectStatement.setObject(1, tipoCell.getStringCellValue());
                selectStatement.setObject(2, faseCell.getStringCellValue());
                selectStatement.setObject(3, autoritaCell.getStringCellValue());

                ResultSet resultSet = selectStatement.executeQuery();
                resultSet.next();
                int rowCount = resultSet.getInt("count");

                resultSet.close();

                // Se la tripletta non è presente, genera la query di inserimento
                if (rowCount == 0) {
                    StringBuilder insertQuery = new StringBuilder("INSERT INTO cms_anas_be.gradi_giudizio (");

                    // Aggiunge i nomi delle colonne dinamicamente
                    for (int i = 0; i < colCount; i++) {
                        // Seleziona il nome della colonna solo se è presente nell'intestazione
                        if (i < headerRow.getPhysicalNumberOfCells()) {
                            String columnName = columnNames[i];
                            insertQuery.append(columnName + ", ");
                        }
                    }

                    // Verifica se le colonne VALORE_MINIMO e FLAG_OLD sono presenti
                    if (containsColumn(columnNames, "VALORE_MINIMO")) {
                        insertQuery.append("VALORE_MINIMO, ");
                    }

                    if (containsColumn(columnNames, "FLAG_OLD")) {
                        insertQuery.append("FLAG_OLD, ");
                    }

                    // Rimuovi l'ultima virgola e spazio
                    insertQuery.setLength(insertQuery.length() - 2);

                    insertQuery.append(") VALUES (");

                    // Aggiunge l'ID alfanumerico generato casualmente
                    String randomId = generateRandomId();
                    insertQuery.append("'" + randomId + "', ");

                    // Aggiunge i valori delle colonne dinamicamente
                    for (int i = 0; i < colCount; i++) {
                        Cell dataCell = row.getCell(i);

                        Object cellValue;

                        if (dataCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            if (dataCell.getNumericCellValue() == Math.floor(dataCell.getNumericCellValue())) {
                                cellValue = (int) dataCell.getNumericCellValue();
                            } else {
                                cellValue = dataCell.getNumericCellValue();
                            }
                        } else {
                            cellValue = dataCell.getStringCellValue();
                        }

                        // Aggiungi il valore della cella
                        insertQuery.append("'" + cellValue + "', ");
                    }

                    // Rimuovi l'ultima virgola e spazio
                    insertQuery.setLength(insertQuery.length() - 2);

                    // Aggiunge il valore fisso 'N' per FLAG_OLD
                    insertQuery.append(",'N')");

                    // Aggiunge la parentesi chiusa e punto e virgola
                    insertQuery.append(";");

                    writer.write(insertQuery.toString());
                    writer.newLine();
                }
            }

            System.out.println("Query di inserimento generate con successo.");

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (selectStatement != null) {
                selectStatement.close();
            }
            if (connection != null) {
                connection.close();
            }
            if (workbook != null) {
                workbook.close();
            }
            if (excelFile != null) {
                excelFile.close();
            }
            if (writer != null) {
                writer.close();
            }
        }
    }
}
