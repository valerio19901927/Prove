package org.example;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.UUID;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToInsertQueries {

    public static void main(String[] args) {
        String excelFilePath = "src/main/resources/Book1.xlsx";
        String outputFilePath = "src/main/resources/insert.txt";
        try {
            generateInsertQueriesFromExcel(excelFilePath, outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void generateInsertQueriesFromExcel(String excelFilePath, String outputFilePath)
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
            if (headerRow == null) {
                throw new IllegalStateException("Header row not found in the Excel sheet.");
            }
            int colCount = headerRow.getPhysicalNumberOfCells();

            int tipoIndex = getColumnIndex(headerRow, "TIPO");
            int faseIndex = getColumnIndex(headerRow, "FASE");
            int autoritaIndex = getColumnIndex(headerRow, "AUTORITA");

            StringBuilder insertQuery = new StringBuilder("INSERT INTO cms_anas_be.gradi_giudizio (ID, ");

            // Aggiunge i nomi delle colonne dinamicamente
            for (int i = 0; i < colCount; i++) {
                insertQuery.append(headerRow.getCell(i).getStringCellValue() + ", ");
            }
            // Aggiunge l'ultima parte della query dopo l'elenco delle colonne
            insertQuery.setLength(insertQuery.length() - 2);
            insertQuery.append(") VALUES ");

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                Cell tipoCell = row.getCell(tipoIndex);
                Cell faseCell = row.getCell(faseIndex);
                Cell autoritaCell = row.getCell(autoritaIndex);

                selectStatement.setObject(1, getCellValueAsString(tipoCell));
                selectStatement.setObject(2, getCellValueAsString(faseCell));
                selectStatement.setObject(3, getCellValueAsString(autoritaCell));

                ResultSet resultSet = selectStatement.executeQuery();
                resultSet.next();
                int rowCount = resultSet.getInt("count");

                resultSet.close();

                // Se la tripletta non è presente, aggiungi i valori alla query
                if (rowCount == 0) {
                    insertQuery.append("(");

                    // Aggiunge l'ID alfanumerico generato casualmente
                    String randomId = generateRandomId();
                    insertQuery.append("'" + randomId + "', ");

                    // Aggiunge i valori delle colonne dinamicamente
                    DataFormatter dataFormatter = new DataFormatter();
                    for (int i = 0; i < colCount; i++) {
                        Cell dataCell = row.getCell(i);

                        // Aggiungi il valore della cella solo se non è nullo
                        if (dataCell != null) {
                            String cellValue;

                            if (dataCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                cellValue = dataFormatter.formatCellValue(dataCell);
                            } else {
                                cellValue = getCellValueAsString(dataCell);
                            }

                            // Aggiungi il valore della cella
                            insertQuery.append("'" + cellValue + "', ");
                        }
                    }

                    // Rimuovi l'ultima virgola e spazio
                    insertQuery.setLength(insertQuery.length() - 2);

                    // Aggiunge la parentesi chiusa
                    insertQuery.append("), ");
                }
            }

            // Rimuovi l'ultima virgola e spazio
            insertQuery.setLength(insertQuery.length() - 2);

            writer.write(insertQuery.toString());
            writer.newLine();
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

    private static String getCellValueAsString(Cell cell) {
        if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            return ""; // Restituisci una stringa vuota se la cella è vuota
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                return dateFormat.format(cell.getDateCellValue());
            } else {
                return NumberToTextConverter.toText(cell.getNumericCellValue());
            }
        } else {
            return cell.getStringCellValue();
        }
    }
}
