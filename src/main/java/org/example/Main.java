package org.example;

import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class Main {
    public static void main(String[] args) throws IOException, SQLException {
        //Acessar a planilha
//        String BANCO = JOptionPane.showInputDialog("Digite o banco do cliente db***");
        JFileChooser fileChooser = new JFileChooser();
        int result = fileChooser.showOpenDialog(null);
        String filePath;
        if (result == JFileChooser.APPROVE_OPTION) {
            filePath = fileChooser.getSelectedFile().getAbsolutePath();
            JOptionPane.showMessageDialog(null, "Arquivo Selecionado" + filePath);
            Workbook workbook = WorkbookFactory.create(new File(filePath));
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter dataFormatter = new DataFormatter();
            //Acessar o banco
            String url = "jdbc:mysql://localhost:3306/" + "db000";
            String userName = "root";

            String password = "@soma+";//JOptionPane.showInputDialog("Insira a da Soma");
            Connection connection = DriverManager.getConnection(url, userName, password);

            //UPDATE no banco com os dados da planilha
            //Loop na planilha
            int rowIndex;
            for (rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (isRowEmpty(row)) {
                    break;
                } else {
                    String id = dataFormatter.formatCellValue(row.getCell(0));
                    String ncm = dataFormatter.formatCellValue(row.getCell(9));
                    String cfop = dataFormatter.formatCellValue(row.getCell(10));
                    String cest = dataFormatter.formatCellValue(row.getCell(11));
                    String cst = dataFormatter.formatCellValue(row.getCell(12));

                    String icmsString = dataFormatter.formatCellValue(row.getCell(13));
                    icmsString = icmsString.replace(",", "");
                    icmsString = icmsString + "0" + "0";
                    int icms = Integer.parseInt(icmsString);

                    String pisCod = dataFormatter.formatCellValue(row.getCell(14));
                    pisCod = "0" + pisCod;
                    String pisAString = dataFormatter.formatCellValue(row.getCell(15));
                    pisAString = pisAString.replace(",", "");
                    pisAString = pisAString + "0";
                    int pisA = Integer.parseInt(pisAString);

                    String cofinsCod = dataFormatter.formatCellValue(row.getCell(16));
                    cofinsCod = "0" + cofinsCod;
                    String cofinsAString = dataFormatter.formatCellValue(row.getCell(17));
                    cofinsAString = cofinsAString.replace(",", "");
                    cofinsAString = cofinsAString + "0" + "0" + "0";
                    int cofinsA = Integer.parseInt(cofinsAString);

                    String sql = "UPDATE product " +
                            "SET ncm = ?, " +
                            "cfop = ?, " +
                            "tax4_code = ?, " +
                            "tax1_code = ?, " +
                            "tax1 = ?, " +
                            "tax2_code = ?, " +
                            "tax2 = ?, " +
                            "tax3_code = ?, " +
                            "tax3 = ? " +
                            "WHERE internal_code = ?";
                    PreparedStatement preparedStatement = connection.prepareStatement(sql);
                    preparedStatement.setString(1, ncm);
                    preparedStatement.setString(2, cfop);
                    preparedStatement.setString(3, cest);
                    preparedStatement.setString(4, cst);
                    preparedStatement.setInt(5, icms);
                    preparedStatement.setString(6, pisCod);
                    preparedStatement.setInt(7, pisA);
                    preparedStatement.setString(8, cofinsCod);
                    preparedStatement.setInt(9, cofinsA);
                    preparedStatement.setString(10, id);

                    preparedStatement.executeUpdate();
                    System.out.println(preparedStatement);
                }
            }
        }
    }
    public static boolean isRowEmpty(Row row) {
        if (row != null) {
            for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
                Cell cell = row.getCell(cellIndex);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    return false;
                }
            }
        }
        return true;
    }
}