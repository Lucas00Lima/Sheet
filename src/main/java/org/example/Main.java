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
        String BANCO = JOptionPane.showInputDialog("Digite o banco do cliente db***");
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
            String url = "jdbc:mysql://localhost:3306/" + BANCO;
            String userName = "root";

            String password = JOptionPane.showInputDialog("Insira a da Soma");
            Connection connection = DriverManager.getConnection(url, userName, password);

            //UPDATE no banco com os dados da planilha
            //Loop na planilha
            int rowIndex;
            for (rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (isRowEmpty(row)) {
                    break;
                } else {
                    String id = dataFormatter.formatCellValue(row.getCell(2));
                    String ncm = dataFormatter.formatCellValue(row.getCell(3));
                    String cfop = dataFormatter.formatCellValue(row.getCell(4));
                    String cest = dataFormatter.formatCellValue(row.getCell(12));
                    String cst = dataFormatter.formatCellValue(row.getCell(9));
                    String icms = dataFormatter.formatCellValue(row.getCell(5));
                    String pisCod = dataFormatter.formatCellValue(row.getCell(10));
                    String pisA = dataFormatter.formatCellValue(row.getCell(6));
                    String cofinsCod = dataFormatter.formatCellValue(row.getCell(11));
                    String cofinsA = dataFormatter.formatCellValue(row.getCell(7));
                    PreparedStatement preparedStatement = connection.prepareStatement("UPDATE product " +
                            "SET ncm = " + ncm
                            + ", cfop = " + cfop
                            + ", tax4_code = " + cest
                            + ", tax1_code = " + cst
                            + ", tax1 = " + icms
                            + ", tax2_code = " + pisCod
                            + ", tax2 = " + pisA
                            + ", tax3_code = " + cofinsCod
                            + ", tax3 = " + cofinsA
                            + " WHERE internal_code = " + id);
                    preparedStatement.execute();
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