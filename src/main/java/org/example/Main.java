package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class Main {
    public static void main(String[] args) throws IOException, SQLException {
    //Acessar a planilha
        String filePath = "C:\\Users\\lukin\\OneDrive\\Área de Trabalho\\Pizzaria\\planilha.xls";
        FileInputStream fileInput = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(fileInput);
        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();
    //Acessar o banco
        String url = "jdbc:mysql://localhost:3306/db000";
        String userName = "root";
        String password = "@soma+";

        Connection connection = DriverManager.getConnection(url,userName,password);


    //UPDATE no banco com os dados da planilha
        //Loop na planilha
        int rowIndex = 2;
        for (rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            String id = dataFormatter.formatCellValue(row.getCell(0));
            String name = dataFormatter.formatCellValue(row.getCell(2));
            String ncm = dataFormatter.formatCellValue(row.getCell(9));
            String cfop = dataFormatter.formatCellValue(row.getCell(10));
            String cest = dataFormatter.formatCellValue(row.getCell(11));
            String cst = dataFormatter.formatCellValue(row.getCell(12));
            String icms = dataFormatter.formatCellValue(row.getCell(13));
            String pisCod = dataFormatter.formatCellValue(row.getCell(14));
            String pisA = dataFormatter.formatCellValue(row.getCell(15));
            String cofinsCod = dataFormatter.formatCellValue(row.getCell(16));
            String cofinsA = dataFormatter.formatCellValue(row.getCell(17));

            PreparedStatement preparedStatement = connection.prepareStatement("UPDATE product " +
                    "SET ncm = " + ncm
                    + ", cfop = " + cfop
                    + ", tax4_code = " + cest
                    + ", tax1_code = " + cst
                    + ", tax1 = " + icms
                    + ", tax2_code = " + pisCod
                    + ", tax3_code = " + cofinsCod
                    + ", tax2 = " + pisA
            + " WHERE internal_code = " + id);
            preparedStatement.execute();
            System.out.println(preparedStatement.toString());
        }
    }
}