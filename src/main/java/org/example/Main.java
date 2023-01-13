package org.example;

import com.github.javafaker.Faker;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class Main {
    public static void main(String[] args) {
        String relativePath = "imePrezime.xlsx";
        try {
            readFile(relativePath);
            writeFile(relativePath);
            readPerson(relativePath);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    public static void readFile(String relativePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(relativePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
      /*  CellReference cellReferenceA1 = new CellReference("A1");
        XSSFRow row = sheet.getRow(cellReferenceA1.getRow());
        XSSFCell cell = row.getCell(cellReferenceA1.getCol()); */

        for (int i = 0; i < 5; i++) {
            XSSFRow row1 = sheet.getRow(i);
            for (int j = 0; j < 2; j++) {
                XSSFCell cell1 = row1.getCell(j);
                System.out.print(cell1.getStringCellValue() + " ");
            }
            System.out.println();
        }
    }

    public static void writeFile(String filename) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        for (int i = 0; i < 5; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.getCell(0);
            XSSFCell cell1 = row.getCell(1);
            XSSFCell cell2 = row.createCell(2);
            XSSFCell cell3 = row.createCell(3);
            cell2.setCellValue(String.valueOf(cell));
            cell3.setCellValue(String.valueOf(cell1));
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filename);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public static void addPerson(String filename) throws IOException {
        Faker faker = new Faker();
        String name1 = faker.name().firstName();
        String name2 = faker.name().firstName();
        String name3 = faker.name().firstName();
        String name4 = faker.name().firstName();
        String name5 = faker.name().firstName();
        String lastName1 = faker.name().lastName();
        String lastName2 = faker.name().lastName();
        String lastName3 = faker.name().lastName();
        String lastName4 = faker.name().lastName();
        String lastName5 = faker.name().lastName();
        ArrayList<String> firstname = new ArrayList<>();
        firstname.add(name1);
        firstname.add(name2);
        firstname.add(name3);
        firstname.add(name4);
        firstname.add(name5);
        ArrayList<String> lastname = new ArrayList<>();
        lastname.add(lastName1);
        lastname.add(lastName2);
        lastname.add(lastName3);
        lastname.add(lastName4);
        lastname.add(lastName5);
        FileInputStream fileInputStream = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        for (int i = 0; i < 5; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell2 = row.createCell(4);
            XSSFCell cell3 = row.createCell(5);
            cell2.setCellValue(firstname.get(i));
            cell3.setCellValue(lastname.get(i));
        }
        FileOutputStream fileOutputStream = new FileOutputStream(filename);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public static void readPerson(String relativePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(relativePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        for (int i = 0; i < 5; i++) {
            XSSFRow row1 = sheet.getRow(i);
            for (int j = 4; j < 6; j++) {
                XSSFCell cell1 = row1.getCell(j);
                System.out.print(cell1.getStringCellValue() + " ");
            }
            System.out.println();
        }
    }
}