import com.github.javafaker.Faker;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {



        try {
            readDataFromWorkbook("workbook1.xlsx");
        } catch (FileNotFoundException exception) {
            System.out.println("Nevalidna putanja.");
        } catch (IOException exception) {
            System.out.println("nevalidan excel fajl.");

        }


        try {
            writeDataInWorkbook("workbook1.xlsx");
        } catch (FileNotFoundException exception) {
            System.out.println("Nevalidna putanja!");
        } catch (IOException exception) {
            System.out.println("Nevalidan excel fajl!");
        }


    }

    public static void readDataFromWorkbook (String relativePath) throws IOException {

        FileInputStream inputStream = new FileInputStream("workbook1.xlsx");
        XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook1.getSheet("imenaiprezimena");

        for (int i = 0; i < 5; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cellIme = row.getCell(0);
            XSSFCell cellPrezime = row.getCell(1);

            String ime = cellIme.getStringCellValue();
            String prezime = cellPrezime.getStringCellValue();

            System.out.println(ime + " " + prezime);
        }
    }


    public static void writeDataInWorkbook (String relativePath) throws IOException {  //FileNotFoundException je podklasa od IOException

        Faker faker = new Faker();
        FileInputStream inputStream = new FileInputStream(relativePath);
        XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook1.getSheet("imenaiprezimena");

        for (int i = 5; i < 10; i++) {
            XSSFRow row = sheet.createRow(i);
            XSSFCell cellIme = row.createCell(0);
            cellIme.setCellValue(faker.name().firstName());
            XSSFCell cellPrezime = row.createCell(1);
            cellPrezime.setCellValue(faker.name().lastName());

            System.out.println(cellIme + " " + cellPrezime);

        }

        FileOutputStream outputStream = new FileOutputStream("workbook1");
        workbook1.write(outputStream);
        outputStream.close();

    }


}