package readxlsx;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.math.BigDecimal;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ReadXLSX {
    
    public static CurrencyWriter cw = new CurrencyWriter();
    public static Integer line = 2;
    public static String path = "D:\\Cutrale\\0.Wheslley\\Archives\\08112018\\Real_Safra\\";
    public static String archiveRead = "Real_Safra.xls";
    public static String subNameArchive = "REAL_SAFRA";
    
    public static void main(String[] args) throws IOException, BiffException {

        File arquivo = new File(path + archiveRead);

        Workbook workbook = Workbook.getWorkbook(arquivo);

        Sheet sheet = workbook.getSheet(0);
        int linhas = sheet.getRows();

        String nameArchive = null;
        String nameArchiveOld = null;
        String key = null;
        String data = "";

        for (int i = 1; i < linhas; i++) {

            if (key == null) {
                key = sheet.getCell(0, i).getContents();
                nameArchive = getNameArchive(sheet, i);
                nameArchiveOld = nameArchive;
            } else {
                nameArchive = getNameArchive(sheet, i);
            }

            if (!sheet.getCell(0, i).getContents().equalsIgnoreCase(key)) {
                generateTxt(nameArchiveOld, subNameArchive, data);
                key = null;
                data = "";
            }

            data += lineXlx(sheet, i);
        }

        workbook.close();

    }

    public static String getNameArchive(Sheet sheet, Integer i) {
        if (sheet.getCell(1, i).getContents().length() > 20) {
            return sheet.getCell(0, i).getContents() + "_" + sheet.getCell(1, i).getContents().substring(0, 20);
        } else {
            return sheet.getCell(0, i).getContents() + "_" + sheet.getCell(1, i).getContents();
        }
    }

    public static String lineXlx(Sheet sheet, Integer i) {

        System.out.println("Linha: " + line);
        line++;
        
        if(sheet.getCell(0, i).getContents().replaceAll("%", "").equals("")){
            return "";
        }

        return sheet.getCell(0, i).getContents().replaceAll("%", "") + ";"
                + sheet.getCell(1, i).getContents().replaceAll("%", "") + ";"
                + sheet.getCell(2, i).getContents().replaceAll("%", "") + ";"
                + sheet.getCell(3, i).getContents().replaceAll("%", "") + ";"
                + sheet.getCell(4, i).getContents().replaceAll("%", "") + ";"
                + sheet.getCell(5, i).getContents().replaceAll("%", "") + ";"
                + sheet.getCell(6, i).getContents().replaceAll("%", "") + ";"
                + sheet.getCell(7, i).getContents().replaceAll("%", "") + ";"
                + cw.write(new BigDecimal(sheet.getCell(7, i).getContents().replaceAll("%", ""))) + ";\r\n";

    }

    public static void generateTxt(String nameArchive, String typeArchive, String data) throws IOException {

        FileWriter arq = new FileWriter(path + typeArchive + "_" + nameArchive + ".txt");

        PrintWriter gravarArq = new PrintWriter(arq);

        gravarArq.printf(data);

        arq.close();

    }

}
