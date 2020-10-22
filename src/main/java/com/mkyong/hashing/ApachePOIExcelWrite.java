package com.mkyong.hashing;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ApachePOIExcelWrite {

   // private static final String FILE_NAME = "C:\\Users\\damiano\\Downloads\\risultatoMerge.xls";

    public  void write(ArrayList<RowForExcelMerge> list ,String pathDestination) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
     
        int rowNum = 0;
        System.out.println("Creating excel");
        //intesatzione 
        int colNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cellSocietaTitolo = row.createCell(colNum++);
        cellSocietaTitolo.setCellValue("SOCIETA'");
        Cell cellDataTitolo = row.createCell(colNum++);
        cellDataTitolo.setCellValue("DATA");
        Cell cellCentroDicostoTitolo = row.createCell(colNum++);
        cellCentroDicostoTitolo.setCellValue("CENTRO DI COSTO");
        Cell cellDipendenteTitolo = row.createCell(colNum++);
        cellDipendenteTitolo.setCellValue("DIPENDENTE");
        Cell cellCFTitolo = row.createCell(colNum++);
        cellCFTitolo.setCellValue("CODCIE FISCALE");
        Cell cellCodiceTitolo = row.createCell(colNum++);
        cellCodiceTitolo.setCellValue("CODICE");
        Cell cellDescrizioneTitolo = row.createCell(colNum++);
        cellDescrizioneTitolo.setCellValue("DESCRIZIONE");
        Cell cellQuantitaTitolo = row.createCell(colNum++);
        cellQuantitaTitolo.setCellValue("QUANTITA'");
        
        for (RowForExcelMerge elem : list) {
            row = sheet.createRow(rowNum++);
            colNum = 0;
            Cell cellSocieta = row.createCell(colNum++);
            cellSocieta.setCellValue((String) elem.getSocieta());
            Cell cellData = row.createCell(colNum++);
            cellData.setCellValue(elem.getData());
            Cell cellCentroDicosto = row.createCell(colNum++);
            cellCentroDicosto.setCellValue((String) elem.getCentroDiCosto());
            Cell cellDipendente = row.createCell(colNum++);
            cellDipendente.setCellValue((String) elem.getDipendente());
            Cell cellCF = row.createCell(colNum++);
            cellCF.setCellValue((String) elem.getCF());
            Cell cellCodice = row.createCell(colNum++);
            cellCodice.setCellValue((String) elem.getCodice());
            Cell cellDescrizione = row.createCell(colNum++);
            cellDescrizione.setCellValue((String) elem.getDesrizione());
            Cell cellQuantita = row.createCell(colNum++);
            cellQuantita.setCellValue((Double) Double.valueOf(elem.getQuantita()));
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(pathDestination);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }
}
