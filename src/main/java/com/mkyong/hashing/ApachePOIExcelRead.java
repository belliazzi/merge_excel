package com.mkyong.hashing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePOIExcelRead {

	//private final String FILE_NAME = "C:\\Users\\damiano\\Downloads\\CON_BERNARDO_08-09-2020.xlsm";

	public ArrayList<RowForExcelMerge> read(CordinateDaLeggere cordinate, String filename) {
		ArrayList<RowForExcelMerge> res = new ArrayList<RowForExcelMerge>();
		try {
			
			FileInputStream excelFile = new FileInputStream(new File(filename));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			HashMap<String,String> hash = new HashMap<String,String>();
			for(CellaDaLeggere elem : cordinate.getList()) {
				Row currentRow = datatypeSheet.getRow(elem.getRow());				
				hash.put(elem.getNome(), ExcelHandler.getCellValue(currentRow.getCell(elem.getCellNum())));
			}
			
			
			int i = 17; //dalal riga 18-1 ci sono i record
			while (!ExcelHandler.isRowEmpty(datatypeSheet.getRow(i))) {
				RowForExcelMerge rowforExcelMerge = new RowForExcelMerge();
				rowforExcelMerge.setCentroDiCosto(hash.get("centroDiCosto"));
				rowforExcelMerge.setSocieta(hash.get("societa"));
				rowforExcelMerge.setDipendente(hash.get("dipendente"));
				rowforExcelMerge.setCF(hash.get("CF"));
				rowforExcelMerge.setData(hash.get("data"));
				Row currentRow = datatypeSheet.getRow(i);
				if (ExcelHandler.getCellValue(currentRow.getCell(2)).trim().equals("")) {
					System.out.println("la riga i è vuota" + i);
					break;
				}
				rowforExcelMerge.setCodice(ExcelHandler.getCellValue(currentRow.getCell(0)));
				rowforExcelMerge.setDesrizione(ExcelHandler.getCellValue(currentRow.getCell(2)));
				rowforExcelMerge.setQuantita(ExcelHandler.getCellValue(currentRow.getCell(9)));
				System.out.println("riga :" + i);
				i++;
				res.add(rowforExcelMerge);
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return res;

	}
}