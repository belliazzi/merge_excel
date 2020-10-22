package com.mkyong.hashing;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.stream.Stream;

public class App {

	public static void main(String[] args) {

		System.out.print("START ");
		System.out.println(" digita percorso dove sta la cartella dei file excel ");
		System.out.println(" esempio C:\\Users\\francesco\\fileexcell");
		Scanner sc = new Scanner(System.in); // System.in is a standard input stream
		System.out.print(" scrivi percorso alla cartella : ");
		String str = sc.nextLine(); // reads string
		System.out.println(" percorso scelto : " + str);
		System.out.println(" rimuovo file  scelto : risultatoMerge.xls");
		
        try {
        	if (Files.exists(Paths.get(str + "\\risultatoMerge.xls"))) {
        		 Files.delete(Paths.get(str + "\\risultatoMerge.xls"));
        	}
           
        } catch (IOException e) {
            e.printStackTrace();
        }
		CordinateDaLeggere cordinate = new CordinateDaLeggere();
		CellaDaLeggere societa = new CellaDaLeggere(3, 0, org.apache.poi.ss.usermodel.CellType.FORMULA, "societa");
		CellaDaLeggere data = new CellaDaLeggere(8, 1, org.apache.poi.ss.usermodel.CellType.STRING, "data");
		CellaDaLeggere centroDiCosto = new CellaDaLeggere(10, 7, org.apache.poi.ss.usermodel.CellType.FORMULA,
				"centroDiCosto");
		CellaDaLeggere dipendente = new CellaDaLeggere(8, 7, org.apache.poi.ss.usermodel.CellType.STRING, "dipendente");
		CellaDaLeggere CF = new CellaDaLeggere(9, 7, org.apache.poi.ss.usermodel.CellType.STRING, "CF");
		cordinate.getList().add(societa);
		cordinate.getList().add(data);
		cordinate.getList().add(centroDiCosto);
		cordinate.getList().add(dipendente);
		cordinate.getList().add(CF);

		ApachePOIExcelRead apr = new ApachePOIExcelRead();
		ApachePOIExcelWrite writers = new ApachePOIExcelWrite();

		File folder = new File(str);
		File[] listOfFiles = folder.listFiles();
		ArrayList<RowForExcelMerge> list = new ArrayList<RowForExcelMerge>();
		for (File file : listOfFiles) {
			if (file.isFile()) {
				System.out.println("elaboro file :" + str + "\\" + file.getName());
				list.addAll(apr.read(cordinate, str + "\\" + file.getName()));
			}
			writers.write(list, str + "\\risultatoMerge.xls");
		}

	}

}