package com.test.java.OfferIssue;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		System.out.println("Finding Offer Issue");
		readOfferExcelSheet();
	}

	public static void readOfferExcelSheet() {
		System.out.println("inside method readOfferExcelSheet ");
		// File Path # C:\yuvi\java\STS\OfferIssue\src\main\resources\InvalidOffers.xlsx
		try {
			String filePath = "C:\\yuvi\\java\\STS\\OfferIssue\\src\\main\\resources\\InvalidOffers.xlsx";
			FileInputStream file = new FileInputStream(new File(filePath));
			// We know there are 33 special characters. So we will use them.
			Pattern p = Pattern.compile("[ !\"#$%&'()*+,-./:;<=>?@\\[\\]^_`{|}~]");

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			int rowNumber = 0;
			List<String> offersList;
			while (rowIterator.hasNext()) {
				offersList = new ArrayList<String>();
				// System.out.println("Row Number # " + ++rowNumber);
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				int inmpactedOffers = 0;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					// cell.getRowIndex() + " "
					if (cell.getColumnIndex() <= 1)
						offersList.add(cell.toString());
					if (cell.getColumnIndex() > 1) {
						Matcher m = p.matcher(cell.toString());
						if (m.find()) {
							inmpactedOffers++;
							offersList.add(cell.toString());
							// System.out.print(cell.toString() + " ");
						}
					}
				}

				if (inmpactedOffers > 0) {
					// System.out.print(" " + inmpactedOffers + "\n");
					writeInvalidOffers(offersList, inmpactedOffers);
				}

			}
			file.close();
		} catch (Exception e) {
			System.out.println("Exception in  method readOfferExcelSheet()");
		}

	}

	public static void writeInvalidOffers(List<String> offersList, int numbers) {
		System.out.println("Finding Invalid Offers " + numbers);
		try {
			List<String> listOfOffers = new ArrayList<String>();
			listOfOffers.addAll(offersList);
			listOfOffers.forEach(element -> System.out.print(element + " # "));
			System.out.println("");
			// Blank workbook
			XSSFWorkbook workbook = new XSSFWorkbook();
			// Create a blank sheet
			XSSFSheet sheet = workbook.createSheet("InValid Offers Data");
			String filePath = "C:\\yuvi\\java\\STS\\OfferIssue\\src\\main\\resources\\output.xlsx";
			// This data needs to be written (Object[])
			Map<String, Object[]> data = new TreeMap<String, Object[]>();
			data.put("1", new Object[] { "ID", "NAME", "LASTNAME" });
			data.put("2", new Object[] { 1, "Amit", "Shukla" });
			data.put("3", new Object[] { 2, "Lokesh", "Gupta" });
			data.put("4", new Object[] { 3, "John", "Adwards" });
			data.put("5", new Object[] { 4, "Brian", "Schultz" });

			// Iterate over data and write to sheet
			Set<String> keyset = data.keySet();
			int rownum = 0;
			for (String key : keyset) {
				Row row = sheet.createRow(rownum++);
				Object[] objArr = data.get(key);
				int cellnum = 0;
				for (Object obj : objArr) {
					Cell cell = row.createCell(cellnum++);
					if (obj instanceof String)
						cell.setCellValue((String) obj);
					else if (obj instanceof Integer)
						cell.setCellValue((Integer) obj);
				}
			}
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new File(filePath));
			workbook.write(out);
			out.close();
			// System.out.println("output.xlsx written successfully on disk.");
		} catch (Exception e) {
			System.out.println("Exception in  method readOfferExcelSheet()");
		}

	}

}
