package com.test.java.OfferIssue;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
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
			while (rowIterator.hasNext()) {
				System.out.println("Row Number # " + ++rowNumber);
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				int inmpactedOffers = 0;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					// cell.getRowIndex() + " "
					if (cell.getColumnIndex() > 1) {
						Matcher m = p.matcher(cell.toString());
						if (m.find()) {
							inmpactedOffers++;
							System.out.print(cell.toString() + " ");
						}
					}
				}
				
				if(inmpactedOffers>0)
				System.out.print(" " + inmpactedOffers + "\n");
			}
			file.close();
		} catch (Exception e) {
			System.out.println("Exception in  method readOfferExcelSheet ");
		}

	}
	
	public static void writeIvvalidOffers()
	{
		System.out.println("Finding Invalid Offers");
	}

}
