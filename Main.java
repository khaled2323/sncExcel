package snc.com;
import java.io.File;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellRange;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 
import java.util.ArrayList; 
import java.util.Scanner;

public class Main {
	
	// Has all numbers that appear on excel cells
			static ArrayList<Double> EXCELCELLS = new ArrayList<Double>();
			// has all quantities in sheet
			static ArrayList<Double> Quantity = new ArrayList<Double>();
			// has all unit prices in sheet
			static ArrayList<Double> UnitPriceinRiyals = new ArrayList<Double>();
			// will have updated total price 
			static ArrayList<Double> TotalPriceEach = new ArrayList<Double>();
			static double totalPriceForAllItems = 0;

	
	static void riyalsFunction() {
	    System.out.println("Unit price in Riyals");
				
		try {
			File file = new File("./src/snc/com/sample.xlsx");
			FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
			//creating Workbook instance that refers to .xlsx file  
			XSSFWorkbook wb = new XSSFWorkbook(fis);   
			XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
			Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
			while(itr.hasNext()) {
				Row row = itr.next();  
				Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();  
					switch (cell.getCellType()) {						
					case STRING: {
						// do something
//						System.out.print(cell.getStringCellValue() + "\t\t\t");
						break;
					}
					case NUMERIC: {
						// do something else
//						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						EXCELCELLS.add(cell.getNumericCellValue());
						break;
					}
					default: {
						// do nothing
					}	
//					System.out.println(""); 
				}
			}
		}	
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		for(int num = 0; num < EXCELCELLS.size(); num++) {
			num++;
			Quantity.add(EXCELCELLS.get(num));
			num++;
			UnitPriceinRiyals.add(EXCELCELLS.get(num));
			num++;
		}
		
		for(int i = 0; i < Quantity.size(); i++) {
			double temp = 0;
			temp = Quantity.get(i)*UnitPriceinRiyals.get(i);
			TotalPriceEach.add(temp);
		}
		
		for (int j = 0; j < TotalPriceEach.size(); j++) {
			totalPriceForAllItems = totalPriceForAllItems + TotalPriceEach.get(j);
		}

		System.out.println(totalPriceForAllItems);
	  }
	
	static void nonRiyalsFunction(double x) {
		// put code here
		
	    System.out.println("Rate is: " + x);
		
		try {
			File file = new File("./src/snc/com/sample.xlsx");
			FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
			//creating Workbook instance that refers to .xlsx file  
			XSSFWorkbook wb = new XSSFWorkbook(fis);   
			XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
			Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
			while(itr.hasNext()) {
				Row row = itr.next();  
				Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();  
					switch (cell.getCellType()) {						
					case STRING: {
						// do something
//						System.out.print(cell.getStringCellValue() + "\t\t\t");
						break;
					}
					case NUMERIC: {
						// do something else
//						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						EXCELCELLS.add(cell.getNumericCellValue());
						break;
					}
					default: {
						// do nothing
					}	
//					System.out.println(""); 
				}
			}
		}	
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		for(int num = 0; num < EXCELCELLS.size(); num++) {
			num++;
			Quantity.add(EXCELCELLS.get(num));
			num++;
			UnitPriceinRiyals.add(EXCELCELLS.get(num));
			num++;
		}
		
		for(int i = 0; i < Quantity.size(); i++) {
			double temp = 0;
			temp = Quantity.get(i)*UnitPriceinRiyals.get(i)*x;
			TotalPriceEach.add(temp);
		}
		
		for (int j = 0; j < TotalPriceEach.size(); j++) {
			totalPriceForAllItems = totalPriceForAllItems + TotalPriceEach.get(j);
		}
		
//		System.out.println(Quantity.get(0));
//		System.out.println(UnitPriceinRiyals.get(0));
//		System.out.println(Quantity.get(1));
//		System.out.println(UnitPriceinRiyals.get(1));
//		System.out.println(TotalPriceEach.get(0));
//		System.out.println(TotalPriceEach.get(1));
//		System.out.println(TotalPriceEach.get(16));
		System.out.println(totalPriceForAllItems);
	
	}
	
	static void writeColumnAdditions(Double cRate) throws EncryptedDocumentException, IOException {
		File file = new File("./src/snc/com/sam.xlsx");
		try {
			FileInputStream fis = new FileInputStream(file);
			
			Workbook workbook = WorkbookFactory.create(fis) ;
			Sheet sheet = workbook.getSheetAt(0);
			
			int rowCount = 2; // get last row num
			int columnCount;
			if(cRate == 0) {
		     columnCount = sheet.getDefaultColumnWidth() ; // find current column width for Riyals
			}else {
				   columnCount = sheet.getDefaultColumnWidth() + 1 ;
			}
			
			Double inputTemp [] = new Double[TotalPriceEach.size()]; // get array size of this arrayList since this stupid library doesnt support arrayLists
			inputTemp = TotalPriceEach.toArray(new Double[0]);
			
			for (int i=0; i<TotalPriceEach.size();i++) {
			
			Row row =  sheet.getRow(rowCount);
			  Cell cell = row.createCell(columnCount);
			  rowCount++;
			  cell.setCellValue(inputTemp[i]);
			  
			}
			   Row rowTotal = sheet.createRow(rowCount); 
			   Cell cellTotal = rowTotal.createCell(columnCount);
			   cellTotal.setCellValue(totalPriceForAllItems);
			  
	           FileOutputStream outputStream = new FileOutputStream("./src/snc/com/sam.xlsx");
	           workbook.write(outputStream);
	           
	           workbook.close();
	          
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		}   //obtaining bytes from the file  
	}
	
	
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
			System.out.print("Program 1 started\n");
		 Scanner myObj = new Scanner(System.in);  // Create a Scanner object
		    System.out.println("Enter Exchange rate, 0 if already in Saudi Riyals");

		    String rate = myObj.nextLine();  // Read user input
		    System.out.println("Exchange rate is: " + rate);  // Output user input
		    double cRate = Double.parseDouble(rate);
		    if(cRate == 0) {
		    	// means that its in saudi riyals
		    	riyalsFunction();
		    	writeColumnAdditions(cRate);
		    	
		    } else {
		    	// means it is not in Saudi riyals
		    	nonRiyalsFunction(cRate);
		    	writeColumnAdditions(cRate);
		    }	
		    myObj.close();
	}
}