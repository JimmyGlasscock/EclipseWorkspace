package excelFormatter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {
	
	int columnOne, columnTwo;
	
	String choice, filePath, fileOutPath;

	public static void main(String args[]){
		Main main = new Main();
		main.prompt();
	}
	
	public void prompt(){
		Scanner scanner = new Scanner(System.in);
		System.out.println("Which columns are we combining?");
		System.out.println("First column:");
		columnOne = scanner.nextInt();
		System.out.println("Second column:");
		columnTwo = scanner.nextInt();
		
		System.out.println("First Column: " + columnOne + " Second Column: " + columnTwo + "\n");
		System.out.println("is this correct? (y for yes, n for no)\n");
		choice = scanner.next();
		
		if(choice.equalsIgnoreCase("y")){
			combineColumns(columnOne, columnTwo);
		}else{
			prompt();
		}
	}
	
	public void combineColumns(int colOne, int colTwo){
		try{
			
			Scanner scanner = new Scanner(System.in);
			
			System.out.println("\n Enter file path: \n");
			
			filePath = scanner.nextLine();
			
			fileOutPath = filePath.substring(0,filePath.length()-5) + "(UPDATED).xlsx";
			
			FileInputStream inputStream = new FileInputStream(new File(filePath));
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			
			int numberOfRows = sheet.getLastRowNum();
			
			for(int i = 1; i < numberOfRows; i++){
				
				if(sheet.getRow(i).getCell(colOne) != null && sheet.getRow(i).getCell(colTwo) != null){
					if(sheet.getRow(i).getCell(colOne).toString().equals(sheet.getRow(i).getCell(colTwo).toString())){
						System.out.println("(" + colOne + "," + i + ") and (" + colTwo + "," + i + ") are equal");
					}
					
					if(sheet.getRow(i).getCell(colOne).toString().length() < sheet.getRow(i).getCell(colTwo).toString().length()){
						sheet.getRow(i).getCell(colOne).setCellValue(fixFormfactor(sheet.getRow(i).getCell(colTwo).toString()));
						System.out.println("Copied cell (" + colTwo + "," + i + ") to (" + colOne + "," + i + ")");
					}
				}else{
					if(sheet.getRow(i).getCell(colTwo) != null){
						sheet.getRow(i).createCell(colOne);
						sheet.getRow(i).getCell(colOne).setCellValue(fixFormfactor(sheet.getRow(i).getCell(colTwo).toString()));
						System.out.println("Copied cell (" + colTwo + "," + i + ") to (" + colOne + "," + i + ")");
					}
				}
			}
			
			System.out.println("\n Done! \n \n Complete data is stored in column " + colOne);
			
			inputStream.close();
			
			FileOutputStream outputStream = new FileOutputStream(fileOutPath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
		
		}catch(EncryptedDocumentException ex){
			ex.printStackTrace();
		}catch(IOException e) {
			System.out.println("\n File not found... Try again. \n");
			combineColumns(colOne, colTwo);
		}
		
		System.out.println("\n Output file is stored at " + fileOutPath);
	}
	
	public String fixFormfactor(String str) {
		if(str.substring(1,2).equals(".")&&str.substring(str.length()-2, str.length()-1).equalsIgnoreCase("E")) {
			str = str.substring(0,1) + str.substring(2,str.length()-2);
		}
		
		if(str.substring(str.length()-2).equals(".0")) {
			str = str.substring(0, str.length()-2);
		}
		
		return str;
	}
	
}
