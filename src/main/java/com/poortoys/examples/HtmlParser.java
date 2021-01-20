package com.poortoys.examples;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

public class HtmlParser {

	public static void main(String[] args) throws IOException {
		
		LocalDate date = LocalDate.now();
				
		DateTimeFormatter myFormatObj = DateTimeFormatter.ofPattern("dd.MM.yyyy");

		String formattedDate = date.format(myFormatObj);
		    
		    
		updateExcel(0,calculateLaunchTime("BiP-KPI",formattedDate), calculateLaunchTime("WhatsApp-KPI",formattedDate),formattedDate);    
		
		updateExcel(1,calculateInstallTime("BiP-KPI",formattedDate), calculateInstallTime("WhatsApp-KPI",formattedDate),formattedDate);    
				
	}
	
	//Calculate launch time for BiP and WhatsApp
	public static ArrayList<Double> calculateLaunchTime(String appName, String today) throws IOException {
		ArrayList<Double> list1 = new ArrayList<>();
		
		for(int i=1; i <= 10; i++) {
			
			String fileName = "C:\\Users\\omer_\\Desktop\\KPI\\" + appName + "\\" + today + "\\index" + i + ".html";
			
			Document doc = Jsoup.parse(new File(fileName), "utf-8");
			
			double totalTime=0;
			
			for(int j=10; j <= 15; j++) {
				totalTime += Double.parseDouble(doc.getElementsByClass("panel panel-success").get(j).getElementsByClass("panel-body").text().split(" ")[2]);
			}
			
			list1.add(totalTime);
		}
		return list1;
	}
	
	//Calculate install time for BiP and WhatsApp
	public static ArrayList<Double> calculateInstallTime(String appName, String today) throws IOException {
		ArrayList<Double> list = new ArrayList<>();
				
		for(int i=1; i <= 10; i++) {
			
			String fileName = "C:\\Users\\omer_\\Desktop\\KPI\\" + appName + "\\" + today + "\\index" + i + ".html";
			
			Document doc = Jsoup.parse(new File(fileName), "utf-8");
			
			Double installTime = Double.parseDouble(doc.getElementsByClass("panel panel-success").get(9).getElementsByClass("panel-body").text().split(" ")[2]);
			
			list.add(installTime);
		}
		return list;
	}
	
	
	//Updates existing Excel file.
	public static void updateExcel(int a, ArrayList<Double> timeBip, ArrayList<Double> timeWhatsApp,String date) {
		 String excelFilePath = "C:\\Users\\omer_\\Desktop\\KPI\\RunData.xlsx";
         
	        try {
	            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
	            Workbook workbook = WorkbookFactory.create(inputStream);
	 
	            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(a);
	            
	            
	            sheet.shiftRows(1, sheet.getLastRowNum(), 2, true, true);
	            
	            Row row = sheet.createRow(1);
	            Row row1 = sheet.createRow(2);
	            	            
	            for(int i=0;i<timeBip.size();i++) {
	            	
	            	row.createCell(i+2).setCellValue(timeBip.get(i));
	            	row1.createCell(i+2).setCellValue(timeWhatsApp.get(i));
	            }
	            
	            row.createCell(0).setCellValue(date);
	            
	            row.createCell(1).setCellValue("Bip");
	            row1.createCell(1).setCellValue("WhatsApp");
	           
	            
	            inputStream.close();
	 
	            FileOutputStream outputStream = new FileOutputStream("C:\\Users\\omer_\\Desktop\\KPI\\RunData.xlsx");
	            workbook.write(outputStream);
	            workbook.close();
	            outputStream.close();
	             
	        } catch (IOException | EncryptedDocumentException
	                | InvalidFormatException ex) {
	            ex.printStackTrace();
	        }
	}
	
}
