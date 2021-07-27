package org.web.report;
 
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class createworkbook { 
public static void main(String args[]) throws IOException{
//To create a new WorkBook with xlsx extension: HSSFWorkbook can be used for .xls format
Workbook wb = new XSSFWorkbook(); 
FileOutputStream fileOut = new FileOutputStream("D:\\TestCase_Framework_Reporting_Bugs.xlsx"); 
CreationHelper createHelper = wb.getCreationHelper();
Sheet sheet = wb.createSheet("Defects");

//Create the first row for headers
Row row = sheet.createRow((short) 0);
for (int i=0;i<10;i++) {
	String a[] = {"S.No.","Issue","Issue-Description","ReportedBy","AssignedTo","CreationDate","UpdatedDate","ClosedDate","Release","Status"};
	row.createCell(i).setCellValue(createHelper.createRichTextString(a[i])); 	
}
wb.write(fileOut);
fileOut.close();
}
}