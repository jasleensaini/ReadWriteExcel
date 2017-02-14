import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.List;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.util.StringTokenizer;

public class ReadExcel
{
   static XSSFRow row;
   public static void main(String[] args) throws Exception 
   {
      FileInputStream fis = new FileInputStream(new File("C:/Users/Jasleen Saini/workspace/ReadTest/INC.xlsx"));
  	XSSFWorkbook workbook = new XSSFWorkbook(fis);
      XSSFSheet spreadsheet = workbook.getSheetAt(0);
      Iterator < Row > rowIterator = spreadsheet.iterator();
      ArrayList<String> Bid = new ArrayList<String>();
      ArrayList<String> Bname = new ArrayList<String>();
      String  yy= null;
      String ddd= null;
      String mm=null;
      String dd= null;
      while (rowIterator.hasNext()) 
      {
         row = (XSSFRow) rowIterator.next();
         Iterator < Cell > cellIterator = row.cellIterator();
         while ( cellIterator.hasNext()) 
         {
            Cell cell = cellIterator.next();
            String fullName="";                     
            switch (cell.getCellType()) 
            {
               case Cell.CELL_TYPE_NUMERIC:
            	   int col=0;
            	   Cell test1= row.getCell(col);   
           	   if(test1.getNumericCellValue()==1.0){
            		String middle;
            		String name= row.getCell(3).toString();
            	  String last= row.getCell(5).toString();
            		Cell testmid= row.getCell(4);
            		if(testmid!=null)
            		 middle = row.getCell(4).toString(); 
            		else
            			middle="";
            		fullName= name+" "+middle+" "+last;
            	   	Bname.add(fullName);
            	   
           	   }
            	   	if(test1.getNumericCellValue()==0.0){
            	   String Id= row.getCell(1).toString();    
            	   String[] clientId= Id.split("0000");
            	   Bid.add(clientId[1]);
            	   ddd= row.getCell(2).toString();
            	   }
            	   
                break;
               case Cell.CELL_TYPE_STRING:
              // System.out.print(cell.getStringCellValue() + " \t\t " );
               break;          
            }  
         }           
      }
  	//System.out.println(Bname);
      String[] temp;
      temp = ddd.split("");
      for(int i =0; i < temp.length ; i++){
    	   dd= temp[0]+temp[1];
    	   mm= temp[2]+temp[3];
    	   yy= temp[6]+temp[7];
      }
      switch(mm){
      case "01": mm="Jan";
      			break;
      case "02": mm="Feb";
		break;
      case "03": mm="Mar";
		break;
      case "04": mm="Apr";
		break;
      case "05": mm="May";
		break;
      case "06": mm="Jun";
		break;
      case "07": mm="Jul";
		break;
      case "08": mm="Aug";
		break;
      case "09": mm="Sep";
		break;
      case "10": mm="Oct";
		break;
      case "11": mm="Nov";
		break;
      case "12": mm="Dec";
		break;

			
      }
      fis.close();
      
      // Writing in Daily.xlsx
      String excelFilePath = "C:/Users/Jasleen Saini/workspace/ReadTest/Daily.xlsx";
  	try {
  		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
  		Workbook wb = WorkbookFactory.create(inputStream);
  	      XSSFSheet sheet = (XSSFSheet) wb.getSheetAt(0);   
  	     
  		int rowCount = sheet.getLastRowNum();
  		//System.out.println(rowCount);
  		for (int i=0; i<Bid.size(); i++) {
  			XSSFRow row = sheet.createRow(++rowCount);
  			int columnCount = 0;
  			Cell cell = row.createCell(columnCount);
  			cell.setCellValue(yy+"."+mm+"."+rowCount+1);	
  			columnCount = 1;
 			 cell = row.createCell(columnCount);
 			cell.setCellValue(mm+"."+yy);
  			columnCount = 2;
  			 cell = row.createCell(columnCount);
  			cell.setCellValue(dd+"."+mm+"."+yy);
  			columnCount = 5;
 			 cell = row.createCell(columnCount);
 			cell.setCellValue(Bid.get(i));
 			columnCount = 7;
			 cell = row.createCell(columnCount);
			cell.setCellValue("OK");
  			}
  		inputStream.close();
  		FileOutputStream outputStream = new FileOutputStream("C:/Users/Jasleen Saini/workspace/ReadTest/Daily.xlsx");
  			wb.write(outputStream);
  			System.out.println("Written");
  			wb.close();
  			outputStream.close();
  			}
  	catch(Exception e){
  		System.out.println("Hello Exception");
  	}

      
      
      
      
   }
}