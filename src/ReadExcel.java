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
  	String path= "C:\\Users\\Jasleen Saini\\workspace\\ReadTest\\INC.xlsx";
//  	try{
//		Process proc = Runtime.getRuntime().exec("cmd /c dir \"" + path + "\" /tc");
//		BufferedReader br = new BufferedReader( new InputStreamReader(proc.getInputStream()));
//		String data ="";
//		for(int i=0; i<6; i++){
//			data = br.readLine();
//		}
//		System.out.println("Extracted value : " + data);
//		//split by space
//		StringTokenizer st = new StringTokenizer(data);
//		String date = st.nextToken();//Get date
//		String time = st.nextToken();//Get time
//
//		System.out.println("Creation Date  : " + date);
//		System.out.println("Creation Time  : " + time);
//	}catch(IOException e){
//		e.printStackTrace();
//	}
  	
      XSSFWorkbook workbook = new XSSFWorkbook(fis);
      XSSFSheet spreadsheet = workbook.getSheetAt(0);
      Iterator < Row > rowIterator = spreadsheet.iterator();
      String[] Bname = new String[200];
      int s=0;
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
            	   float id= (float) test1.getNumericCellValue();   
            	   if(test1.getNumericCellValue()==1.0){
            	  // String add = (String) test1.getAddress();
            	  System.out.println(id + test1.getAddress().toString());
//            	   if(id==1.0){
            		String middle;
            		String name= row.getCell(3).toString();
            		Cell testmid= row.getCell(4);
            		if(testmid==null)
            		 middle= "";
            		else
            		 middle = row.getCell(4).toString();
            		String last= row.getCell(5).toString();
            		fullName= name+" "+middle+" "+last;
            		//System.out.println(s);
            		Bname[s]=fullName;  
            		//System.out.println(Bname[s]);
            		s++;	
            		break;
            	   }
//            	   else
             //  System.out.print(cell.getNumericCellValue() + " \t\t " );
              // break;
               case Cell.CELL_TYPE_STRING:
              // System.out.print(cell.getStringCellValue() + " \t\t " );
               break;          
            }
          
         }
        
//         Bname = Arrays.stream(Bname)                                   //Removing null values from Bname array
//                 .filter(q -> (q != null && q.length() > 0))
//                 .toArray(String[]::new);    
//         java.util.List<String> list= new ArrayList<String>();
//        // List<String> list = new ArrayList<String>();
        //for(int j=0; j<Bname.length; j++){
        //	System.out.println(Bname[j]);
//        	 if(Bname[j]!=null){
//        		 list.add(Bname[j]);
       	// }
//         }
//         String[] Sname= new String[200];
//         Sname = list.toArray(new String[list.size()]);
//         for(int j=0; j<Sname.length; j++){
//        	 System.out.println(Sname[j]);
//         }
         System.out.println();
         
      }
      fis.close();
   }
}