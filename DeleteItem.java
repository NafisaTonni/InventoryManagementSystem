/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package invertorymanagement;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author partha
 */
public class DeleteItem {
    
 
    
   static FileInputStream fis;
   static  File file;
   static  File f;
   static HSSFWorkbook hssfw;
   static Workbook wb;
   static HSSFSheet hssfs;
   static FormulaEvaluator formulaEvaluator;
   static String Code,name;
   static String cellid;
   static double Price,TotalPrice;
   static boolean check;
   static Sheet s;
    public DeleteItem( String Item_Code) throws FileNotFoundException, IOException, InvalidFormatException{
          
        Code = Item_Code;
       // Quantity =item_Quantity;
        
        
        //System.out.println("Id:"+this.Id+"\n Pass: "+this.Password);
         
           f =new File("Items.xlsx");
           fis = new FileInputStream(f);
           file =new File("Items.xlsx");
            if(file.exists() && f.exists()){
       hssfw = new HSSFWorkbook(fis);
           hssfs = hssfw.getSheetAt(0);
         //  s=hssfw.getSheetAt(0);
     }
     
                  for (int rowIndex = 0; rowIndex <= hssfs.getLastRowNum(); rowIndex++) {
                           Row row = hssfs.getRow(rowIndex);
                  if (row != null) {
                     
                     Cell cell0 =row.getCell(0);
                     Cell cell = row.getCell(1);
                     Cell cell1 = row.getCell(3);
                    
                  if (cell != null) {
                        // Found column and there is value in the cell.
                      cellid = cell.getStringCellValue();
                      
                      if(cellid.equals(Code)){
                           
                            Row rowh = hssfs.getRow(rowIndex);
                            name = cell0.getStringCellValue();
                            DeleteItem.Price =cell1.getNumericCellValue();
                            hssfs.removeRow(rowh);
                            
                            
                             row = hssfs.createRow(rowIndex);
                            row.createCell(0).setCellValue(this.name);
                            row.createCell(1).setCellValue(cellid);
                            row.createCell(2).setCellValue(0);
                            row.createCell(3).setCellValue(this.Price);
                            row.createCell(4).setCellValue(0);
                            check = true;    
                            System.out.println(name);
              }
       
  }
}
             
    }
         FileOutputStream fos = new FileOutputStream(file);
       
        hssfw.write(fos);
        hssfw.close();
        fos.close();
        
   // return false;
    
       
      
    
}
            
         
}   
    
   