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
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author partha
 */
public class MatchData {
    
   static FileInputStream fis;
   static  File file;
   static  File f;
   static HSSFWorkbook hssfw;
   static Workbook wb;
   static HSSFSheet hssfs;
   static String Code,cellid,name;
   static int Quantity,q,RemainQuantity;
   static double Price,TotalPrice;
   static boolean  res;
    
    public MatchData(String ItemCode,int Quan ) throws FileNotFoundException, IOException{
     
        
         Code = ItemCode;
        Quantity =Quan;
       
        
        //System.out.println("Id:"+this.Id+"\n Pass: "+this.Password);
         
           f =new File("Items.xlsx");
           fis = new FileInputStream(f);
           file =new File("Items.xlsx");
            if(file.exists() && f.exists()){
       hssfw = new HSSFWorkbook(fis);
           hssfs = hssfw.getSheetAt(0);
         //  s=hssfw.getSheetAt(0);
     }
           
      //  formulaEvaluator= hssfw.getCreationHelper().createFormulaEvaluator();
       

    
        
 //   }
                   
//    public boolean matchdata() throws IOException{
                  for (int rowIndex = 0; rowIndex <= hssfs.getLastRowNum(); rowIndex++) {
                           Row row = hssfs.getRow(rowIndex);
                  if (row != null) {
                     
                      Cell cell0 =row.getCell(0);
                     Cell cell = row.getCell(1);
                     Cell cell1 = row.getCell(2);
                     Cell cell2 = row.getCell(3);
                     Cell cell3 = row.getCell(4);
                  if (cell != null) {
                        // Found column and there is value in the cell.
                      cellid = cell.getStringCellValue();
                     
                      if(cellid.equals(Code)){
                            name = cell0.getStringCellValue();
                            q = (int)cell1.getNumericCellValue();
                            
                            if(q>0 && (q-Quantity)>0){
                            Price= (int) cell2.getNumericCellValue();
                            RemainQuantity= q-Quantity;
                            TotalPrice = RemainQuantity*Price;
                           
                            Row rowh = hssfs.getRow(rowIndex);
                    
                            hssfs.removeRow(rowh);
                            
                            row = hssfs.createRow(rowIndex);
                            row.createCell(0).setCellValue(this.name);
                            row.createCell(1).setCellValue(cellid);
                            row.createCell(2).setCellValue(this.RemainQuantity);
                            row.createCell(3).setCellValue(this.Price);
                            row.createCell(4).setCellValue(this.TotalPrice);

                           res = true;
                            }
                            else{
                              res = false;
                            }
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
