/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package invertorymanagement;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import jdk.nashorn.internal.codegen.CompilerConstants;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.CORBA.DomainManagerOperations;

/**
 *
 * @author partha
 */
public class logindatacheck  extends  AdminPage{
    static FileInputStream fis;
   static HSSFWorkbook hssfw;
   static HSSFSheet hssfs;
   static FormulaEvaluator formulaEvaluator;
   static String Id,Password;
   static String cellid,cellpass;
    public logindatacheck( String Id,String Password) throws FileNotFoundException, IOException{

        this.Id=Id;
        this.Password=Password;
        
        //System.out.println("Id:"+this.Id+"\n Pass: "+this.Password);
        
          fis = new FileInputStream("Login.xlsx");
           hssfw = new HSSFWorkbook(fis);
        hssfs = hssfw.getSheetAt(0);
      //  formulaEvaluator= hssfw.getCreationHelper().createFormulaEvaluator();
       

    
        
    }
    public boolean matchlogin(){
                  for (int rowIndex = 0; rowIndex <= hssfs.getLastRowNum(); rowIndex++) {
                  Row row = hssfs.getRow(rowIndex);
                  if (row != null) {
               Cell cell = row.getCell(2);
               Cell cell1 = row.getCell(3);
               Cell cell2 = row.getCell(0);
    if (cell != null) {
      // Found column and there is value in the cell.
      cellid = cell.getStringCellValue();
      // Do something with the cellValueMaybeNull here ...
       // System.out.println(cellid);
    }
      if (cell1 != null) {
      // Found column and there is value in the cell.
     cellpass = cell1.getStringCellValue();
      // Do something with the cellValueMaybeNull here ...
       // System.out.println(cellpass);
    }
      if(cellid.equals(Id) && cellpass.equals(Password)){
                  
         System.out.println(cell2.getStringCellValue());
                 
             
               usrname = cell2.getStringCellValue();
               return true;
             
              }
  }
}
              return false;
    }
         
   
  
        
}
