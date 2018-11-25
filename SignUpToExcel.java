/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package invertorymanagement;

import java.io.*;
import java.io.FileReader;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.List;
 
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author partha
 */
public class SignUpToExcel {
     ////////////////////////////////////// variables /////////////////////////
   static String [] Column = {"Name","Phone Number","ID","Password","Gender"};
  // static ArrayList<SignUp> Signup =new ArrayList<>();
    static List<SignUp> signup = new ArrayList<SignUp>();
   
/////////////////////////// io //////////////////////////////
     static File file=new File("Login.xlsx");
    static FileWriter fileWriter;
    static FileReader fileReader;
  //  static WritableWorkbook myWorkbook ;
    static HSSFWorkbook workbook;
    static Row hearderRow;
    static CellStyle headCellStyle;
    //////////////////////////////////////////
        static String Name;
        static String PhoneNum;
        static String ID;
        static String Password;
        static String Gender;
        static HSSFSheet mysFSheet;
        static org.apache.poi.ss.usermodel.Font font;
        static File f= new File("Rowdata.txt");
        static FileWriter fw;
        static FileReader fr;
        static int rownum=1;
        static String rowread;
        static int r;
        static BufferedReader bf;
 
        
    public SignUpToExcel(String Name,String PhoneNum,String ID,String Password,String Gender) throws NullPointerException, FileNotFoundException, IOException{
         this.Name =Name;
         this.PhoneNum=PhoneNum;
        this.ID = ID;
        this.Password =Password;
        this.Gender=Gender;
        
       
        signup.add(new SignUp(Name,PhoneNum, ID, Password,Gender));
        
     if(!file.exists()){
        workbook = new HSSFWorkbook();
        mysFSheet = workbook.createSheet("Contacts");
        
     }
     else{
        
         workbook =new HSSFWorkbook(new FileInputStream(file));
         
         mysFSheet=workbook.getSheet("Contacts");
     }
        font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short)14);
        font.setColor(IndexedColors.BLUE.getIndex());
        
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        
        Row headerRow = mysFSheet.createRow(0);
        for(int i=0;i<Column.length;i++){
        Cell cell = headerRow.createCell(i);
        cell.setCellValue(Column[i]);
        }
       
        for (SignUp signUp :signup ) {
            if(f.exists()){
                bf=new BufferedReader( new FileReader(f));
                rowread=bf.readLine();
                rownum=Integer.parseInt(rowread);
                rownum++;
                fw=new FileWriter(f);
               fw.write(""+rownum);
               fw.close();
               
                System.out.println(rowread);
            }
            else{
                fw=new FileWriter(f);
                
                fw.write(""+rownum);
                fw.close();
                bf=new BufferedReader( new FileReader(f));
                rowread=bf.readLine();
                rownum=Integer.parseInt(rowread);
                System.out.println(rowread);
            }
             
            Row row = mysFSheet.createRow(rownum);
            row.createCell(0).setCellValue(SignUp.Name);
            row.createCell(1).setCellValue(SignUp.PhoneNum);
            row.createCell(2).setCellValue(SignUp.ID);
            row.createCell(3).setCellValue(SignUp.Password);
            row.createCell(4).setCellValue(signUp.Gender);
        }
        signup.clear();
        for(int i=0;i<Column.length;i++){
        mysFSheet.autoSizeColumn(i);
    }
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        workbook.close();
        fos.close();
        
        
  }
          
    
 
    
    
}
