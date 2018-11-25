
package invertorymanagement;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author partha
 */
public class InsertItemToExcel {
    
    
    
     static String [] Column = {"Item Name","Item Code","Item Quantity","Per Item Price","Total Price"};
  
    static List<Itemdata> InsertItem = new ArrayList<Itemdata>();
   
/////////////////////////// io //////////////////////////////
     static File file=new File("Items.xlsx");
    static FileWriter fileWriter;
    static FileReader fileReader;
  //  static WritableWorkbook myWorkbook ;
    static HSSFWorkbook workbook;
    static Row hearderRow;
    static CellStyle headCellStyle;
    //////////////////////////////////////////
        static String Item_Name;
        static String Item_Code;
        static  int  Item_Quantity;
        static double Item_Price;
        static Double TotalPrice;
        
        static HSSFSheet mysFSheet;
        static org.apache.poi.ss.usermodel.Font font;
        static File f= new File("InsertSheetDataList.txt");
        static FileWriter fw;
        static FileReader fr;
        static int rownum=1;
        static String rowread;
        static int r;
        static BufferedReader bf;
 
    
     public InsertItemToExcel( String Item_Name, String Item_Code,int Item_Quantity,double Item_Price) throws FileNotFoundException, IOException{
        
        this.Item_Name = Item_Name;
        this.Item_Code = Item_Code;
        this.Item_Quantity = Item_Quantity;
        this.Item_Price = Item_Price;
        this.TotalPrice = Item_Price*Item_Quantity;
        InsertItem.add(new Itemdata(this.Item_Name,this.Item_Code,this.Item_Quantity,this.Item_Price));
        
        
        
          if(!file.exists()){
        workbook = new HSSFWorkbook();
        mysFSheet = workbook.createSheet("Items");
        
     }
     else{
        
         workbook =new HSSFWorkbook(new FileInputStream(file));
         
         mysFSheet=workbook.getSheet("Items");
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
       
        for (Itemdata itemdata : InsertItem ) {
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
            row.createCell(0).setCellValue(itemdata.Item_Name);
            row.createCell(1).setCellValue(itemdata.Item_Code);
            row.createCell(2).setCellValue(itemdata.Item_Quantity);
            row.createCell(3).setCellValue(itemdata.Item_Price);
            row.createCell(4).setCellValue(TotalPrice);
            
        }
       InsertItem.clear();
        for(int i=0;i<Column.length;i++){
        mysFSheet.autoSizeColumn(i);
    }
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        workbook.close();
        fos.close();
        
        
        
        
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
}
