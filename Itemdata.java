/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package invertorymanagement;

/**
 *
 * @author partha
 */
public class Itemdata {
   
     
        static String Item_Name;
        static String Item_Code;
        static int Item_Quantity;
        static double Item_Price;
      
    public Itemdata( String Item_Name, String Item_Code,int Item_Quantity,double Item_Price){
        
        this.Item_Name = Item_Name;
        this.Item_Code = Item_Code;
        this.Item_Quantity = Item_Quantity;
        this.Item_Price = Item_Price;
        
    }
    
    
}
