package Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Excel_Utils1 {
    // static String filebath="C:\\Users\\j416\\eclipse-workspace\\Page_Load\\Excel\\Page_Load.xlsx";
static Excel_Reader reader;
    
    public static  ArrayList<Object[]> getDataFromexcel() {
    

        ArrayList<Object[]> myData=new ArrayList<Object[]>();
        try {
        
        	
            reader=new Excel_Reader("E:\\selenium\\SEO_content\\SEO_Automation\\Excel\\Book1.xlsx");
            
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        for (int rowNum = 2; rowNum <= reader.getRowCount("Sheet2"); rowNum++) {
          
         
            
             
            Object ob[]= {};
            		
            myData.add(ob);
        }
        return myData;

    }

    
    public static void writeinexcel2(String value, int INDEX) throws Exception {
        Thread.sleep(2000);
        
        File fis = new File("E:\\selenium\\SEO_content\\SEO_Automation\\Excel\\Book1.xlsx");
        FileInputStream excelloc = new FileInputStream(fis);
        
        XSSFWorkbook wb = new XSSFWorkbook(excelloc);
        XSSFSheet s = wb.getSheetAt(1);
        XSSFRow row = s.getRow(INDEX);
        XSSFCell c = row.createCell(2);
        
        c.setCellValue(value);
        
        FileOutputStream out = new FileOutputStream(fis);
        wb.write(out);
        out.close();
    }
    public static void writeinexcel3(String value, int INDEX) throws Exception {
    	Thread.sleep(2000);
        File fis = new File("E:\\selenium\\SEO_content\\SEO_Automation\\Excel\\Book1.xlsx");
        FileInputStream excelloc = new FileInputStream(fis);
        XSSFWorkbook wb = new XSSFWorkbook(excelloc);
        XSSFSheet s = wb.getSheetAt(1);
        XSSFRow row = s.getRow(INDEX);
        XSSFCell c = row.createCell(3);
        
        c.setCellValue(value);
        FileOutputStream out = new FileOutputStream(fis);
        wb.write(out);
        out.close();
    }
    public static void writeinexcel4(String value, int INDEX) throws Exception {
        Thread.sleep(2000);
        File fis = new File("E:\\selenium\\SEO_content\\SEO_Automation\\Excel\\Book1.xlsx");
        FileInputStream excelloc = new FileInputStream(fis);
        XSSFWorkbook wb = new XSSFWorkbook(excelloc);
        XSSFSheet s = wb.getSheetAt(1);
        XSSFRow row = s.getRow(INDEX);
        XSSFCell c = row.createCell(4);
        
        c.setCellValue(value);
        FileOutputStream out = new FileOutputStream(fis);
        wb.write(out);
        out.close();
    }
    public static void writeinexcel5(String value, int INDEX) throws Exception {
        Thread.sleep(2000);
        File fis = new File("E:\\selenium\\SEO_content\\SEO_Automation\\Excel\\Book1.xlsx");
        FileInputStream excelloc = new FileInputStream(fis);
        XSSFWorkbook wb = new XSSFWorkbook(excelloc);
        XSSFSheet s = wb.getSheetAt(1);
        XSSFRow row = s.getRow(INDEX);
        
        XSSFCellStyle my_style = wb.createCellStyle();
        XSSFFont my_font=wb.createFont();
        my_font.setColor(XSSFFont.COLOR_RED);
        my_style.setFont(my_font);
        
        XSSFCell c = row.createCell(4);
        
        c.setCellValue(value);
        c.setCellStyle(my_style);
        FileOutputStream out = new FileOutputStream(fis);
        wb.write(out);
        out.close();
    }
    public static void writeinexcel6(String value, int INDEX) throws Exception {
        Thread.sleep(2000);
        File fis = new File("E:\\selenium\\SEO_content\\SEO_Automation\\Excel\\Book1.xlsx");
        FileInputStream excelloc = new FileInputStream(fis);
        XSSFWorkbook wb = new XSSFWorkbook(excelloc);
        XSSFSheet s = wb.getSheetAt(1);
        XSSFRow row = s.getRow(INDEX);
        XSSFCell c = row.createCell(0);
        
        c.setCellValue(value);
        FileOutputStream out = new FileOutputStream(fis);
        wb.write(out);
        out.close();
    }

    public static void writeinexcel7(String value, int INDEX) throws Exception {
        Thread.sleep(2000);
        File fis = new File("E:\\selenium\\SEO_content\\SEO_Automation\\Excel\\Book1.xlsx");
        FileInputStream excelloc = new FileInputStream(fis);
        XSSFWorkbook wb = new XSSFWorkbook(excelloc);
        XSSFSheet s = wb.getSheetAt(1);
        XSSFRow row = s.getRow(INDEX);
        XSSFCell c = row.createCell(1);
        
        c.setCellValue(value);
        FileOutputStream out = new FileOutputStream(fis);
        wb.write(out);
        out.close();
    }
    public static void writeinexcel8(String value, int INDEX) throws Exception {
        Thread.sleep(2000);
        File fis = new File("E:\\selenium\\SEO_content\\SEO_Automation\\Excel\\Book1.xlsx");
        FileInputStream excelloc = new FileInputStream(fis);
        XSSFWorkbook wb = new XSSFWorkbook(excelloc);
        XSSFSheet s = wb.getSheetAt(1);
        XSSFRow row = s.getRow(INDEX);
        XSSFCell c = row.createCell(5);
        
        c.setCellValue(value);
        FileOutputStream out = new FileOutputStream(fis);
        wb.write(out);
        out.close();
    }


}
