import java.io.File;
import java.io.FileOutputStream;
 
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.*;


public class Operation {
 
	static DataManager dm;
	
	
    public static void main (String[] args) {
    	
    	Operation op = new Operation();
    	
    	
    	dm = new DataManager();	
    	dm.ReadData();
    	dm.PrintData();
    	dm.WriteData();
    }
    
 
}