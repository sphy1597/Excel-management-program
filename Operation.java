import java.io.File;
import java.io.FileOutputStream;
 
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Operation {
 
    public static void main (String[] args) {
    	
    	try {
    		
    		File file = new File("data.xlsx");
    		
    		FileInputStream fis = new FileInputStream(file);
    		XSSFWorkbook workbook = new XSSFWorkbook(fis);
    		
    		// 가로 , 세로
    		int rowindex = 0;
    		int colindex = 0;
    		
    		//시트수
    		XSSFSheet sheet = workbook.getSheetAt(0);
    		//행 수
    		int rows = sheet.getPhysicalNumberOfRows();
    		
    		for(rowindex=1;rowindex<rows;rowindex++) {
    			
    			//행읽기 
    			XSSFRow row = sheet.getRow(rowindex);
    			XSSFCell cell = row.getCell(2);
    			
    			
    			
    			
    		}
    		
    		
    	}catch(Exception e) {
    		e.printStackTrace();
    	}

 
        
    }
 
}