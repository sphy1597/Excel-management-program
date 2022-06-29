import java.io.File;
import java.io.FileOutputStream;
 
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.*;


public class Operation {
 
	static ArrayList<String> colname = new ArrayList<String>(); // 칼럼이름 저장 
	static ArrayList<Dataform> datas = new ArrayList<Dataform>(); // 데이터폼으로 데이터 객체 저장
	
	static File file;
	static FileInputStream fis;
	static XSSFWorkbook workbook;
	
    public static void main (String[] args) {
    	
    	
    	try {
    		file = new File("data.xlsx");
			fis = new FileInputStream(file);
			workbook = new XSSFWorkbook(fis);
			
    	}catch(Exception e) {
    		e.getStackTrace();
    	}
    	Readdata rd = new Readdata(file, fis, workbook);
		
    	rd.read();
    	rd.printdata();
    	
 
    }
 
}