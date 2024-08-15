import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class xlsx_read {

    public static String filePath = "/Users/na/Desktop/letsGoPoiTest";
    public static String fileNm = "테스트.xlsx";

    public static void main(String[] args){
        try(FileInputStream file = new FileInputStream(new File(filePath, fileNm)))    {
            // xlsx 로 Workbook instance 생성
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // workbook 의 첫번째 sheet 를 가져옴
            XSSFSheet sheet = workbook.getSheetAt(0);

            // 모든 row 조회
            for(Row row : sheet){
                Iterator<Cell> cellIterator = row.cellIterator();

                while(cellIterator.hasNext()){
                    Cell cell = cellIterator.next();

                    switch(cell.getCellType()){

                        case NUMERIC:
                            System.out.println((int) cell.getNumericCellValue() + "\t");
                            break;

                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println();
            }

        }catch(IOException e){
            e.printStackTrace();
        }
    }

}
