import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class xlsx_write {
    public static String filePath = "/Users/na/Desktop/letsGoPoiTest";
    public static String fileNm = "테스트.xlsx";

    public static void main(String[] args){
        // 빈 WorkBook 생성
        XSSFWorkbook workbook = new XSSFWorkbook();

        File fullPath = new File(filePath + "/" + fileNm);

        // 상위 디렉토리 존재여부 확인
        // 없으면 생성
        File parentDir = fullPath.getParentFile();
        if(parentDir != null & !parentDir.exists()){
            boolean created = parentDir.mkdirs();
            if(created){
                System.out.println("폴더 생성 : " + parentDir.getAbsolutePath());
            }else{
                System.out.println("폴더 생성 실패");
            }
        }

        // 빈 Sheet 생성
        XSSFSheet sheet = workbook.createSheet("테스트용");
        System.out.println("생성");

        // Sheet 를 채우기 위한 data 들을 Map 에 저장
        Map<String, Object[]> data = new TreeMap<>();
        data.put("1", new Object[]{"ID", "NAME", "PHONE"});
        data.put("2", new Object[]{"1", "최예나", "9992"});
        data.put("3", new Object[]{"2", "구스덕", "124908"});
        data.put("4", new Object[]{"3", "팜핀", "1211"});

        // data 의 keySet 을 가져옴
        // Set 을 조회하면서 data 를 sheet 에 입력
        Set<String> keySet = data.keySet();
        int rownum = 0;

        // TreeMap 은 오름차순
        for(String key : keySet){
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for(Object obj : objArr){
                Cell cell = row.createCell(cellnum++);
                if(obj instanceof  String){
                    cell.setCellValue((String)obj);
                }else if(obj instanceof Integer){
                    cell.setCellValue((Integer)obj);
                }
            }
        }
        try(FileOutputStream out = new FileOutputStream(new File(filePath, fileNm))) {
            System.out.println("생성 완료");
            workbook.write(out);
        }catch(IOException e){
            e.printStackTrace();
        }
    }
}


