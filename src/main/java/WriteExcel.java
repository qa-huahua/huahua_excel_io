import org.apache.poi.hssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class WriteExcel {



    /**
     * 向excel表中插入数据
     * @param contents
     * @throws IOException
     */
    public static void WriteExcel(List<String[]> contents) throws IOException{
        int countColumnNum = contents.size();
        if(countColumnNum <= 0){
            System.out.println("表格内容为空");
            return;
        }
        HSSFWorkbook hwb = new HSSFWorkbook();
        HSSFSheet sheet = hwb.createSheet("dep");

        for (int i = 0; i < countColumnNum; i ++){
            HSSFRow row = sheet.createRow(i);
            String[] content = contents.get(i);
            for (int colu = 0; colu < content.length; colu ++){
                HSSFCell cell = row.createCell(colu);
                cell.setCellValue(content[colu]);
            }
        }

        OutputStream out = new FileOutputStream("POI2Excel/dep.xls");
        hwb.write(out);
        out.close();

    }
}
