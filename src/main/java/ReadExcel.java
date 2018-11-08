import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ReadExcel {
    public static void main(String[] args) throws IOException{
        ReadExcel xlsMain = new ReadExcel("dep.xlsx");
        List<String[]> list = xlsMain.readAllExcel();
        System.out.println("行=="+list.size());
        System.out.println("列=="+list.get(0).length);
        WriteExcel.WriteExcel(list);
    }

    private Workbook workbook = null;

    public ReadExcel(String fileName) throws IOException{
        InputStream is = new FileInputStream(fileName);
        if (fileName.endsWith("xls")){
            workbook = new HSSFWorkbook(is);
        } else if (fileName.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(is);
        }
    }

    /**
     * 读取整个excel表数据，支持xls和xlsx
     * @return
     * @throws IOException
     */
    private List<String[]> readAllExcel(){
        List<String[]> contentList = new ArrayList<String[]>();
        //循环sheet表
        for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet ++){
            Sheet sheet = workbook.getSheetAt(numSheet);
            if (sheet == null){
                continue;
            }

            //获取sheet表中的首行数据
            Row row01 = sheet.getRow(0);
            //获取列数
            int cellsNum = row01.getLastCellNum();

            //循环获取表中的行数据
            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum ++){
                Row row = sheet.getRow(rowNum);
                if (row == null){
                    continue;
                }
                String[] rowContents = new String[cellsNum];
                //循环获取每行中的单元格数据
                for (int i = 0; i < cellsNum; i ++){
                    Cell cell = row.getCell(i);
                    if (cell == null){
                        continue;
                    }
                    rowContents[i] = getValue(cell);
                    System.out.print(getValue(cell)+"=");
                }
                System.out.println("");
                contentList.add(rowContents);

            }
        }

        return contentList;
    }

    /**
     * 获取指定表格指定行的所有数据
     * @param sheetNum
     * @param rowNums
     * @return
     */
    private String[] getRows(int sheetNum, int rowNums){

        Sheet sheet = workbook.getSheetAt(sheetNum);
        if (sheet == null){
            return null;
        }
        Row row = sheet.getRow(rowNums);
        int cellNum = row.getLastCellNum();
        String[] rowContents = new String[cellNum];
        for (int i = 0; i < cellNum; i ++){
            Cell cell = row.getCell(i);
            if (cell == null){
                continue;
            }
            rowContents[i] = getValue(cell);
        }
        return rowContents;
    }

    /**
     * 获取指定表格指定列的所有数据
     * @param sheetNum
     * @param coluNum
     * @return
     */
    private String[] getColumn(int sheetNum, int coluNum){

        Sheet sheet = workbook.getSheetAt(sheetNum);
        int rowNums = sheet.getLastRowNum();
        String[] coluContents = new String[rowNums];
        for (int rowNum = 0; rowNum < rowNums; rowNum++){
            Row row = sheet.getRow(rowNum);
            if (row == null){
                continue;
            }
            Cell cell = row.getCell(coluNum);
            coluContents[rowNum] = getValue(cell);
        }
        return coluContents;
    }

    /**
     * 读取指定单元格的数据，支持xls和xlsx
     * @param rowNum  行号
     * @param columNum 列号
     * @return
     */
    private String readExcelByRC(int sheetNum, int rowNum, int columNum) throws IOException{
        Sheet sheet = workbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(rowNum);
        Cell cell = row.getCell(columNum);
        if (cell == null){
            return null;
        }
        return getValue(cell);
    }

    /**
     * 获取excel文件的sheet表格数
     * @return
     */
    private int getSheets(){
        return workbook.getNumberOfSheets();
    }

    /**
     * 获取指定表格的行数
     * @return
     */
    private int getRowNums(int sheetNum){
        Sheet sheet = workbook.getSheetAt(sheetNum);
        return sheet.getLastRowNum();
    }

    /**
     * 获取指定表格的列数
     * @return
     */
    private int getColumNums(int sheetNum){
        Sheet sheet = workbook.getSheetAt(sheetNum);
        return sheet.getRow(0).getLastCellNum();
    }





    private String getValue(Cell cell){
        if (cell.getCellType() == CellType.BOOLEAN){
            return String.valueOf(cell.getBooleanCellValue());
        }else if(cell.getCellType() == CellType.NUMERIC){
            String result = null;
            if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                SimpleDateFormat sdf = null;
                System.out.println("日期格式："+cell.getCellStyle().getDataFormat());
                if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                    sdf = new SimpleDateFormat("HH:mm");
                } else {// 日期
                    sdf = new SimpleDateFormat("yyyy-MM-dd");
                }
                Date date = cell.getDateCellValue();
                result = sdf.format(date);
            } else if (cell.getCellStyle().getDataFormat() == 58) {
                // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                double value = cell.getNumericCellValue();
                Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
                result = sdf.format(date);
            } else {
                double value = cell.getNumericCellValue();
                CellStyle style = cell.getCellStyle();
                DecimalFormat format = new DecimalFormat();
                String temp = style.getDataFormatString();
                // 单元格设置成常规
                if (temp.equals("General")) {
                    format.applyPattern("#");
                }
                result = format.format(value);
            }
            return result;
        }else{
            return String.valueOf(cell.getStringCellValue());
        }
    }
}
