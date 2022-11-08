import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;

public class getExcel {

    @Test
    public void getExcel() throws IOException {

        /*
         * 导入表
         * */
        FileInputStream Fis = new FileInputStream("C:\\Users\\liyifan\\Desktop\\海南网点花名册人员清单数据-需补全信息.xlsx");
        FileOutputStream fileOutputStream = null;
        XSSFWorkbook oldTable = new XSSFWorkbook(Fis);
        XSSFWorkbook newWorkBook = new XSSFWorkbook();
        XSSFSheet person = newWorkBook.createSheet("人员明细");

        int j = 1;
        for (int i = 1; i < oldTable.getSheetAt(0).getLastRowNum(); i++) {
            if (oldTable.getSheetAt(0).getRow(i).getCell(3).getStringCellValue().equals
                    (oldTable.getSheetAt(0).getRow(i - 1).getCell(3).getStringCellValue())) {
                XSSFRow row = person.createRow(j);
                row.createCell(0).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(0).getNumericCellValue());
                row.createCell(1).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(1).getStringCellValue());
                row.createCell(2).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(2).getNumericCellValue());
                row.createCell(3).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(3).getStringCellValue());
                row.createCell(4).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(4).getNumericCellValue());
                row.createCell(5).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(5).getStringCellValue());
                row.createCell(6).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(6).getStringCellValue());
                row.createCell(7).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(7).getNumericCellValue());
                row.createCell(8).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(8).getStringCellValue());
                row.createCell(9).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(9).getStringCellValue());
                row.createCell(10).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(10).getStringCellValue());
                j += 1;
                System.out.println("i am your father");
            } else {

//              设置尾行
                XSSFRow row = person.createRow(j);
                row.createCell(0).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(0).getNumericCellValue());
                row.createCell(1).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(1).getStringCellValue());
                row.createCell(2).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(2).getNumericCellValue());
                row.createCell(3).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(3).getStringCellValue());
                row.createCell(4).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(4).getNumericCellValue());
                row.createCell(5).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(5).getStringCellValue());
                row.createCell(6).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(6).getStringCellValue());
                row.createCell(7).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(7).getNumericCellValue());
                row.createCell(8).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(8).getNumericCellValue());
                row.createCell(9).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(9).getStringCellValue());
                row.createCell(10).setCellValue(oldTable.getSheetAt(0).getRow(i).getCell(10).getStringCellValue());

//              置零，新建工作蒲
                j = 0;
                fileOutputStream = new FileOutputStream("C:\\Users\\liyifan\\Desktop\\网点人员明细\\"
                        + oldTable.getSheetAt(0).getRow(i - 1).getCell(3).getStringCellValue() + ".xlsx");
                newWorkBook.write(fileOutputStream);
                newWorkBook = new XSSFWorkbook();
                person = newWorkBook.createSheet("人员明细");
//              调用设置表头函数
                setTitle(person,oldTable);

            }
        }
    }

    /*
    * 设置标题函数
    * */
     static void setTitle(XSSFSheet sheet,XSSFWorkbook oldTable){
        sheet.createRow(0).createCell(0).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        sheet.createRow(0).createCell(1).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(1).getStringCellValue());
        sheet.createRow(0).createCell(2).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(2).getStringCellValue());
        sheet.createRow(0).createCell(3).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(3).getStringCellValue());
        sheet.createRow(0).createCell(4).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(4).getStringCellValue());
        sheet.createRow(0).createCell(5).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(5).getStringCellValue());
        sheet.createRow(0).createCell(6).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(6).getStringCellValue());
        sheet.createRow(0).createCell(7).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(7).getStringCellValue());
        sheet.createRow(0).createCell(8).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(8).getStringCellValue());
        sheet.createRow(0).createCell(9).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(9).getStringCellValue());
        sheet.createRow(0).createCell(10).setCellValue(oldTable.getSheetAt(0).getRow(0).getCell(10).getStringCellValue());
    }
}
