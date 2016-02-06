package jp.designrule;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.io.FileInputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFormatter {

  public static void main(String... args) throws IOException, InvalidFormatException {
    File file = new File(args[0]);
    InputStream is = new FileInputStream(file);
    Workbook workbook = WorkbookFactory.create(is);
    is.close();

    Sheet sheet = workbook.getSheet("Sheet1");
    
    //B31の参照先を書き換え
    CellReference reference = new CellReference("B31"); // A1形式
    Row row = sheet.getRow(reference.getRow());
    if (row != null) {
      Cell cell = row.getCell(reference.getCol());
      if (cell != null) {
        cell.setCellFormula("'01_0表紙'!t28");
      }
    }
    
    //上書き保存
    try (FileOutputStream fos = new FileOutputStream(file)) {
        workbook.write(fos);
      }
  }
}
