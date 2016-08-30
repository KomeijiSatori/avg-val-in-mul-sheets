package pak;
/**
 * Created by Satori on 2016/8/30.
 */

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;

public class SheetUtils
{
    static ArrayList<String> readError = new ArrayList<>();

    public static double[][] readSheet(String filePath, int rowCnt, int colCnt, int startRow, int startCol)
            throws IOException
    {
        String fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
        InputStream stream = new FileInputStream(filePath);
        Workbook wb = null;
        double[][] vals = new double[rowCnt][colCnt];

        if (fileType.equals("xls"))
        {
            wb = new HSSFWorkbook(stream);
        }
        else if (fileType.equals("xlsx"))
        {
            wb = new XSSFWorkbook(stream);
        }

        assert wb != null;
        Sheet sheet = wb.getSheetAt(0);
        for (int i = 0; i < rowCnt; i++)
        {
            for (int j = 0; j < colCnt; j++)
            {
                try
                {
                    double val = sheet.getRow(i + startRow).getCell(j + startCol).getNumericCellValue();
                    vals[i][j] = val;
                }
                catch (IllegalStateException | NumberFormatException e)
                {
                    vals[i][j] = -1;
                    readError.add("Cannot parse to double in file "
                            + filePath + " at row " + String.valueOf(i + startRow) + " at col " + String.valueOf(j + startCol));
                }
            }
        }
        return vals;
    }

    public static void writeSheet(String outPath, double[][] res, int rowCnt, int colCnt) throws IOException
    {
        String fileType = outPath.substring(outPath.lastIndexOf(".") + 1, outPath.length());

        Workbook wb = null;
        if (fileType.equals("xls"))
        {
            wb = new HSSFWorkbook();
        }
        else if (fileType.equals("xlsx"))
        {
            wb = new XSSFWorkbook();
        }
        Sheet sheet = (Sheet) wb.createSheet("sheet1");

        for (int i = 0; i < rowCnt; i++)
        {
            Row row = (Row) sheet.createRow(i);
            for (int j = 0; j < colCnt; j++)
            {
                Cell cell = row.createCell(j);
                cell.setCellValue(res[i][j]);
            }
        }

        OutputStream stream = new FileOutputStream(outPath);
        wb.write(stream);
        stream.close();
    }
}