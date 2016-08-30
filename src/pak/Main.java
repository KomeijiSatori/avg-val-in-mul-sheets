package pak;

import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Created by Satori on 2016/8/30.
 */

public class Main
{
    final static String inPath = "./input/";//folder path
    final static String outPath = "./result/result.xlsx";//output file
    final static String errPath = "./result/error.txt";
    final static int rowCnt = 29;
    final static int colCnt = 5;
    final static int startRow = 2;
    final static int startCol = 3;

    public static void main(String [] args)
    {
        ArrayList<String> fileNames = FileName.getExcelFiles(inPath);
        double[][] res = new double[rowCnt][colCnt];
        ArrayList<String> errList = new ArrayList<>();

        for (String fileName : fileNames)
        {
            try
            {
                double[][] vals = SheetUtils.readSheet(inPath + fileName, rowCnt, colCnt, startRow, startCol);
                for (int i = 0; i < rowCnt; i++)
                {
                    for (int j = 0; j < colCnt; j++)
                    {
                        res[i][j] += vals[i][j];
                    }
                }
            }
            catch (IOException e1)
            {
                errList.add("IO error at file: " + fileName);
            }
        }
        errList.addAll(SheetUtils.readError);

        if (!fileNames.isEmpty() && errList.isEmpty())
        {
            for (int i = 0; i < rowCnt; i++)
            {
                for (int j = 0; j < colCnt; j++)
                {
                    res[i][j] /= fileNames.size();
                }
            }
            //write the result
            try
            {
                SheetUtils.writeSheet(outPath, res, rowCnt, colCnt);
            }
            catch (IOException e)
            {
                errList.add("Output file error");
            }
        }

        if (!errList.isEmpty())
        {
            try
            {
                FileWriter writer = new FileWriter(errPath);
                for (String str : errList)
                {
                    writer.write(str);
                }
                writer.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }
}
