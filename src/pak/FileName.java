package pak;

import java.io.File;
import java.util.ArrayList;

/**
 * Created by Satori on 2016/8/30.
 */
public class FileName
{
    //get all excel files
    public static ArrayList<String> getExcelFiles(String path)
    {
        File folder = new File(path);
        File[] listOfFiles = folder.listFiles();
        ArrayList<String> excelNames = new ArrayList<>();

        if (listOfFiles != null)
        {
            //find all excel files
            for (File file : listOfFiles)
            {
                if (file.isFile())
                {
                    String name = file.getName();
                    String sub = name.substring(name.lastIndexOf(".") + 1, name.length());
                    if (sub.equals("xls") || sub.equals("xlsx"))
                    {
                        excelNames.add(name);
                    }
                }
            }
        }
        return excelNames;
    }
}
