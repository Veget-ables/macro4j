package pangram;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class OutputUtil {

    // excelファイルに保存
    public static void book(String outPath, Workbook book) {
        try {
            FileOutputStream fileOut = new FileOutputStream(outPath);
            book.write(fileOut);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
