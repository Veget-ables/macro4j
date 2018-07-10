package pangram.analyze.sentence.bubble;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import pangram.OutputUtil;

import java.io.File;
import java.io.IOException;

public class BubbleLengthMacro {
    private static String INPUT_FILE_PATH;
    private static String OUTPUT_FILE_PATH;

    static {
        INPUT_FILE_PATH = System.getenv("INPUT_DIR_PATH");
        OUTPUT_FILE_PATH = System.getenv("OUTPUT_DIR_PATH");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        for(int i=1; i< 37; i++){
            Workbook book = WorkbookFactory.create(new File(INPUT_FILE_PATH +i+ "/pangram"+ i + ".xlsx"));
            Sheet sheet = book.getSheetAt(0);
            BubbleLength.execute(sheet);

            OutputUtil.book(OUTPUT_FILE_PATH +i+ "/pangram"+ i + "_out.xlsx", book);
        }
    }
}
