package inputdata.sentence;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import pangram.OutputUtil;

import java.io.File;
import java.io.IOException;

public class MatchDefaultSentenceMacro {
    private static String INPUT_DIR_PATH;
    private static String OUTPUT_DIR_PATH;

    static {
        INPUT_DIR_PATH = System.getenv("INPUT_DIR_PATH");
        OUTPUT_DIR_PATH = System.getenv("OUTPUT_DIR_PATH");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook pangramBook = WorkbookFactory.create(new File(INPUT_DIR_PATH + "pangram12.xlsx"));
        Workbook inputDataBook = WorkbookFactory.create(new File(INPUT_DIR_PATH  + "SentenceData12.xlsm"));

        SentenceMatcher.execute(pangramBook.getSheetAt(0), inputDataBook.getSheetAt(0));
        OutputUtil.book(OUTPUT_DIR_PATH + "SentenceData12_out.xlsm", inputDataBook);
    }
}
