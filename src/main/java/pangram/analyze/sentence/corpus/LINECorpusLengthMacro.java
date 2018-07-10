package pangram.analyze.sentence.corpus;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import pangram.OutputUtil;

import java.io.File;
import java.io.IOException;

public class LINECorpusLengthMacro {
    private static String INPUT_FILE_PATH;
    private static String OUTPUT_FILE_PATH;

    static {
        INPUT_FILE_PATH = System.getenv("INPUT_FILE_PATH");
        OUTPUT_FILE_PATH = System.getenv("OUTPUT_FILE_PATH");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook book = WorkbookFactory.create(new File(INPUT_FILE_PATH));
        Sheet sheet = book.getSheetAt(0);
        LINECorpusLength.execute(sheet);

        OutputUtil.book(OUTPUT_FILE_PATH, book);
    }
}
