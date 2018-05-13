package pangram.analyze.appearance.kana;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

public class KanaCountMacro {
    private static String INPUT_FILE_PATH;
    private static String OUTPUT_FILE_PATH;
    private static final int KANA_CELL = 20;
    private static final int KANA_COUNT_CELL = 21;
    private static final int KANA_APPEARANCE_RATE_CELL = 22;

    static {
        INPUT_FILE_PATH = System.getenv("INPUT_FILE_PATH");
        OUTPUT_FILE_PATH = System.getenv("OUTPUT_FILE_PATH");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook book = WorkbookFactory.create(new File(INPUT_FILE_PATH));
        Sheet sheet = book.getSheetAt(0);
        Map<String, Integer> kanaMap = KanaCounter.execute(sheet);
        kanaMap.forEach((k, v) -> System.out.println(k + " :" + v));

        // 出現割合を算出するために，合計文字数を計算
        int countAll =0;
        for(String key : kanaMap.keySet()){
            if(key.equals("、") || key.equals("。"))continue;
            countAll += kanaMap.get(key);
        }

        int rowNum=1;
        for(String key : kanaMap.keySet()){
            Row row = sheet.getRow(rowNum++);
            row.createCell(KANA_CELL).setCellValue(key);

            int count = kanaMap.get(key);
            row.createCell(KANA_COUNT_CELL).setCellValue(count);
            row.createCell(KANA_APPEARANCE_RATE_CELL).setCellValue((double) count/countAll);
        }

        FileOutputStream fileOut = new FileOutputStream(OUTPUT_FILE_PATH);
        book.write(fileOut);
    }
}
