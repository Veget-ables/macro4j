package pangram.analyze.appearance.kana;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class KanaCountMacro {
    private static String FILE_PATH;

    static {
        FILE_PATH = System.getenv("TARGET_FILE_PATH"); // 読み込むファイルを設定する
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook book = WorkbookFactory.create(new File(FILE_PATH));
        Sheet mSheet = book.getSheetAt(0);

        Map<Character, Integer> kanaMap = new HashMap<>();
        int rowNum = 1;
        while (true) {
            Row row = mSheet.getRow(rowNum);
            if (row == null) break;
            String text = row.getCell(1).getStringCellValue();
            for (Character kana : text.toCharArray()) {
                if (kanaMap.containsKey(kana)) {
                    int count = kanaMap.get(kana) + 1;
                    kanaMap.put(kana, count);
                } else {
                    kanaMap.put(kana, 0);
                }
            }
            rowNum++;
        }
        kanaMap.forEach((k, v) -> System.out.println(k + " :" + v));
    }
}
