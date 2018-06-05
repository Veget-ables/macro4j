package inputdata.sentence;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

class SentenceMatcher {
    static void execute(Sheet pangramSheet, Sheet dataSheet) {
        List<String> sentenceList = new ArrayList<>();
        StringBuilder pangrams = new StringBuilder();

        int rowNum = 0;
        while (true) {
            Row row = pangramSheet.getRow(rowNum);
            if (row == null) {
                pangrams.delete(pangrams.toString().length() - 1, pangrams.toString().length());
                sentenceList.add(pangrams.toString());
                break;
            }

            Cell cell = row.getCell(0);
            if (cell != null) {
                if (cell.getNumericCellValue() != 0) { // 文番号が存在する
                    if (!pangrams.toString().isEmpty()) {
                        pangrams.delete(pangrams.toString().length() - 1, pangrams.toString().length());
                        sentenceList.add(pangrams.toString());
                    }
                    pangrams = new StringBuilder();
                }
            }
            String text = row.getCell(1).getStringCellValue();
            pangrams.append(text).append("|");
            rowNum++;
        }
    }
}
