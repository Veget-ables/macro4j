package pangram.analyze.appearance.sentence;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

class MeasureLength {
    private static final String SPLIT_REGEX = " |\\n|\\?|!|？|！|。、*|w+|わらい、*|（わらい）、*";

    private static final int LENGTH_CELL = 20;
    private static final int LENGTH_AVERAGE_CELL = 21;

    static void execute(Sheet sheet) {
        int rowNum = 1;
        int lenRowNum = 1;
        while (true) {
            Row row = sheet.getRow(rowNum);
            if (row == null) break;
            if (row.getCell(0) == null) break;

            String text = row.getCell(0).getStringCellValue();
            String[] sentences = text.split(SPLIT_REGEX, 0);
            lenRowNum = countLength(sheet, sentences, lenRowNum);
            rowNum++;
        }
    }

    private static int countLength(Sheet sheet, String[] sentences, int lenRowNum) {
        for (String sentence : sentences) {
            if (sentence.isEmpty() | sentence.length() == 1) continue;

            Row lenRow = sheet.getRow(lenRowNum);
            if (lenRow == null) {
                lenRow = sheet.createRow(lenRowNum);
            }
            lenRow.createCell(LENGTH_CELL).setCellValue(sentence);
            lenRow.createCell(LENGTH_AVERAGE_CELL).setCellValue(sentence.length());
            lenRowNum++;
        }
        return lenRowNum;
    }
}
