package pangram.analyze.sentence.bubble;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

class BubbleLength {
    private static final int SENTENCE_LENGTH_CELL = 3;
    private static final int DAY_LENGTH_SUM_CELL = 4;

    static void execute(Sheet sheet) {
        sentenceLength(sheet);
        daySentenceLength(sheet);
    }

    private static void sentenceLength(Sheet sheet) {
        int rowNum = 1;
        while (true) {
            Row row = sheet.getRow(rowNum);
            if (row == null) break;

            String text = row.getCell(1).getStringCellValue();
            row.createCell(SENTENCE_LENGTH_CELL).setCellValue(text.length());
            rowNum++;
        }
    }

    private static void daySentenceLength(Sheet sheet) {
        int rowNum = 1;
        int sumNum = 1;
        int textLength = 0;
        while (true) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                outDayLength(sheet, sumNum, textLength);
                return;
            }
            textLength += row.getCell(1).getStringCellValue().length();
            if (rowNum % 18 == 0) {
                outDayLength(sheet, sumNum, textLength);
                textLength = 0;
                sumNum++;
            }
            rowNum++;
        }
    }

    private static void outDayLength(Sheet sheet, int sumNum, int textLength){
        sheet.getRow(sumNum).createCell(DAY_LENGTH_SUM_CELL).setCellValue(textLength);
        sheet.getRow(sumNum).createCell(DAY_LENGTH_SUM_CELL + 1).setCellValue(textLength + 17);
    }
}
