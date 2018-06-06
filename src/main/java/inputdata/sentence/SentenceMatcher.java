package inputdata.sentence;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

class SentenceMatcher {
    static void execute(Sheet pangramSheet, Sheet dataSheet) {
        StringBuilder pangrams = new StringBuilder();

        int rowNum = 0;
        int day = 1;
        while (true) {
            Row row = pangramSheet.getRow(rowNum);
            if (row == null) {
                pangrams.delete(pangrams.toString().length() - 1, pangrams.toString().length());
                break;
            }

            if (rowNum != 0 && rowNum % 18 == 0) {
                pangrams.delete(pangrams.toString().length() - 1, pangrams.toString().length());

                Row rowDay = dataSheet.getRow(day);
                Cell inputCell = rowDay.getCell(4);
                if (inputCell != null) {
                    String inputText = inputCell.getStringCellValue();
                    char[] inputChars = inputText.toCharArray();
                    char[] pangramChars = pangrams.toString().toCharArray();
                    int minLength = inputChars.length < pangramChars.length ? inputChars.length : pangramChars.length;
                    for (int i = 0; i < minLength; i++) {
                        if (inputChars[i] == pangramChars[i]) continue;
                        rowDay.createCell(6).setCellValue(inputText.substring(0, i));
                        break;
                    }
                }
                day++;
                pangrams = new StringBuilder();
            }
            String text = row.getCell(1).getStringCellValue();
            pangrams.append(text).append("l");
            rowNum++;
        }
    }
}
