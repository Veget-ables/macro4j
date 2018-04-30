import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class SortMacro {
    private static Sheet mSheet;
    private static final int RANDOM_CELL = 10;

    public static void main(String[] args) throws IOException, InvalidFormatException {

//        String inputPath = args[0];
//        String outputPath = args[1];
        String inputPath = "/Users/Tojo/Desktop/files/test.xlsx";
        String outputPath = "/Users/Tojo/Desktop/files/test2.xlsx";

        Workbook book = WorkbookFactory.create(new File(inputPath));
        mSheet = book.getSheetAt(0);

        int startNextRow = 0;
        int currentRow = 0;
        while (true) {
            currentRow = sortSet(startNextRow, currentRow);

            if (mSheet.getRow(currentRow + 2) == null) break; // 空行が2行続く場合，データがないとみなし終了．
            currentRow++;
            startNextRow = currentRow;
        }

        FileOutputStream fileOut = new FileOutputStream(outputPath);
        book.write(fileOut);
    }

    // 1セットずつソートする．セットの終わりは改行で区別する．
    private static int sortSet(int startRow, int currentRow){
        while (true) {
            Row row = mSheet.getRow(currentRow);
            if (row == null) break; // 空行
            currentRow++;
        }
        randomSort(startRow, currentRow);
        return currentRow;
    }

    private static void randomSort(int startRow, int setMax) {
        for (int i = startRow; i < setMax; i++) {
            Row row = mSheet.getRow(i);
            String kana = row.getCell(1).getStringCellValue();
            String kanji = row.getCell(2).getStringCellValue();
            while (true) {
                int value = (int) (Math.random() * 100) % (setMax - startRow) + startRow;
                Cell cellKana = mSheet.getRow(value).getCell(RANDOM_CELL);
                if (cellKana != null) continue;

                cellKana = mSheet.getRow(value).createCell(RANDOM_CELL);
                cellKana.setCellValue(kana);
                Cell cellKanji = mSheet.getRow(value).createCell(RANDOM_CELL +1);
                cellKanji.setCellValue(kanji);
                break;
            }
        }
    }
}
