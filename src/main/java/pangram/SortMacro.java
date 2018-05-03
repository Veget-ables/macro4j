package pangram;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class SortMacro {
    private static final String INPUT_DIR = "/Users/Tojo/Desktop/files/";
    private static final String OUTPUT_DIR = "/Users/Tojo/Desktop/files/";

    private static Sheet mSheet;
    private static final int RANDOM_CELL = 10;

    public static void main(String[] args) {
        for (int i = 1; i < 13; i++) {
            try {
                String fileName = "pangram";
                Workbook book = randomSortXSLX(INPUT_DIR + fileName + ".xlsx");
                FileOutputStream fileOut = new FileOutputStream(OUTPUT_DIR + fileName + i + ".xlsx");
                book.write(fileOut);

                PrintWord.createWord(OUTPUT_DIR + fileName + i + ".docx", book);
            } catch (IOException | InvalidFormatException e) {
                throw new RuntimeException(e);
            }
        }
    }

    static private Workbook randomSortXSLX(String fileName) throws IOException, InvalidFormatException {
        Workbook book = WorkbookFactory.create(new File(fileName));
        mSheet = book.getSheetAt(0);

        int startNextRow = 0;
        int currentRow = 0;
        while (true) {
            currentRow = sortSet(startNextRow, currentRow + 1);
            if (currentRow == -1) break; // 空行の場合，データがないとみなし終了．
            startNextRow = currentRow;
        }
        return book;
    }

    // 1セットずつソートする．セットの終わりは改行で区別する．
    private static int sortSet(int startRow, int currentRow) {
        while (true) {
            Row row = mSheet.getRow(currentRow);
            if (row == null) { // 最後の空行がきた．
                randomSort(startRow, currentRow);
                return -1;
            }
            Cell cell = row.getCell(0);
            if (cell == null) {
                currentRow++;
            } else {
                if (cell.getNumericCellValue() == 0) { // 意図しない空文字列が含まれている場合がある．
                    currentRow++;
                    continue;
                }
                break;
            }
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
                Cell cellKanji = mSheet.getRow(value).createCell(RANDOM_CELL + 1);
                cellKanji.setCellValue(kanji);
                break;
            }
        }
        removeOldCells(startRow, setMax);
    }

    // ソート後の値を古い列に上書きする．
    private static void removeOldCells(int startRow, int setMax) {
        for (int i = startRow; i < setMax; i++) {
            Row row = mSheet.getRow(i);
            Cell kanaCell = row.getCell(RANDOM_CELL);
            Cell kanjiCell = row.getCell(RANDOM_CELL + 1);
            row.getCell(1).setCellValue(kanaCell.getStringCellValue());
            row.getCell(2).setCellValue(kanjiCell.getStringCellValue());
            kanaCell.removeCellComment();
            kanjiCell.removeCellComment();
        }
    }
}
