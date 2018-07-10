package pangram.print;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import pangram.OutputUtil;

import java.io.File;
import java.io.IOException;

public class RandomSortMacro {
    private static final String INPUT_FILE_PATH;
    private static final int RANDOM_KANA_CELL = 100;
    private static final int RANDOM_KANJI_CELL = 101;

    private static Sheet mSheet;
    private static boolean isFinish = false;

    static {
        INPUT_FILE_PATH = System.getenv("INPUT_FILE_PATH");
    }

    public static void main(String[] args) {
        final String outputFile = "/Users/Tojo/Desktop/files/pangram";

        for (int i = 13; i < 37; i++) { // 12セット分のランダムファイルを作成
            try {
                Workbook book = randomSortXSLX(INPUT_FILE_PATH);
                OutputUtil.book(outputFile + i + ".xlsx", book);

                PrintWord.createPangramList(outputFile + i + ".docx", book);
            } catch (IOException | InvalidFormatException e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * 1セットごとにソートする．
     *
     * @param fileName ソート対象のファイル名．
     */
    static private Workbook randomSortXSLX(String fileName) throws IOException, InvalidFormatException {
        Workbook book = WorkbookFactory.create(new File(fileName));
        mSheet = book.getSheetAt(0);
        isFinish = false;

        int startNextRow = 0;
        int currentRow = 0;
        while (true) {
            currentRow = countSetSentence(currentRow + 1);
            randomSortRange(startNextRow, currentRow);
            if (isFinish) break; // 空行の場合，データがないとみなし終了．
            startNextRow = currentRow;
        }
        return book;
    }

    /**
     * セットに含まれる短文数をカウント．セットの区切りは0番目のCellの番号で判断．
     */
    private static int countSetSentence(int currentRow) {
        while (true) {
            Row row = mSheet.getRow(currentRow);
            if (row == null) { // 最後の空行がきた．
                isFinish = true;
                return currentRow;
            }
            Cell cell = row.getCell(0); // 1番目のcellの値で日数を区別．
            if (cell != null) {
                if (cell.getNumericCellValue() == 0) { // 意図しない空文字列が含まれている場合がある．
                    currentRow++;
                    continue;
                }
                break;
            }
            currentRow++;
        }
        return currentRow;
    }

    private static void randomSortRange(int start, int end) {
        for (int i = start; i < end; i++) {
            Row row = mSheet.getRow(i);
            String kana = row.getCell(1).getStringCellValue();
            String kanji = row.getCell(2).getStringCellValue();
            while (true) {
                int value = (int) (Math.random() * 100) % (end - start) + start;
                Cell cellKana = mSheet.getRow(value).getCell(RANDOM_KANA_CELL);
                if (cellKana != null) continue;

                cellKana = mSheet.getRow(value).createCell(RANDOM_KANA_CELL);
                cellKana.setCellValue(kana);
                Cell cellKanji = mSheet.getRow(value).createCell(RANDOM_KANJI_CELL);
                cellKanji.setCellValue(kanji);
                break;
            }
        }
        removeOldCells(start, end);
    }

    /**
     * ソート後の値を古い列に上書きする．
     */
    private static void removeOldCells(int startRow, int setMax) {
        for (int i = startRow; i < setMax; i++) {
            Row row = mSheet.getRow(i);
            Cell kanaCell = row.getCell(RANDOM_KANA_CELL);
            Cell kanjiCell = row.getCell(RANDOM_KANJI_CELL);
            row.getCell(1).setCellValue(kanaCell.getStringCellValue());
            row.getCell(2).setCellValue(kanjiCell.getStringCellValue());
            kanaCell.removeCellComment();
            kanjiCell.removeCellComment();
        }
    }
}
