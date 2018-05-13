package pangram.print;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

class PrintWord {
    private static XWPFDocument mDocument;

    static void createPangramList(String absPath, Workbook book) {
        try (FileOutputStream out = new FileOutputStream(new File(absPath))) {
            mDocument = new XWPFDocument();
            addPangrams(book);
            mDocument.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static XWPFDocument addPangrams(Workbook book) {
        Sheet sheet = book.getSheetAt(0);
        int wordId = 1;
        int rowNum = 0;
        int day = 1;

        while (true) {
            Row row = sheet.getRow(rowNum);
            if (row == null) return mDocument;

            if (wordId == 1) {
                mDocument.createParagraph().setPageBreak(true);
                insertDayTitle(day++);
            }
            insertText(row, wordId);

            if (wordId == 18) {
                mDocument.createParagraph().createRun().addBreak();
                insertTable();
                wordId = 1;
            } else {
                wordId++;
            }
            rowNum++;
        }
    }

    private static void insertDayTitle(int day) {
        XWPFParagraph paragraph = mDocument.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun runTitle = paragraph.createRun();
        runTitle.setFontSize(24);
        runTitle.setText("計測" + day + "日目");
    }

    private static void insertText(Row row, int wordId) {
        XWPFParagraph paragraph = mDocument.createParagraph();
        paragraph.setSpacingBeforeLines(35);
        XWPFRun run = paragraph.createRun();

        String format = wordId < 10 ? "%s.    %s (%s) \n" : "%s.  %s (%s) \n";
        String kana = row.getCell(1).getStringCellValue();
        String kanji = row.getCell(2).getStringCellValue();
        run.setText(String.format(format, wordId, kana, kanji));
    }

    private static void insertTable() {
        XWPFTable table = mDocument.createTable(2, 1);
        XWPFTableCell timeCell = table.getRow(0).getCell(0);
        setCellWidth(timeCell);
        XWPFParagraph timeParagraph = timeCell.getParagraphs().get(0);
        XWPFRun run = timeParagraph.createRun();
        run.setText("開始時刻:");

        XWPFTableRow row = table.getRow(1);
        row.setHeight(3000);
        XWPFTableCell commentCell = row.getCell(0);
        setCellWidth(commentCell);
    }

    private static void setCellWidth(XWPFTableCell cell) {
        cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(10500L));
    }
}
