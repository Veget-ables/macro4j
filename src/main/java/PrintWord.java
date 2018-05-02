import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPageMargins;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

class PrintWord {

    static void createWord(String absPath, Workbook book) {
        try (FileOutputStream out = new FileOutputStream(new File(absPath))) {
            XWPFDocument document = addPangramList(book);
            document.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static XWPFDocument addPangramList(Workbook book) {
        Sheet sheet = book.getSheetAt(0);
        XWPFDocument document = new XWPFDocument();

        int wordId = 1;
        int rowNum =0;
        while (true) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            Row row = sheet.getRow(rowNum);
            if (row == null) return document;

            String kana = row.getCell(1).getStringCellValue();
            String kanji = row.getCell(2).getStringCellValue();
            String text;
            if (wordId < 10) {
                text = String.format("%s.    %s (%s) \n", wordId, kana, kanji);
            } else {
                text = String.format("%s.  %s (%s) \n", wordId, kana, kanji);
            }
            run.setText(text);

            if(wordId == 18){
                document.createParagraph().setPageBreak(true);
                wordId = 1;
            }else{
                wordId++;
            }
            rowNum++;
        }
    }
}
