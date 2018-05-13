package pangram.analyze;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;

public class ChunkCountMacro {
    private static String REQUEST_URL;
    private static String FILE_PATH;

    static {
        FILE_PATH = System.getenv("TARGET_FILE_PATH"); // 読み込むファイルを設定する
        String APP_ID = System.getenv("YAHOO_APP_ID"); // 登録したAPP_IDを設定する
        REQUEST_URL = "https://jlp.yahooapis.jp/DAService/V1/parse?appid=" + APP_ID + "&sentence=";
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook book = WorkbookFactory.create(new File(FILE_PATH));
        Sheet sheet = book.getSheetAt(0);

        int rowNum = 1;
        while (true) {
            Row row = sheet.getRow(rowNum);
            if (row == null) break;

            String kanji = row.getCell(2).getStringCellValue();
            String encodedText = URLEncoder.encode(kanji, "UTF-8");
            HttpURLConnection conn = (HttpURLConnection) new URL(REQUEST_URL + encodedText).openConnection();
            InputStream in = conn.getInputStream();
            ChunkCounter.build(in);

            rowNum++;
        }
        System.out.println("Chunk:" + ChunkCounter.getCountChunk());
    }
}
