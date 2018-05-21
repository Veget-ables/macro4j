package pangram.analyze.appearance.kana;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

class KanaCounter {

    static Map<String, Integer> execute(Sheet sheet) {
        Map<String, Integer> charMap = countCharMap(sheet);
        return countKanaMap(charMap);
    }

    private static Map<String, Integer> countCharMap(Sheet sheet){
        Map<String, Integer> countMap = new HashMap<>();
        int rowNum = 1;
        while (true) {
            Row row = sheet.getRow(rowNum);
            if (row == null) break;
            String text = row.getCell(1).getStringCellValue();
            for (char kana : text.toCharArray()) {
                String k = String.valueOf(kana);
                if (countMap.containsKey(k)) {
                    int count = countMap.get(k) + 1;
                    countMap.put(k, count);
                } else {
                    countMap.put(k, 1);
                }
            }
            rowNum++;
        }
        return countMap;
    }

    private static Map<String, Integer> countKanaMap(Map<String, Integer> map) {
        Map<String, Integer> countKanaMap = new LinkedHashMap<>();
        for (KanaType kanaType : KanaType.values()) {
            for (int i = 0; i < kanaType.getKana().length; i++) {
                String key = "";
                int count = 0;
                String kana = kanaType.getKana()[i];
                key = kana;
                if (kana.isEmpty()) continue;
                if (map.containsKey(kana)) count = map.get(kana);

                if (kanaType.getSecondKana() != null) {
                    String secondKana = kanaType.getSecondKana()[i];
                    if (map.containsKey(secondKana)) count += map.get(secondKana);
                }
                if (kanaType.getThirdKana() != null) {
                    String thirdKana = kanaType.getThirdKana()[i];
                    if (map.containsKey(thirdKana)) count += map.get(thirdKana);

                }
                countKanaMap.put(key, count);
            }
        }
        return countKanaMap;
    }
}
