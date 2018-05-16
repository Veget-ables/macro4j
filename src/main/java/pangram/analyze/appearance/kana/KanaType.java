package pangram.analyze.appearance.kana;

enum KanaType {

    KANA_A(
            new String[]{"あ", "い", "う", "え", "お"},
            new String[]{"ぁ", "ぃ", "ぅ", "ぇ", "ぉ"}),

    KANA_K(
            new String[]{"か", "き", "く", "け", "こ"},
            new String[]{"が", "ぎ", "ぐ", "げ", "ご"}),

    KANA_S(
            new String[]{"さ", "し", "す", "せ", "そ"},
            new String[]{"ざ", "じ", "ず", "ぜ", "ぞ"}),

    KANA_T(
            new String[]{"た", "ち", "つ", "て", "と"},
            new String[]{"だ", "ぢ", "っ", "で", "ど"},
            new String[]{"", "", "づ", "", ""}),

    KANA_N(
            new String[]{"な", "に", "ぬ", "ね", "の"}),

    KANA_H(
            new String[]{"は", "ひ", "ふ", "へ", "ほ"},
            new String[]{"ば", "び", "ぶ", "べ", "ぼ"},
            new String[]{"ぱ", "ぴ", "ぷ", "ぺ", "ぽ"}),

    KANA_M(
            new String[]{"ま", "み", "む", "め", "も"}),

    KANA_Y(
            new String[]{"や", "ゆ", "よ", "", ""},
            new String[]{"ゃ", "ゅ", "ょ", "", ""}),

    KANA_R(
            new String[]{"ら", "り", "る", "れ", "ろ"}),

    KANA_W(
            new String[]{"わ", "を", "ん", "", ""},
            new String[]{"ゎ", "", "", "", ""}),

    KANA_MARK(
            new String[]{"", "", "", "。", "、"});

    private String[] kana;
    private String[] secondKana;
    private String[] thirdKana;

    KanaType(String[] kana) {
        this(kana, null, null);
    }

    KanaType(String[] kana, String[] secondKana) {
        this(kana, secondKana, null);
    }

    KanaType(String[] kana, String[] secondKana, String[] thirdKana) {
        this.kana = kana;
        this.secondKana = secondKana;
        this.thirdKana = thirdKana;
    }

    public String[] getKana() {
        return kana;
    }

    public String[] getSecondKana() {
        return secondKana;
    }

    public String[] getThirdKana() {
        return thirdKana;
    }
}

