//检查器，用来检查格式


package com.loong.ai.test;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.docx4j.TraversalUtil;
import org.docx4j.TraversalUtil.CallbackImpl;
import java.math.BigInteger;
import java.util.List;

public class FormatChecker {

    public static void main(String[] args) {
        try {
            // 加载你的DOCX文件
            String filePath = "你的文件路径.docx";
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(filePath));

            // 创建一个Callback实例来检索文档中的所有文本元素
            TextExtractor textExtractor = new TextExtractor();
            new TraversalUtil(wordMLPackage.getMainDocumentPart(), textExtractor);

            // 遍历所有的文本元素，检查“摘要”
            for (Text text : textExtractor.getTexts()) {
                if (text.getValue().contains("摘要")) {
                    R run = (R) text.getParent();
                    P paragraph = (P) run.getParent(); // 获取Run的父段落

                    // 检查格式
                    checkFontSize(run, BigInteger.valueOf(24)); // 小三号通常是24半点
                    checkFontType(run, "黑体");
                    checkAlignment(paragraph, JcEnumeration.CENTER);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void checkFontSize(R run, BigInteger expectedSize) {
        if (run.getRPr() != null && run.getRPr().getSz() != null &&
                expectedSize.equals(run.getRPr().getSz().getVal())) {
            System.out.println("‘摘要’的字号符合预期");
        } else {
            System.out.println("‘摘要’的字号不符合预期");
        }
    }

    private static void checkFontType(R run, String expectedFont) {
        if (run.getRPr() != null && run.getRPr().getRFonts() != null &&
                expectedFont.equals(run.getRPr().getRFonts().getAscii())) {
            System.out.println("‘摘要’的字体符合预期");
        } else {
            System.out.println("‘摘要’的字体不符合预期");
        }
    }

    private static void checkAlignment(P paragraph, JcEnumeration expectedAlignment) {
        if (paragraph.getPPr() != null && paragraph.getPPr().getJc() != null &&
                expectedAlignment == paragraph.getPPr().getJc().getVal()) {
            System.out.println("‘摘要’的对齐方式符合预期");
        } else {
            System.out.println("‘摘要’的对齐方式不符合预期");
        }
    }

    private static class TextExtractor extends CallbackImpl {
        private final List<Text> texts = new java.util.ArrayList<>();

        @Override
        public List<Object> apply(Object o) {
            if (o instanceof Text) {
                texts.add((Text) o);
            }
            return null;
        }

        public List<Text> getTexts() {
            return texts;
        }
    }
}
