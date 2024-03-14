package com.loong.ai.test;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Text;
import java.util.List;

public class CheckDocument {

    public static void main(String[] args) {
        String filePath = "classpath:数据科学与大数据技术.docx"; // 确保路径是正确的
        try {
            // 使用DocumentLoader加载文档
            WordprocessingMLPackage wordMLPackage = DocumentLoader.loadDocument(filePath);
            if (wordMLPackage == null) {
                System.out.println("Failed to load the document.");
                return;
            }

            // 获取文档的所有文本元素
            List<Object> texts = getAllElementFromObject(wordMLPackage.getMainDocumentPart(), Text.class);

            int chineseCount = 0;
            int englishCount = 0;
            StringBuilder abstractText = new StringBuilder();
            boolean foundAbstract = false;

            for (Object obj : texts) {
                Text textElement = (Text) obj;
                String text = textElement.getValue();
                for (char c : text.toCharArray()) {
                    if (isChinese(c)) {
                        chineseCount++;
                    } else if (isEnglish(c)) {
                        englishCount++;
                    }
                }

                // 提取摘要
                if (text.startsWith("摘要")) {
                    foundAbstract = true;
                } else if (foundAbstract) {
                    abstractText.append(text).append("\n");
                }
            }

            System.out.println("中文字符数：" + chineseCount);
            System.out.println("英文字符数：" + englishCount);
            System.out.println("摘要内容：\n" + abstractText.toString().trim());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static boolean isChinese(char c) {
        // 判断字符是否为中文字符（Unicode 编码范围）
        return (c >= 0x4E00 && c <= 0x9FA5);
    }

    public static boolean isEnglish(char c) {
        // 判断字符是否为英文字符（ASCII 编码范围）
        return (c >= 0x0020 && c <= 0x007E);
    }

    // 通用方法，用于从给定对象中获取所有指定类型的元素
    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        // 此处应包含之前实现的getAllElementFromObject方法的代码
    }
}
