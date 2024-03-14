package com.loong.ai.test;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import java.util.List;

public class ContentClassifier {

    public static void classifyContent(WordprocessingMLPackage wordMLPackage) {
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

        // 获取文档中的所有对象
        List<Object> contents = mainDocumentPart.getContent();

        for (Object content : contents) {
            // 检查内容是否为段落
            if (content instanceof P) {
                P paragraph = (P) content;
                classifyParagraph(paragraph);
            }
            // 其他类型的内容可以在这里添加更多的检查
        }
    }

    private static void classifyParagraph(P paragraph) {
        // 检查段落是否为标题
        if (isHeading(paragraph)) {
            System.out.println("Found a heading: " + getParagraphText(paragraph));
        } else if (containsImage(paragraph)) {
            System.out.println("Found an image within a paragraph.");
        } else {
            System.out.println("Regular paragraph: " + getParagraphText(paragraph));
        }
        // 添加对公式等其他内容的检查
    }

    private static boolean isHeading(P paragraph) {
        // 通常标题使用特定的样式，如Heading1, Heading2等
        if (paragraph.getPPr() != null && paragraph.getPPr().getPStyle() != null) {
            String styleId = paragraph.getPPr().getPStyle().getVal();
            return styleId != null && styleId.startsWith("Heading");
        }
        return false;
    }

    private static boolean containsImage(P paragraph) {
        // 检查段落中是否包含图像
        return paragraph.getContent().stream().anyMatch(element -> element instanceof R && ((R) element).getContent().stream().anyMatch(subElement -> subElement instanceof Drawing));
    }

    private static String getParagraphText(P paragraph) {
        // 提取段落中的文本
        StringBuilder sb = new StringBuilder();
        List<Object> texts = getAllElementFromObject(paragraph, Text.class);
        for (Object text : texts) {
            Text t = (Text) text;
            sb.append(t.getValue());
        }
        return sb.toString();
    }

    // getAllElementFromObject 方法与之前相同
    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        // 实现细节与之前相同，用于提取指定类型的所有元素
    }
}

