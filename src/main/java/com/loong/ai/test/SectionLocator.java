package com.loong.ai.test;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.Text;

import java.util.ArrayList;
import java.util.List;
import java.util.HashMap;
import java.util.Map;

public class SectionLocator {

    // 定位并返回文档中的各个部分
    public static Map<String, List<P>> findSections(WordprocessingMLPackage wordMLPackage, List<String> sectionMarkers) {
        Map<String, List<P>> sections = new HashMap<>();
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

        // 获取文档中的所有段落
        List<Object> paragraphs = getAllElementFromObject(mainDocumentPart, P.class);

        // 当前正在处理的部分的名称
        String currentSectionName = null;
        // 当前部分的内容
        List<P> currentSectionContent = new ArrayList<>();

        // 遍历所有段落，寻找含有特定标记的段落
        for (Object obj : paragraphs) {
            P paragraph = (P) obj;
            String paragraphText = getParagraphText(paragraph).trim();

            // 检查段落是否包含某个标记
            if (sectionMarkers.contains(paragraphText)) {
                // 如果当前部分名称不为空，保存并重置当前部分
                if (currentSectionName != null) {
                    sections.put(currentSectionName, new ArrayList<>(currentSectionContent));
                    currentSectionContent.clear();
                }
                // 更新当前部分的名称
                currentSectionName = paragraphText;
            } else if (currentSectionName != null) {
                // 如果当前正在处理某个部分，添加当前段落到这个部分
                currentSectionContent.add(paragraph);
            }
        }

        // 添加最后一个部分
        if (currentSectionName != null && !currentSectionContent.isEmpty()) {
            sections.put(currentSectionName, currentSectionContent);
        }

        return sections;
    }

    // 从对象（如段落）中提取文本内容
    private static String getParagraphText(P paragraph) {
        StringBuilder sb = new StringBuilder();
        List<Object> texts = getAllElementFromObject(paragraph, Text.class);
        for (Object text : texts) {
            Text t = (Text) text;
            sb.append(t.getValue());
        }
        return sb.toString();
    }

    // 通用方法，用于从给定对象中获取所有指定类型的元素
    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }
}
