package com.loong.ai.test;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import java.util.List;
import java.util.Map;

public class CheckDocument {

    public static void main(String[] args) {
        if (args.length < 1) {
            System.out.println("Usage: CheckDocument <path_to_docx>");
            return;
        }
        String filePath = args[0]; // 从命令行参数获取文档路径

        try {
            // 使用DocumentLoader加载文档
            WordprocessingMLPackage wordMLPackage = DocumentLoader.loadDocument(filePath);
            if (wordMLPackage == null) {
                System.out.println("Failed to load the document.");
                return;
            }

            // 定义文档中部分的标记
            List<String> sectionMarkers = List.of("封面如下", "摘要如下", "正文如下", "引用如下", "致谢如下");

            // 使用SectionLocator定位文档的不同部分
            Map<String, List<P>> sections = SectionLocator.findSections(wordMLPackage, sectionMarkers);

            // 使用ContentClassifier分类文档内容
            ContentClassifier.classifyContent(wordMLPackage);

            // 假设FormatChecker和CommentAdder的逻辑已经实现
            // 在这里调用它们对文档进行格式检查和添加批注
            // 例如，对每个部分执行特定的格式检查
            sections.forEach((sectionName, paragraphs) -> {
                System.out.println("Processing section: " + sectionName);
                // 对于每个段落，进行格式检查和添加批注
                // 这里需要根据FormatChecker和CommentAdder的实际实现来调用它们的方法
                // 示例：FormatChecker.checkParagraphFormat(paragraph);
                // 示例：CommentAdder.addCommentToParagraph(wordMLPackage, paragraph, "批注内容");
            });

            // 保存更改后的文档
            wordMLPackage.save(new java.io.File(filePath.replace(".docx", "_modified.docx")));

            System.out.println("Document processing completed.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
