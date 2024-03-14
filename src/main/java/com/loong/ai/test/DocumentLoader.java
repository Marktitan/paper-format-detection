package com.loong.ai.test;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class DocumentLoader {

    /**
     * 加载一个Word文档。
     * @param filePath 文档的文件路径。
     * @return 加载的Word文档包，如果无法加载则返回null。
     */
    public static WordprocessingMLPackage loadDocument(String filePath) {
        try {
            // 使用docx4j的WordprocessingMLPackage加载文档
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.load(new java.io.File(filePath));
            return wordPackage;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public static void main(String[] args) {
        // 示例：加载文档并打印文档中的文字内容
        String docPath = "classpath:数据科学与大数据技术.docx"; // 替换为你的文档路径
        WordprocessingMLPackage wordPackage = loadDocument(docPath);
        if (wordPackage != null) {
            System.out.println("Document loaded successfully.");
            // 这里你可以添加更多的逻辑来处理加载的文档
        } else {
            System.out.println("Failed to load the document.");
        }
    }
}

