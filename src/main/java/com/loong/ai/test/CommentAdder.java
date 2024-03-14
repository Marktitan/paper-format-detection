package com.loong.ai.test;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.CommentsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.math.BigInteger;

public class CommentAdder {

    public static void main(String[] args) {
        try {
            // 加载你的DOCX文件
            String filePath = "你的文件路径.docx";
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(filePath));
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

            // 创建或获取CommentsPart（如果已存在）
            CommentsPart commentsPart;
            if (mainDocumentPart.getCommentsPart() == null) {
                commentsPart = new CommentsPart();
                mainDocumentPart.addTargetPart(commentsPart);
            } else {
                commentsPart = mainDocumentPart.getCommentsPart();
            }

            // 创建一个新的批注
            Comment comment = objectFactory.createComment();
            comment.setId(BigInteger.valueOf(1)); // 设置批注ID，确保每个批注的ID唯一
            comment.setAuthor("作者名"); // 设置批注作者
            comment.setDate(XmlUtils.getXmlGc(System.currentTimeMillis())); // 设置批注日期

            // 设置批注内容
            P commentParagraph = objectFactory.createP();
            R commentRun = objectFactory.createR();
            Text commentText = objectFactory.createText();
            commentText.setValue("这里是批注内容");
            commentRun.getContent().add(commentText);
            commentParagraph.getContent().add(commentRun);
            comment.getContent().add(commentParagraph);

            // 将新批注添加到CommentsPart
            commentsPart.getContents().getComment().add(comment);

            // 在文档的特定位置引用批注（例如，第一个段落）
            P paragraph = (P) mainDocumentPart.getContent().get(0); // 获取文档的第一个段落
            R run = objectFactory.createR();
            CommentRangeStart commentStart = objectFactory.createCommentRangeStart();
            commentStart.setId(comment.getId()); // 批注开始标签
            CommentRangeEnd commentEnd = objectFactory.createCommentRangeEnd();
            commentEnd.setId(comment.getId()); // 批注结束标签

            run.getContent().add(commentStart); // 添加批注开始
            run.getContent().add(objectFactory.createRCommentReference(comment.getId())); // 添加批注引用
            run.getContent().add(commentEnd); // 添加批注结束

            // 添加带有批注的文本到段落
            paragraph.getContent().add(0, run); // 将带有批注的运行添加到段落的开始处

            // 保存更改到文档
            wordMLPackage.save(new java.io.File("路径/到/你的/修改后的文档.docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
