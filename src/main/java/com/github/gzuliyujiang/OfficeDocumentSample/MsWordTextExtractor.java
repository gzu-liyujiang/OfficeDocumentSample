package com.github.gzuliyujiang.OfficeDocumentSample;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * [description]
 * Created by liyujiang on 2020/3/30
 */
public class MsWordTextExtractor {

    public static void main(String[] args) throws IOException {
        File file = new File(System.getProperty("user.dir"), "sample.docx");
        String text = extractWordText(file);
        System.out.println(text);
    }

    private static String extractWordText(File file) throws IOException {
        String path = file.getPath();
        FileInputStream stream = new FileInputStream(file);
        if (path.endsWith(".doc")) {
            WordExtractor extractor = new WordExtractor(stream);
            String text = extractor.getText();
            extractor.close();
            return text;
        } else if (path.endsWith("docx")) {
            XWPFDocument document = new XWPFDocument(stream);
            XWPFWordExtractor extractor = new XWPFWordExtractor(document);
            String text = extractor.getText();
            extractor.close();
            return text;
        } else {
            stream.close();
            throw new IOException("此文件不是Word文件！");
        }
    }

}
