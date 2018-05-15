package com.ma.doc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.CHPBinTable;
import org.apache.poi.hwpf.model.CharIndexTranslator;
import org.apache.poi.hwpf.model.io.HWPFFileSystem;

public class DocTest{
	
	public static void main(String []args) throws IOException {
        File file = new File("f:\\java.doc");
        try {
            FileInputStream fis = new FileInputStream(file);
            HWPFDocument doc = new HWPFDocument(fis);
            String doc1 = doc.getDocumentText();
            System.out.println(doc1);
            int zmCount=0;
            int szCount=0;
            int unCount=0;
            for(int i=0;i<doc1.length();i++){
            	char c=doc1.charAt(i);
            	if(c>='0' && c<='9'){
            		szCount++;
            	}else if(c>='a'&& c<='z' || c>='A'&&c<='Z'){
            		zmCount++;
            	}else{
            		unCount++;
            	}
            }
            FileWriter fw=null;
            fw=new FileWriter("F:\\count.txt");
            fw.write("字母统计个数为："+zmCount+"  "+"数字统计个数为："+szCount+"   "+
            		"汉字统计个数为："+unCount);
            fw.close();
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
