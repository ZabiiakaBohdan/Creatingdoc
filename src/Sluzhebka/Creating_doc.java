/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Sluzhebka;

import java.io.FileOutputStream;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Locale;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.examples.Alignment;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * @author Bogdan
 *Класс в котором описаная основная задача программы, то есть : считывание данных с файла,
 * их форматирование, запиь данных в документ формата .docx.
 */
public class Creating_doc {
    /**
     *Первое поле, используется для редактирования данных заголовка и обращения 
     */
    public String treatment;
    /**
     *Второе поле, используется редактирования данных основного текста
     */
    public String centrText;
    /** 
    * Третье поле, используется для указания сохранения
    */
    public String wave_to_saving_docx;
    /** 
    *Метод в котором мы описываем функцию, позволяющю убирать символ из строки по указанному индексу
    *@param s
    *@param pos
    *@return Возвращает результатирующую строку.
    */
    private static String removeCharAt(String s, int pos) {

       return s.substring(0,pos)+s.substring(pos+1);

    }
    /* Метод для редактирования заголовка/обращения
   
    @return Возвращает коллекцию строк.
    */
    private ArrayList<String> formate_treatment(){
         ArrayList<String> each_string = new ArrayList<String>();
         int kol_voPerenos =0;
         String help = treatment;
         for(int i = 0; i< treatment.length();i++){
            if (treatment.charAt(i) == '\n')
                kol_voPerenos++;
         }
         for(int i=0; i<kol_voPerenos;i++){
             each_string.add(help.substring(0, help.indexOf("\n")));
             help = help.replace(each_string.get(i), "");
             help = removeCharAt(help, 0);
         }
         each_string.add(help);
         return each_string;
    }
    /** 
    Метод для редактирования основной части записки
    @return Возвращает коллекцию строк.
    */
    private ArrayList<String> formate_centrText(){
         ArrayList<String> each_string = new ArrayList<String>();
        
         int kol_voPerenos =0;
         String help = centrText;
         for(int i = 0; i< centrText.length();i++){
            if (centrText.charAt(i) == '\n')
                kol_voPerenos++;
         }
         for(int i=0; i<kol_voPerenos;i++){
             each_string.add(help.substring(0, help.indexOf("\n")));
             help = help.replace(each_string.get(i), "");
             help = removeCharAt(help, 0);
         }
         each_string.add(help);
         return each_string;
    }
    /** 
    * Метод для записи всей информации в файла формата docx
    * JOptionPane класс для создания окон.
    */
    public void generate_document(){
        try{
            wave_to_saving_docx = wave_to_saving_docx+"\\Служебная записка.docx";
            FileOutputStream outStream = new FileOutputStream(wave_to_saving_docx);
            XWPFDocument doc = new XWPFDocument();
            ArrayList<XWPFParagraph> docPar = new ArrayList<XWPFParagraph>();
            ArrayList<String> treats = new ArrayList<String>();
            ArrayList<String> centr = new ArrayList<String>();
            treats = formate_treatment();
            centr = formate_centrText();
            for(int  i = 0; i<treats.size();i++){
                docPar.add(doc.createParagraph());
                docPar.get(i).setIndentationLeft(5000);
                docPar.get(i).setAlignment(ParagraphAlignment.LEFT);
                XWPFRun docRun = docPar.get(i).createRun();
                docRun.setFontFamily("Times New Roman");
                docRun.setFontSize(14);
                docRun.setText(treats.get(i));
            }
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.CENTER);
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.CENTER);
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.CENTER);
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.CENTER);
            XWPFRun docRun = docPar.get(docPar.size()-1).createRun();
            docRun.setFontFamily("Times New Roman");
            docRun.setFontSize(16);
            docRun.setText("СЛУЖЕБНАЯ ЗАПИСКА");
            for(int  i = 0; i<centr.size();i++){
                docPar.add(doc.createParagraph());
                docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.LEFT);
                XWPFRun docRun1 = docPar.get(docPar.size()-1).createRun();
                docRun1.setFontFamily("Times New Roman");
                docRun1.setFontSize(14);
                docRun1.setText(centr.get(i));
            }
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.LEFT);
            
            Date currentDate = new Date();
            Locale local = new Locale("ru","RU");
            DateFormat df = DateFormat.getDateInstance(DateFormat.DEFAULT, local); 
            currentDate = new Date(); 
            String last_str = treats.get(treats.size()-1)+"                                                                   "+df.format(currentDate);
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.LEFT);
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.LEFT);
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.LEFT);
            docPar.add(doc.createParagraph());
            docPar.get(docPar.size()-1).setAlignment(ParagraphAlignment.LEFT);
            XWPFRun docRun1 = docPar.get(docPar.size()-1).createRun();
            docRun1.setFontFamily("Times New Roman");
            docRun1.setFontSize(14);
            docRun1.setText(last_str);
            doc.write(outStream);
            outStream.close();
            JOptionPane.showMessageDialog(null, "Успешно сохранено");
            
            
        }
        catch(Exception e){
            JOptionPane.showMessageDialog(null, e);
        }
}
}
