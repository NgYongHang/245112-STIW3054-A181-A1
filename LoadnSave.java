/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.assignment1;

import static com.mycompany.assignment1.StoreToExcel.excelLog;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 *
 * @author User
 */
public class LoadnSave {

    /**
     *
     * @param args
     */
    public static void main(String[] args)  {
        String html = "https://ms.wikipedia.org/wiki/Malaysia";
        try {
            load(html);
        } catch (IOException ex) {
            Logger.getLogger(LoadnSave.class.getName()).log(Level.SEVERE, null, ex);
        }
 
    }
    public static void load(String html) throws IOException{
        Document web = Jsoup.connect(html).get();
            Element table = web.select("table.wikitable").get(1);
//            System.out.println(table.toString());
            Elements row = table.select("tr");
//            System.out.println(row.size());
            System.out.println("Data storing....");
            excelLog("Trivia", "", 0);
            for (int i = 0; i < row.size(); i++) {
                Elements column1 = row.get(i).select("th");
//                System.out.println(column1.text());
                Elements column2 = row.get(i).select("td");
//                System.out.println(column2.text());

                //apply store to excel class
                excelLog(column1.text(), column2.text(), i + 1);
            }
            System.out.println("Data successfully stored at C:\\Users\\User\\Desktop\\Trivia.xls");
    }
}