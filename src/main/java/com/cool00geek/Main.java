package com.cool00geek;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.Scanner;

public class Main {
    public static void main(String[] args){
        StringBuilder markdownFile = new StringBuilder();

        String filename = "";
        Scanner scanner = new Scanner(System.in);
        System.out.print("Please enter the absolute filepath to the pptx: ");
        filename = scanner.nextLine();

        File file = new File(filename);
        FileInputStream inputstream = null;
        try {
            inputstream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.exit(1);
        }
        XMLSlideShow ppt = null;
        try {
            ppt = new XMLSlideShow(inputstream);
        } catch (IOException e) {
            e.printStackTrace();
            System.exit(1);
        }

        List<XSLFSlide> slides = ppt.getSlides();
        for (int index = 0; index < slides.size(); index++) {
            XSLFSlide slide = slides.get(index);
            XSLFTextShape[] placeHolders = slide.getPlaceholders();
            for (int i = 0; i < placeHolders.length; i++) {
                if (i == 0) {
                    if (index == 0) {
                        markdownFile.append("# ");
                    } else {
                        markdownFile.append("## ");
                    }
                }
                String currentText = placeHolders[i].getText();
                String[] lines = currentText.split("\n");
                if (lines.length == 1) {
                    if (index == 0 && i != 0){
                        markdownFile.append("#### ");
                    }
                    markdownFile.append(lines[0]).append("\n\n");
                } else {
                    for (String line : lines) {
                        markdownFile.append("- ").append(line).append("\n\n");
                    }
                }
            }
        }
        System.out.println(markdownFile);
    }
}
