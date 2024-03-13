package com.example.xlsfileexample.controller;

import org.apache.pdfbox.cos.COSDictionary;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;

@Controller
public class FileConversionController {

    @RequestMapping("/upload")
    public String uploadForm() {
        return "upload";
    }

    @PostMapping("/upload")
    public String uploadFile(@RequestParam("file") MultipartFile file, Model model) {
        try {

            File pdfFile = convertXlsxToPdf(file);


            model.addAttribute("pdfFileName", pdfFile.getName());

            return "success";
        } catch (IOException e) {
            e.printStackTrace();
            model.addAttribute("error", "Error converting file: " + e.getMessage());
            return "upload";
        }
    }

    private File convertXlsxToPdf(MultipartFile xlsxFile) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(xlsxFile.getInputStream())) {
            File pdfFile = File.createTempFile("convertedFile", ".pdf");

            try (PDDocument document = new PDDocument()) {
                PDPage page = new PDPage();
                document.addPage(page);

                try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
                    // Error in (new PDType1Font(new COSDictionary(new COSDictionary()))
                    contentStream.setFont(new PDType1Font(new COSDictionary(new COSDictionary())), 12);
                    contentStream.beginText();
                    contentStream.newLineAtOffset(100, 700);

                    for (Sheet sheet : workbook) {
                        for (Row row : sheet) {
                            for (Cell cell : row) {
                                contentStream.showText(cell.toString() + "\t");
                            }
                            contentStream.newLine();
                        }
                    }

                    contentStream.endText();
                }

                document.save(pdfFile);
            }

            return pdfFile;
        }
    }
}
