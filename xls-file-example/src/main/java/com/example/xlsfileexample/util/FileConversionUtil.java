package com.example.xlsfileexample.util;

import org.apache.pdfbox.io.RandomAccessStreamCache;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public class FileConversionUtil {

    public static Workbook convertPdfToXlsx(MultipartFile pdfFile) throws IOException {
        try (InputStream inputStream = pdfFile.getInputStream()) {
            File tempFile = File.createTempFile("temp", ".pdf");
            pdfFile.transferTo(tempFile);

            try (PDDocument document = new PDDocument((RandomAccessStreamCache.StreamCacheCreateFunction) tempFile)) {
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("PDF Content");

                PDFTextStripper textStripper = new PDFTextStripper();
                int rowNumber = 0;

                for (int pageNumber = 0; pageNumber < document.getNumberOfPages(); pageNumber++) {
                    textStripper.setStartPage(pageNumber + 1);
                    textStripper.setEndPage(pageNumber + 1);
                    String pageText = textStripper.getText(document);

                    Row row = sheet.createRow(rowNumber++);
                    Cell cell = row.createCell(0);
                    cell.setCellValue(pageText);
                }

                return workbook;
            } finally {
                tempFile.delete();  // Clean up temporary file
            }
        }
    }

    public static byte[] workbookToBytes(Workbook workbook) throws IOException {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }
}
