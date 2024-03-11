package com.example.xlsfileexample.controller;

import com.example.xlsfileexample.util.FileConversionUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@Controller
public class FileController {

    @RequestMapping("/")
    public String home() {
        return "upload";
    }

    @PostMapping("/upload")
    public String uploadFile(@RequestParam("file") MultipartFile file, Model model) {
        System.out.println("file uploaded");
        try {
            Workbook workbook = FileConversionUtil.convertPdfToXlsx(file);
            model.addAttribute("workbook", workbook);
            return "success";
        } catch (IOException e) {
            e.printStackTrace();
            model.addAttribute("error", "Error converting file: " + e.getMessage());
            return "upload";
        }
    }

//    @RequestMapping("/download")
//    public ResponseEntity<byte[]> downloadFile(Model model) throws IOException {
//        Workbook workbook = (Workbook) model.getAttribute("workbook");
//        HttpHeaders headers = new HttpHeaders();
//        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
//        headers.setContentDispositionFormData("attachment", "converted_file.xlsx");
//        byte[] workbookBytes = FileConversionUtil.workbookToBytes(workbook);
//        return new ResponseEntity<>(workbookBytes, headers, org.springframework.http.HttpStatus.OK);
//    }
@RequestMapping("/download")
public ResponseEntity<byte[]> downloadFile(Model model) throws IOException {
    Workbook workbook = (Workbook) model.getAttribute("workbook");

    HttpHeaders headers = new HttpHeaders();
    headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
    headers.setContentDispositionFormData("attachment", "converted_file.xlsx");

    byte[] workbookBytes = FileConversionUtil.workbookToBytes(workbook);

    return ResponseEntity
            .ok()
            .headers(headers)
            .body(workbookBytes);
}

}

