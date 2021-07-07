package com.duck.docxtopdf;


import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.ResourceUtils;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.*;

@RestController
public class DuckController {

    @PostMapping
    public ResponseEntity<InputStreamResource> download(@RequestParam("file") MultipartFile multipartFile, HttpServletRequest request) throws IOException {
        String fileName = StringUtils.cleanPath(multipartFile.getOriginalFilename().replace(".docx",""));
        System.out.println(fileName);
        InputStream in = new ByteArrayInputStream(multipartFile.getBytes());
        XWPFDocument document = new XWPFDocument(in);
        PdfOptions options = PdfOptions.create();
        File fileOut = new File(fileName+".pdf");
        OutputStream out = new FileOutputStream(fileOut);
        PdfConverter.getInstance().convert(document, out, options);

        document.close();
        out.close();


        HttpHeaders responseHeader = new HttpHeaders();
        try {
            File file = fileOut;
            byte[] data = FileUtils.readFileToByteArray(file);
            // Set mimeType trả về
            responseHeader.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            // Thiết lập thông tin trả về
            responseHeader.set("Content-disposition", "attachment; filename=" + file.getName());
            responseHeader.setContentLength(data.length);
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(data));
            InputStreamResource inputStreamResource = new InputStreamResource(inputStream);
            return new ResponseEntity<InputStreamResource>(inputStreamResource, responseHeader, HttpStatus.OK);
        } catch (Exception ex) {
            return new ResponseEntity<InputStreamResource>(null, responseHeader, HttpStatus.INTERNAL_SERVER_ERROR);
        } finally {
            fileOut.delete();
        }
    }
}
