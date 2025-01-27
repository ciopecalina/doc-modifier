package org.ciopecalina.docmodifier;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

@RestController
public class DocumentController {

    @Autowired
    private WordDocumentService wordService;

    //http://localhost:8080/download-document?invoiceSeries=ABCSDFG&invoiceNo=123455
    @GetMapping("/download-document")
    public ResponseEntity<byte[]> downloadModifiedDocument(@RequestParam String invoiceSeries, @RequestParam String invoiceNo) throws IOException {
        // Output path
        String outputPath = "src/main/resources/Invoice_" + invoiceSeries + "-" + invoiceNo + ".docx";

        // Call Service to create and modify
        String destinationPath = wordService.createAndModifyWord(invoiceSeries, invoiceNo, outputPath);

        if (destinationPath == null) {
            return ResponseEntity.status(500).body(null);
        }

        // New file for download
        File file = new File(destinationPath);
        InputStream inputStream = new FileInputStream(file);

        byte[] fileContent = inputStream.readAllBytes();

        // Response headers for file download
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + file.getName());
        headers.add(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_OCTET_STREAM_VALUE);

        // Return the file as a downloadable response
        return ResponseEntity.ok()
                .headers(headers)
                .body(fileContent);
    }
}