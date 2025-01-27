package org.ciopecalina.docmodifier;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;

@Service
public class WordDocumentService {

    public String createAndModifyWord(String invoiceSeries, String invoiceNo, String outputPath) {
        String templatePath = "src/main/resources/template/invoice-template.docx";

        try (InputStream resourceStream = new FileInputStream(templatePath)) {
            // Load the word document
            XWPFDocument document = new XWPFDocument(resourceStream);

            // Invoice data
            findAndReplaceText(document, "<s>", invoiceSeries);
            findAndReplaceText(document, "<no>", invoiceNo);
            findAndReplaceText(document, "<d>", "10/12/2024");

            // User data
            findAndReplaceText(document, "<u_name>", "-");
            findAndReplaceText(document, "<u_reg_no>", "-");
            findAndReplaceText(document, "<u_f_code", "-");
            findAndReplaceText(document, "<u_address>", "-");
            findAndReplaceText(document, "<u_iban", "-");
            findAndReplaceText(document, "<u_bank", "-");

            // Client data
            findAndReplaceText(document, "<c_name>", "-");
            findAndReplaceText(document, "<c_reg_no>", "-");
            findAndReplaceText(document, "<c_f_code", "-");
            findAndReplaceText(document, "<c_address>", "-");

            try (OutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
            }

            // Return file path for the controller
            return outputPath;

        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

//    private void findAndReplace(XWPFDocument document, String placeholder, String replacement) {
//        for (XWPFParagraph paragraph : document.getParagraphs()) {
//            for (XWPFRun run : paragraph.getRuns()) {
//                String text = run.getText(0);
//                if (text != null && text.contains(placeholder)) {
//                    text = text.replace(placeholder, replacement);
//                    run.setText(text, 0);
//                }
//            }
//        }
//    }

    private void findAndReplaceText(XWPFDocument document, String placeholder, String replacement) {
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        for (XWPFRun run : paragraph.getRuns()) {
                            String text = run.getText(0);
                            if (text != null && text.contains(placeholder)) {
                                text = text.replace(placeholder, replacement);
                                run.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
    }
}