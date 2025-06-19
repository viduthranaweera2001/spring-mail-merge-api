package com.example.mail_merge_api.controller;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.core.io.*;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;
import java.util.zip.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@RestController
public class MailMergeController {

    private static final Logger logger = LoggerFactory.getLogger(MailMergeController.class);

    // In-memory storage for merged documents and images
    private final Map<String, List<byte[]>> mergedDocumentsStore = new HashMap<>();
    private final Map<String, Map<String, byte[]>> imageStore = new HashMap<>();

    @PostMapping("/test")
    public String test() {
        return "test";
    }

    @PostMapping("/mail-merge")
    public ResponseEntity<Map<String, String>> performMailMerge(
            @RequestParam("wordTemplate") MultipartFile wordFile,
            @RequestParam("excelData") MultipartFile excelFile) throws IOException, InvalidFormatException {

        String sessionId = UUID.randomUUID().toString();
        logger.info("Starting mail merge for session: {}", sessionId);

        List<Map<String, String>> personDetails = readExcelFile(excelFile);
        logger.info("Read {} data rows from Excel", personDetails.size());

        List<byte[]> mergedDocuments = new ArrayList<>();
        for (Map<String, String> person : personDetails) {
            byte[] mergedDoc = processWordTemplate(wordFile, person);
            mergedDocuments.add(mergedDoc);
        }
        logger.info("Generated {} merged documents", mergedDocuments.size());

        mergedDocumentsStore.put(sessionId, mergedDocuments);

        Map<String, String> response = new HashMap<>();
        response.put("sessionId", sessionId);
        return ResponseEntity.ok().body(response);
    }

    @GetMapping("/mail-merge/preview/{sessionId}")
    public ResponseEntity<List<String>> getPreview(@PathVariable String sessionId) throws IOException {
        logger.info("Generating preview for session: {}", sessionId);
        List<byte[]> mergedDocuments = mergedDocumentsStore.get(sessionId);
        if (mergedDocuments == null || mergedDocuments.isEmpty()) {
            logger.warn("No documents found for session: {}", sessionId);
            return ResponseEntity.status(404).body(List.of("No documents found for session: " + sessionId));
        }

        List<String> previews = new ArrayList<>();
        Map<String, byte[]> sessionImages = new HashMap<>();
        imageStore.put(sessionId, sessionImages);

        for (int docIndex = 0; docIndex < mergedDocuments.size(); docIndex++) {
            try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(mergedDocuments.get(docIndex)))) {
                StringBuilder html = new StringBuilder();
                html.append("<div class='document-preview'>");
                boolean hasImages = false;

                // Process paragraphs for text and inline images
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    StringBuilder paraHtml = new StringBuilder("<p");
                    if (paragraph.getAlignment() != null) {
                        switch (paragraph.getAlignment()) {
                            case CENTER:
                                paraHtml.append(" style='text-align: center;'");
                                break;
                            case RIGHT:
                                paraHtml.append(" style='text-align: right;'");
                                break;
                            default:
                                paraHtml.append(" style='text-align: left;'");
                        }
                    }
                    paraHtml.append(">");

                    for (XWPFRun run : paragraph.getRuns()) {
                        String text = run.getText(0);
                        if (text != null) {
                            text = text.replace("&", "&").replace("<", "<").replace(">", ">");
                            StringBuilder runHtml = new StringBuilder();
                            if (run.isBold()) runHtml.append("<strong>");
                            if (run.isItalic()) runHtml.append("<em>");
                            runHtml.append(text);
                            if (run.isItalic()) runHtml.append("</em>");
                            if (run.isBold()) runHtml.append("</strong>");
                            paraHtml.append(runHtml);
                        }
                        // Check for inline images
                        for (XWPFPicture picture : run.getEmbeddedPictures()) {
                            XWPFPictureData pictureData = picture.getPictureData();
                            if (pictureData != null && pictureData.getData() != null && pictureData.getData().length > 0) {
                                String imageId = UUID.randomUUID().toString();
                                sessionImages.put(imageId, pictureData.getData());
//                                String imageUrl = String.format("/mail-merge/image/%s/%d/%s", sessionId, docIndex, imageId);
//                                paraHtml.append(String.format("<img src='%s' alt='Inline Image' style='max-width: 100%%; height: auto;'/>", imageUrl));
                                logger.info("Added inline image for document {}, imageId: {}, size: {} bytes", docIndex + 1, imageId, pictureData.getData().length);
                                hasImages = true;
                            } else {
                                logger.warn("Invalid or empty picture data in run for document {}", docIndex + 1);
                            }
                        }
                    }
                    paraHtml.append("</p>");
                    html.append(paraHtml);
                }

                // Process document-level images
                List<XWPFPictureData> pictures = document.getAllPictures();
                logger.info("Found {} document-level images in document {}", pictures.size(), docIndex + 1);
                for (XWPFPictureData picture : pictures) {
                    if (picture.getData() != null && picture.getData().length > 0) {
                        String imageId = UUID.randomUUID().toString();
                        sessionImages.put(imageId, picture.getData());
                        String imageUrl = String.format("/mail-merge/image/%s/%d/%s", sessionId, docIndex, imageId);
                        html.append(String.format("<p><img src='%s' alt='Document Image' style='max-width: 100%%; height: auto;'/></p>", imageUrl));
                        logger.info("Added document-level image for document {}, imageId: {}, size: {} bytes", docIndex + 1, imageId, picture.getData().length);
                        hasImages = true;
                    } else {
                        logger.warn("Invalid or empty picture data at document level for document {}", docIndex + 1);
                    }
                }

                if (!hasImages) {
                    logger.warn("No valid images found in document {} for session {}", docIndex + 1, sessionId);
                    html.append("<p class='image-error'>No images found in this document. Ensure images are embedded in the Word document.</p>");
                }

                html.append("</div>");
                previews.add(html.toString());
            } catch (Exception e) {
                logger.error("Error processing document {} for session {}: {}", docIndex + 1, sessionId, e.getMessage(), e);
                previews.add("<div class='document-preview'><p class='error'>Error generating preview: " + e.getMessage() + "</p></div>");
            }
        }

        logger.info("Generated {} previews for session {}", previews.size(), sessionId);
        return ResponseEntity.ok(previews);
    }

    @GetMapping("/mail-merge/download/{sessionId}/{index}")
    public ResponseEntity<Resource> downloadSingleDocument(@PathVariable String sessionId, @PathVariable int index) {
        logger.info("Downloading document for session: {}, index: {}", sessionId, index);
        List<byte[]> mergedDocuments = mergedDocumentsStore.get(sessionId);
        if (mergedDocuments == null || index < 0 || index >= mergedDocuments.size()) {
            logger.error("No document found for session: {}, index: {}", sessionId, index);
            return ResponseEntity.status(404)
                    .body(new ByteArrayResource(("No document found for session: " + sessionId + ", index: " + index).getBytes()));
        }

        ByteArrayResource resource = new ByteArrayResource(mergedDocuments.get(index));
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Merged_Letter_" + (index + 1) + ".docx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resource);
    }

    @GetMapping("/mail-merge/download-zip/{sessionId}")
    public ResponseEntity<Resource> downloadZip(@PathVariable String sessionId) throws IOException {
        logger.info("Downloading ZIP for session: {}", sessionId);
        List<byte[]> mergedDocuments = mergedDocumentsStore.get(sessionId);
        if (mergedDocuments == null || mergedDocuments.isEmpty()) {
            logger.error("No documents found for session: {}", sessionId);
            return ResponseEntity.status(404)
                    .body(new ByteArrayResource(("No documents found for session: " + sessionId).getBytes()));
        }

        ByteArrayOutputStream zipOutputStream = new ByteArrayOutputStream();
        try (ZipOutputStream zip = new ZipOutputStream(zipOutputStream)) {
            for (int i = 0; i < mergedDocuments.size(); i++) {
                ZipEntry entry = new ZipEntry("Merged_Letter_" + (i + 1) + ".docx");
                zip.putNextEntry(entry);
                zip.write(mergedDocuments.get(i));
                zip.closeEntry();
            }
        }

        // Clean up after download
        mergedDocumentsStore.remove(sessionId);
        imageStore.remove(sessionId);
        logger.info("Cleaned up storage for session: {}", sessionId);

        ByteArrayResource resource = new ByteArrayResource(zipOutputStream.toByteArray());
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=merged_documents.zip")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resource);
    }

    private List<Map<String, String>> readExcelFile(MultipartFile excelFile) throws IOException {
        List<Map<String, String>> personDetails = new ArrayList<>();
        try (XSSFWorkbook workbook = new XSSFWorkbook(excelFile.getInputStream())) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow headerRow = sheet.getRow(0);
            int rowCount = sheet.getPhysicalNumberOfRows();

            List<String> headers = new ArrayList<>();
            for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
                headers.add(headerRow.getCell(i).getStringCellValue());
            }

            for (int i = 1; i < rowCount; i++) {
                XSSFRow row = sheet.getRow(i);
                Map<String, String> person = new HashMap<>();
                for (int j = 0; j < headers.size(); j++) {
                    String value = row.getCell(j) != null ? row.getCell(j).toString() : "";
                    person.put(headers.get(j), value);
                }
                personDetails.add(person);
            }
        }
        return personDetails;
    }

    private byte[] processWordTemplate(MultipartFile wordFile, Map<String, String> person)
            throws IOException, InvalidFormatException {
        try (XWPFDocument document = new XWPFDocument(wordFile.getInputStream())) {
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText();
                if (text.contains("${")) {
                    for (XWPFRun run : paragraph.getRuns()) {
                        String runText = run.getText(0);
                        if (runText != null) {
                            for (Map.Entry<String, String> entry : person.entrySet()) {
                                String placeholder = "${" + entry.getKey() + "}";
                                runText = runText.replace(placeholder, entry.getValue());
                            }
                            run.setText(runText, 0);
                        }
                    }
                }
            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            document.write(outputStream);
            return outputStream.toByteArray();
        }
    }
}