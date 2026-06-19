package com.bkademy.excelconverter.controller;

import com.bkademy.excelconverter.service.UploadSessionService;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

@RestController
@RequestMapping("/api/upload")
public class UploadController {

    private final UploadSessionService uploadSessionService;

    public UploadController(UploadSessionService uploadSessionService) {
        this.uploadSessionService = uploadSessionService;
    }

    @PostMapping("/init")
    public Map<String, String> initSession() throws IOException {
        return Map.of("sessionId", uploadSessionService.createSession());
    }

    @PostMapping("/{sessionId}/chunk")
    public Map<String, String> uploadChunk(
            @PathVariable String sessionId,
            @RequestParam String relativePath,
            @RequestParam int chunkIndex,
            @RequestParam int totalChunks,
            @RequestParam("file") MultipartFile file) throws IOException {
        uploadSessionService.saveChunk(sessionId, relativePath, chunkIndex, totalChunks, file);
        return Map.of("status", "ok");
    }

    @PostMapping("/{sessionId}/complete")
    public Map<String, String> completeUpload(@PathVariable String sessionId) throws IOException {
        uploadSessionService.processSession(sessionId);
        return Map.of("status", "ready", "sessionId", sessionId);
    }

    @GetMapping("/{sessionId}/download")
    public ResponseEntity<InputStreamResource> download(@PathVariable String sessionId) throws IOException {
        Path resultPath = uploadSessionService.getResultPath(sessionId);
        long size = Files.size(resultPath);

        InputStreamResource resource = new InputStreamResource(Files.newInputStream(resultPath));
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=converted_files.zip")
                .header(HttpHeaders.CONTENT_LENGTH, String.valueOf(size))
                .body(resource);
    }

    @DeleteMapping("/{sessionId}")
    public Map<String, String> cancelSession(@PathVariable String sessionId) {
        uploadSessionService.cleanupSession(sessionId);
        return Map.of("status", "deleted");
    }
}
