package com.bkademy.excelconverter.service;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.*;
import java.time.Instant;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

@Service
public class UploadSessionService {

    private static final long SESSION_TTL_HOURS = 2;

    private final Path baseDir;
    private final ExcelService excelService;
    private final Map<String, SessionInfo> sessions = new ConcurrentHashMap<>();

    public UploadSessionService(
            ExcelService excelService,
            @Value("${excel-converter.upload.temp-dir:}") String tempDir) throws IOException {
        this.excelService = excelService;
        if (tempDir == null || tempDir.isBlank()) {
            this.baseDir = Files.createTempDirectory("excel-converter-sessions");
        } else {
            this.baseDir = Paths.get(tempDir);
            Files.createDirectories(this.baseDir);
        }
    }

    public String createSession() throws IOException {
        String sessionId = UUID.randomUUID().toString();
        Path sessionDir = baseDir.resolve(sessionId);
        Files.createDirectories(sessionDir.resolve("files"));
        Files.createDirectories(sessionDir.resolve("chunks"));
        sessions.put(sessionId, new SessionInfo(sessionDir));
        return sessionId;
    }

    public void saveChunk(String sessionId, String relativePath, int chunkIndex, int totalChunks,
                          MultipartFile chunk) throws IOException {
        SessionInfo session = requireSession(sessionId);
        Path targetFile = resolveSafeFile(session.getDir(), relativePath);

        if (totalChunks <= 1) {
            Files.createDirectories(targetFile.getParent());
            try (InputStream in = chunk.getInputStream()) {
                Files.copy(in, targetFile, StandardCopyOption.REPLACE_EXISTING);
            }
            session.registerFile(relativePath);
            return;
        }

        Path chunkDir = session.getDir().resolve("chunks").resolve(hashPath(relativePath));
        Files.createDirectories(chunkDir);
        Path chunkFile = chunkDir.resolve("part-" + chunkIndex);
        try (InputStream in = chunk.getInputStream()) {
            Files.copy(in, chunkFile, StandardCopyOption.REPLACE_EXISTING);
        }

        FileChunkState state = session.getChunkStates()
                .computeIfAbsent(relativePath, key -> new FileChunkState(totalChunks));
        state.getReceived().add(chunkIndex);

        if (state.getReceived().size() == totalChunks) {
            mergeChunks(chunkDir, targetFile, totalChunks);
            session.registerFile(relativePath);
            deleteDirectory(chunkDir);
            session.getChunkStates().remove(relativePath);
        }
    }

    public byte[] processSession(String sessionId) throws IOException {
        SessionInfo session = requireSession(sessionId);
        if (session.getUploadedFiles().isEmpty()) {
            throw new IllegalStateException("Chưa có file nào được upload");
        }

        Path filesDir = session.getDir().resolve("files");
        byte[] result = excelService.processDirectory(filesDir);

        Path resultFile = session.getDir().resolve("result.zip");
        Files.write(resultFile, result);
        session.setProcessed(true);
        return result;
    }

    public Path getResultPath(String sessionId) {
        SessionInfo session = requireSession(sessionId);
        if (!session.isProcessed()) {
            throw new IllegalStateException("File kết quả chưa sẵn sàng");
        }
        return session.getDir().resolve("result.zip");
    }

    public void cleanupSession(String sessionId) {
        SessionInfo session = sessions.remove(sessionId);
        if (session != null) {
            deleteDirectoryQuietly(session.getDir());
        }
    }

    @Scheduled(fixedRate = 60 * 60 * 1000)
    public void cleanupExpiredSessions() {
        Instant cutoff = Instant.now().minusSeconds(SESSION_TTL_HOURS * 3600L);
        sessions.entrySet().removeIf(entry -> {
            if (entry.getValue().getCreatedAt().isBefore(cutoff)) {
                deleteDirectoryQuietly(entry.getValue().getDir());
                return true;
            }
            return false;
        });

        try (DirectoryStream<Path> stream = Files.newDirectoryStream(baseDir)) {
            for (Path path : stream) {
                if (!Files.isDirectory(path)) {
                    continue;
                }
                String id = path.getFileName().toString();
                if (!sessions.containsKey(id)) {
                    deleteDirectoryQuietly(path);
                }
            }
        } catch (IOException ignored) {
        }
    }

    private SessionInfo requireSession(String sessionId) {
        SessionInfo session = sessions.get(sessionId);
        if (session == null) {
            throw new IllegalArgumentException("Session không tồn tại hoặc đã hết hạn");
        }
        return session;
    }

    private Path resolveSafeFile(Path sessionDir, String relativePath) {
        Path base = sessionDir.resolve("files").normalize().toAbsolutePath();
        Path resolved = base.resolve(relativePath.replace("\\", "/")).normalize().toAbsolutePath();
        if (!resolved.startsWith(base)) {
            throw new IllegalArgumentException("Đường dẫn file không hợp lệ");
        }
        return resolved;
    }

    private void mergeChunks(Path chunkDir, Path targetFile, int totalChunks) throws IOException {
        Files.createDirectories(targetFile.getParent());
        try (OutputStream out = Files.newOutputStream(targetFile)) {
            for (int i = 0; i < totalChunks; i++) {
                Path part = chunkDir.resolve("part-" + i);
                if (!Files.exists(part)) {
                    throw new IOException("Thiếu chunk " + i);
                }
                Files.copy(part, out);
            }
        }
    }

    private String hashPath(String relativePath) {
        return Integer.toHexString(relativePath.replace("\\", "/").hashCode());
    }

    static void deleteDirectoryQuietly(Path dir) {
        if (dir == null || !Files.exists(dir)) {
            return;
        }
        try {
            deleteDirectory(dir);
        } catch (IOException ignored) {
        }
    }

    private static void deleteDirectory(Path dir) throws IOException {
        try (var walk = Files.walk(dir)) {
            walk.sorted(Comparator.reverseOrder()).forEach(path -> {
                try {
                    Files.deleteIfExists(path);
                } catch (IOException ignored) {
                }
            });
        }
    }

    static class SessionInfo {
        private final Path dir;
        private final Instant createdAt;
        private final Set<String> uploadedFiles = ConcurrentHashMap.newKeySet();
        private final Map<String, FileChunkState> chunkStates = new ConcurrentHashMap<>();
        private boolean processed;

        SessionInfo(Path dir) {
            this.dir = dir;
            this.createdAt = Instant.now();
        }

        Path getDir() {
            return dir;
        }

        Instant getCreatedAt() {
            return createdAt;
        }

        Set<String> getUploadedFiles() {
            return uploadedFiles;
        }

        Map<String, FileChunkState> getChunkStates() {
            return chunkStates;
        }

        boolean isProcessed() {
            return processed;
        }

        void setProcessed(boolean processed) {
            this.processed = processed;
        }

        void registerFile(String relativePath) {
            uploadedFiles.add(relativePath.replace("\\", "/"));
        }
    }

    static class FileChunkState {
        private final int totalChunks;
        private final Set<Integer> received = ConcurrentHashMap.newKeySet();

        FileChunkState(int totalChunks) {
            this.totalChunks = totalChunks;
        }

        int getTotalChunks() {
            return totalChunks;
        }

        Set<Integer> getReceived() {
            return received;
        }
    }
}
