package com.bkademy.excelconverter.controller;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.multipart.MaxUploadSizeExceededException;

@RestControllerAdvice
public class UploadExceptionHandler {

    private static final Logger log = LoggerFactory.getLogger(UploadExceptionHandler.class);
    private static final String GENERIC_MESSAGE = "Hệ thống gặp lỗi bất ngờ. Vui lòng thử lại sau.";

    @ExceptionHandler(MaxUploadSizeExceededException.class)
    public ResponseEntity<String> handleMaxUploadSize(MaxUploadSizeExceededException ex) {
        log.warn("Upload vượt quá giới hạn dung lượng", ex);
        return ResponseEntity.status(HttpStatus.PAYLOAD_TOO_LARGE).body(GENERIC_MESSAGE);
    }

    @ExceptionHandler(Exception.class)
    public ResponseEntity<String> handleUnexpected(Exception ex) {
        log.error("Lỗi xử lý request", ex);
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(GENERIC_MESSAGE);
    }
}
