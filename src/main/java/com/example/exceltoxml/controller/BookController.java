package com.example.exceltoxml.controller;

import com.example.exceltoxml.dto.BookDto;
import com.example.exceltoxml.service.BookService;
import jakarta.xml.bind.JAXBException;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@RestController
@RequiredArgsConstructor
@RequestMapping("/book")
public class BookController {
    private final BookService bookServicel;

    @PostMapping("/taoXml")
    public ResponseEntity<Object> createXml1(@RequestParam MultipartFile excel) throws JAXBException, IOException {
        return ResponseEntity.ok(bookServicel.ExcelToObjAndToXml2(excel));
    }
    @GetMapping("/ ")
    public ResponseEntity<List<BookDto>> readExcel(@RequestParam MultipartFile excel) throws IOException {
        return ResponseEntity.ok(bookServicel.readExcelToListObj(excel));
    }
}
