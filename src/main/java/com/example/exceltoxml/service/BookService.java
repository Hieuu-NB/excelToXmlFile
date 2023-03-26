package com.example.exceltoxml.service;

import com.example.exceltoxml.dto.BookDto;
import jakarta.xml.bind.JAXBException;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface BookService {
    Object ExcelToObjAndToXml2(MultipartFile excel) throws IOException, JAXBException;
    List<BookDto> readExcelToListObj(MultipartFile excel) throws IOException;
}
