package com.example.exceltoxml.service;

import com.example.exceltoxml.dto.BookDto;
import com.example.exceltoxml.entity.Book;
import com.example.exceltoxml.entity.Bookstore;
import com.example.exceltoxml.repository.BookRepository;
import jakarta.xml.bind.JAXBContext;
import jakarta.xml.bind.JAXBException;
import jakarta.xml.bind.Marshaller;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.modelmapper.ModelMapper;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
@RequiredArgsConstructor
public class BookServiceImpl implements BookService{

    private final ModelMapper mapper;
    private final BookRepository bookRepository;

    @Override
    public Object ExcelToObjAndToXml2(MultipartFile excel) throws IOException, JAXBException {
        List<BookDto> books = readExcelToListObj(excel);
        List<Book> bookList = new ArrayList<>();
//        List<Book> bookList = books.stream().map(i->new Book(i.getId(),i.getTitle(),i.getQuantity(),i.getPrice(),i.getTotalMoney())).collect(Collectors.toList());
        for (BookDto i : books){
            Book x = mapper.map(i,Book.class);
            bookList.add(x);
        }
        bookRepository.saveAll(bookList);

        // list obj convert new file xml
        Bookstore bookstore = new Bookstore();
        bookstore.setName("dinh minh hieu");
        bookstore.setLocation("ha noi");
        bookstore.setBookList(bookList);

        convertObjectToXML(bookstore);
        return null;
    }
    @Override
    public List<BookDto> readExcelToListObj(MultipartFile excel) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(excel.getInputStream());
        XSSFSheet sheet = workbook.getSheetAt(0);
        List<BookDto> bookDtoList = new ArrayList<>();

        for(int i=1; i<sheet.getPhysicalNumberOfRows();i++) {
            XSSFRow row = sheet.getRow(i);
            BookDto bookDto = new BookDto();
            bookDto.setPrice(String.valueOf(row.getCell(row.getPhysicalNumberOfCells()-(row.getPhysicalNumberOfCells()-3))));
            bookDto.setId(String.valueOf(row.getCell(row.getPhysicalNumberOfCells()-(row.getPhysicalNumberOfCells()))));
            bookDto.setQuantity(String.valueOf(row.getCell(row.getPhysicalNumberOfCells()-(row.getPhysicalNumberOfCells()-2))));
            bookDto.setTitle(String.valueOf(row.getCell(row.getPhysicalNumberOfCells()-(row.getPhysicalNumberOfCells()-1))));
            bookDto.setTotalMoney(String.valueOf(row.getCell(row.getPhysicalNumberOfCells()-(row.getPhysicalNumberOfCells()-4))));
            bookDtoList.add(bookDto);
        }
        List<Book> bookList = new ArrayList<>();
        for (BookDto i : bookDtoList){
            Book x = mapper.map(i,Book.class);
            bookList.add(x);
        }

        bookRepository.saveAll(bookList);
        return bookDtoList;
    }

    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_TITLE = 1;
    public static final int COLUMN_INDEX_QUANTITY = 2;
    public static final int COLUMN_INDEX_PRICE = 3;
    public static final int COLUMN_INDEX_TOTAL = 4;

    //    void showExcel()
    public static List<Book> readExcel(String excelFilePath) throws IOException {
        List<Book> listBooks = new ArrayList<>();

        // Get file
        InputStream inputStream = new FileInputStream(new File(excelFilePath));

        //  Get workbook
        Workbook workbook = getWorkbook(inputStream, excelFilePath);

        // Get sheet
        Sheet sheet = workbook.getSheetAt(0);

        // Get all rows
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            if (nextRow.getRowNum() == 0) {
                // Ignore header
                continue;
            }

            // Get all cells
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            // Read cells and set value for book object
            Book book = new Book();
            while (cellIterator.hasNext()) {
                //Read cell
                Cell cell = cellIterator.next();
                Object cellValue = getCellValue(cell);
                if (cellValue == null || cellValue.toString().isEmpty()) {
                    continue;
                }
                // Set value for book object
                int columnIndex = cell.getColumnIndex();
                switch (columnIndex) {
                    case COLUMN_INDEX_ID:
                        book.setId(String.valueOf(new BigDecimal((double) cellValue).intValue()));
                        break;
                    case COLUMN_INDEX_TITLE:
                        book.setTitle((String) getCellValue(cell));
                        break;
                    case COLUMN_INDEX_QUANTITY:
                        book.setQuantity(String.valueOf(new BigDecimal((double) cellValue).intValue()));
                        break;
                    case COLUMN_INDEX_PRICE:
                        book.setPrice(String.valueOf((Double) getCellValue(cell)));
                        break;
                    case COLUMN_INDEX_TOTAL:
                        book.setTotalMoney(String.valueOf((Double) getCellValue(cell)));
                        break;
                    default:
                        break;
                }

            }
            listBooks.add(book);
        }

        workbook.close();
        inputStream.close();

        return listBooks;
    }

    // Get Workbook
    private static Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

    // Get cell value
    private static Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellTypeEnum();
        Object cellValue = null;
        switch (cellType) {
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
                break;
            default:
                break;
        }
        return cellValue;
    }


    private static void convertObjectToXML(Bookstore bookstore) throws JAXBException {
        JAXBContext context = JAXBContext.newInstance(Bookstore.class);
        Marshaller m = context.createMarshaller();
        m.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, Boolean.TRUE);
        m.marshal(bookstore, System.out);
        File newFileXml = new File("ObjToXml10.xml");
        m.marshal(bookstore, newFileXml);

    }
}
