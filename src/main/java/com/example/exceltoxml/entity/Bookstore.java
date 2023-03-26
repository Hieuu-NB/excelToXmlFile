package com.example.exceltoxml.entity;

import jakarta.xml.bind.annotation.XmlElement;
import jakarta.xml.bind.annotation.XmlElementWrapper;
import jakarta.xml.bind.annotation.XmlRootElement;

import java.util.List;

@XmlRootElement(namespace = "net.javaguides.javaxmlparser.jaxb")
public class Bookstore {

    @XmlElementWrapper(name = "bookList")
    @XmlElement(name = "book")
    private List< Book > bookList;
    private String name;
    private String location;

    public void setBookList(List < Book > bookList) {
        this.bookList = bookList;
    }

    public List < Book > getBooksList() {
        return bookList;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getLocation() {
        return location;
    }

    public void setLocation(String location) {
        this.location = location;
    }
}