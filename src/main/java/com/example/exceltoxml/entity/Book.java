package com.example.exceltoxml.entity;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.*;

@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
@Entity
@Builder
@Table(name = "book")
public class Book {
    @Id
    private String id;
    private String title;
    private String quantity;
    private String price;
    private String totalMoney;
}
