package com.demo.Entity;

import lombok.*;

@Setter
@Getter
@AllArgsConstructor
@NoArgsConstructor
@Builder
public class Product {

    private Integer id;
    private String category;
    private String name;
    private Integer quantity;
    private double price;

}