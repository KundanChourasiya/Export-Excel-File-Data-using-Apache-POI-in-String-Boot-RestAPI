package com.demo.Service;

import com.demo.Entity.Product;

import java.io.ByteArrayInputStream;
import java.util.List;

public interface ProductService {

    // GET: Retrieve all Product
    public List<Product> getAllProduct();

    // POST: Save a new Product
    public Product saveProduct(Product product);

    // GET: Generate a product list Report
    public ByteArrayInputStream generateExcel();
}
