package com.demo.Service.Impl;

import com.demo.Entity.Product;
import com.demo.Service.ProductService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


@Service
public class ProductServiceImpl implements ProductService {

    // Excel sheets headers
    public static String[] HEADERS = {"id", "category", "name", "quantity", "price", "total cost"};

    private static List<Product> productList = new ArrayList<>();
    private static int lastId = 100; // starting from the highest ID already used

    static {
        // it will be store the product value during the class loading.
    }

    @Override
    public List<Product> getAllProduct() {
        return productList;
    }

    @Override
    public Product saveProduct(Product product) {
        product.setId(++lastId);
        productList.add(product);
        return product;
    }

    @Override
    public ByteArrayInputStream generateExcel() {

        ByteArrayOutputStream out= null;

        try {
            // create excel page worksheet
            Workbook workbook = new XSSFWorkbook();
            out = new ByteArrayOutputStream();

            // write the sheet name
            Sheet sheet = workbook.createSheet("Product");

            // set header font color and style
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.RED.getIndex());

            // set cell style
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            // create row for header
            Row row = sheet.createRow(0);
            for (int i = 0; i < HEADERS.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(HEADERS[i]);
                cell.setCellStyle(headerCellStyle);
            }

            int rowIndex = 1;

            for (Product prod : productList) {
                double totalCost = prod.getQuantity() * prod.getPrice();

                //create row and cell to insert data
                Row dataRow = sheet.createRow(rowIndex);

                // set data to each cell
                dataRow.createCell(0).setCellValue(prod.getId());
                dataRow.createCell(1).setCellValue(prod.getCategory());
                dataRow.createCell(2).setCellValue(prod.getName());
                dataRow.createCell(3).setCellValue(prod.getQuantity());
                dataRow.createCell(4).setCellValue(prod.getPrice());
                dataRow.createCell(5).setCellValue(totalCost);

                // increment the row index
                rowIndex++;
            }

            // write the data into Excel file
            workbook.write(out);

        } catch (IOException e) {
            e.printStackTrace();
        }

        // return data into the ByteArray
        return new ByteArrayInputStream(out.toByteArray());
    }
}