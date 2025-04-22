# Export-Excel-File-Data-using-Apache-POI-in-String-Boot-RestAPI

> [!NOTE]
> ### In this Api we generate Excel file using Apache POI in spring boot RestApi.

## Tech Stack
- Java-17
- Spring Boot-3x
- lombok
- Apache-POI
- PostMan

## Modules
* Product Module

## API Root Endpoint
```
https://localhost:8080/
user this data for checking purpose.
```
## Step To Be Followed
> 1. Create Rest Api will return to Product Details.
>    
>    **Project Documentation**
>    - **Entity** - Product (class)
>    - **Payload** - ApiResponceDto (class)
>    - **Service** - ProductService (interface), ProductServiceImpl (class)
>    - **Controller** - ProductController (Class)
>      
> 2. Add Apache POI Java library in pom.xml file.
> 3. Create a generate excel method in service class.

## Important Dependency to be used
```xml 
 <dependency>
     <groupId>org.springframework.boot</groupId>
     <artifactId>spring-boot-starter-web</artifactId>
 </dependency>

 <dependency>
     <groupId>org.projectlombok</groupId>
     <artifactId>lombok</artifactId>
     <optional>true</optional>
 </dependency>

<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi-ooxml</artifactId>
	<version>5.4.0</version>
</dependency>

```

## Create Product class in Entity Package.
```java
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
```

## Create ProductService interface and ProductServiceImpl class in Service package.

### *ProductService*
```java
public interface ProductService {

    // GET: Retrieve all Product
    public List<Product> getAllProduct();

    // POST: Save a new Product
    public Product saveProduct(Product product);

    // GET: Generate a product list excel Report
    public ByteArrayInputStream generateExcel();
}
```

### *ProductServiceImpl*
```java
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
            return new ByteArrayInputStream(out.toByteArray());

        } catch (IOException e) {
            e.printStackTrace();
        }

        // return data into the ByteArray
        return new ByteArrayInputStream(out.toByteArray());
    }
}
```

##  Create ApiResponse inside the Payload Package.
### *ApiResponseDto* 
```java
@Setter
@Getter
@NoArgsConstructor
public class ApiResponse<T> {
    private boolean status;
    private String message;
    private T data;
    public ApiResponse(boolean status, String message, T data) {
        this.status = status;
        this.message = message;
        this.data = data;
    }
}
```

### *Create ProductController class inside the Controller Package.* 

```java
@RestController
@RequestMapping("/api/v1")
public class ProductController {

    private ProductService productService;

    public ProductController(ProductService productService) {
        this.productService = productService;
    }

    // GET: Retrieve all Product
    // URL: http://localhost:8080/api/v1/product
    @GetMapping("/product")
    public ResponseEntity<ApiResponse<?>> getAllProducts() {
        List<Product> allProduct = productService.getAllProduct();
        if (allProduct.isEmpty()) {
            ApiResponse<Object> response = new ApiResponse<>(false, "Product List Empty!!!", null);
            return ResponseEntity.status(HttpStatus.NOT_FOUND).body(response);
        }
        ApiResponse<List<Product>> response = new ApiResponse<>(true, "Product List Found!!!", allProduct);
        return ResponseEntity.status(HttpStatus.OK).body(response);
    }

    // POST: Save a new Product
    // URL: http://localhost:8080/api/v1/product
    @PostMapping(value = "/product")
    public ResponseEntity<ApiResponse<?>> saveProduct(@RequestBody Product product) {
        try {
            Product saveProduct = productService.saveProduct(product);
            ApiResponse<Product> response = new ApiResponse<>(true, "Product saved successfully!!!", saveProduct);
            return ResponseEntity.status(HttpStatus.CREATED).body(response);

        } catch (Exception e) {
            e.printStackTrace();
            ApiResponse<Product> response = new ApiResponse<>(false, "Product Not Saved!!!", null);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(response);
        }
    }

    // GET: Generate a product list Report
    // URL: http://localhost:8080/api/v1/product/excel-report
    @GetMapping(value = "/product/excel-report", produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public ResponseEntity<InputStreamResource> productExcelReport() throws IOException {
        ByteArrayInputStream generateExcel = productService.generateExcel();
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Product_List.xlsx");
        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(new InputStreamResource(generateExcel));
    }

}

```

### Following pictures will help to understand flow of API

### *Postman Test Cases*

Url - http://localhost:8080/api/v1/product
![image](https://github.com/user-attachments/assets/6ae5a8e5-69c0-4e01-bfef-d2a24afe8441)

Url - http://localhost:8080/api/v1/product
![image](https://github.com/user-attachments/assets/a07ee790-5d42-4d9c-8449-8825d03ff0f9)

Url - http://localhost:8080/api/v1/product/excel-report
**Note**: copy and paste the browser search bar. Because File not showing in Postman. Postman doesn't render binary Excel files as previewable content	Save the response manually and open it in Excel.

![image](https://github.com/user-attachments/assets/f0172935-9903-4d8c-a99d-f62e642e16dc)

