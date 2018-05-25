# XLS2PDF
This is the library for convert excel `XLS` and `XLSX` files to `HTML` or `PDF` formats

Based on:
- jxls -  for convert XLS,XLSX to HTML
- itext - for convert HTML to PDF

## Usage:
### for HTML (Xls2HtmlConverter)
```java
new Xls2HtmlConverter(
        new FileInputStream("source.xlsx"), // source excel file
        new FileOutputStream("target.html") // result HTML file
).convert();
```

### for PDF (Xls2PdfConverter)
```java
new Xls2PdfConverter(
        new FileInputStream("source.xlsx"), // source excel file
        new FileOutputStream("target.html") // result HTML file
).convert();
```

###Attention!
This is beta. Do not use this on production!