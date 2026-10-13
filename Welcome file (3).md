
# THIẾT KẾ HỆ THỐNG - GREENTECH INFRASTRUCTURE

## PHẦN 1: TƯ DUY HỆ THỐNG & MÔ HÌNH XỬ LÝ (DIKW & IPO)

### 1. Mô hình IPO (Input - Process - Output)

**Input (Đầu vào):**

-   File nhật ký thô kích thước lớn: `log-he-thong.log`.
    
-   Bảng dữ liệu phần cứng: `quan-ly-thiet-bi.xlsx`.
    
-   Tham số cấu hình hệ thống từ tệp ẩn: `.env`.
    

**Process (Xử lý):**

-   Sử dụng lệnh Terminal lồng nhau để tạo cấu trúc cây thư mục.
    
-   Dùng đường ống (Pipe) kết hợp `findstr`/`grep` để lọc lỗi.
    
-   Áp dụng hàm `TRIM` và `VLOOKUP` để xử lý dữ liệu bảng tính.
    
-   Định dạng tài liệu theo chuẩn Heading Styles và Markdown.
    

**Output (Đầu ra):**

-   Tệp báo cáo sự cố tự động: `bao-cao-su-co.txt`.
    
-   Tài liệu vận hành: `README.md`, `srs-he-thong.docx`.
    
-   Cơ sở dữ liệu đối soát phần cứng hoàn chỉnh, không còn lỗi `#N/A`.
    

### 2. Mô hình DIKW (Data - Information - Knowledge - Wisdom)

Cấp độ

Nội dung

**Data (Dữ liệu thô)**

Hàng triệu dòng log ký tự vô nghĩa và bảng mã linh kiện rời rạc.

**Information (Thông tin)**

Các dòng lỗi cụ thể (`ERROR`, `CRITICAL`, `404`, `TIMEOUT`) được lọc ra và gom vào báo cáo.

**Knowledge (Tri thức)**

Hiểu được xu hướng lỗi hệ thống, trạng thái hoạt động của 100 trạm IoT ngoài trang trại dựa trên báo cáo và bảng đối soát thiết bị.

**Wisdom (Sự khôn ngoan/Hành động)**

Đưa ra quyết định bảo trì phần cứng kịp thời, tối ưu hóa hạ tầng mạng và thiết lập quy trình vận hành tự động hóa bền vững cho Startup.

----------

## MODULE 4: ĐỐI SOÁT TÀI NGUYÊN (EXCEL DATABASE MINDSET)

Giải quyết triệt để **Bẫy Excel** (không chứa hàng trống, không dùng Merge Cell) và **Bẫy dữ liệu** (khoảng trắng thừa gây lỗi `#N/A`).

### 1. Xử lý khoảng trắng thừa với hàm TRIM

Trong cơ sở dữ liệu thực tế, các mã linh kiện thường bị lẫn khoảng trắng vô hình ở đầu hoặc cuối chuỗi, ví dụ `"MK01 "` so với `"MK01"`, khiến hàm tìm kiếm trả về lỗi `#N/A`.

**Giải pháp:** Tạo một cột phụ trợ hoặc chuẩn hóa trực tiếp dữ liệu nguồn bằng công thức:

```
=TRIM(A2)
```

### 2. Tự động điền tên thiết bị bằng hàm VLOOKUP chuẩn xác

Tại sheet `trang-thai-tram`, để điền cột `ten_thiet_bi` tự động từ sheet `danh-muc-linh-kien`, áp dụng công thức sau tại ô tương ứng, ví dụ cột B:

```
=VLOOKUP(TRIM(B2),'danh-muc-linh-kien'!$A$2:$D$100,2,FALSE)
```

**Giải thích các tham số chống lỗi:**

-   `**TRIM(B2)**`**:** Làm sạch mã linh kiện tại trạm trước khi dò tìm, khắc phục lỗi khoảng trắng thừa.
    
-   `**'danh-muc-linh-kien'!$A$2:$D$100**`**:** Vùng dữ liệu bảng chuẩn, cố định tuyệt đối bằng dấu `$` để khi kéo công thức xuống không bị lệch vùng chọn.
    
-   `**2**`**:** Lấy giá trị trả về ở cột thứ 2 (tên thiết bị) trong bảng danh mục.
    
-   `**FALSE**`**:** Bắt buộc tìm kiếm chính xác tuyệt đối (_Exact Match_), tránh việc trả về giá trị gần đúng sai lệch.
    

**Lưu ý:** Cột chứa mã linh kiện cần dò tìm phải là cột đầu tiên trong vùng dữ liệu VLOOKUP. Nếu dữ liệu nguồn cũng có khoảng trắng thừa, hãy chuẩn hóa cột mã trong danh mục bằng `TRIM` trước khi đối soát.
