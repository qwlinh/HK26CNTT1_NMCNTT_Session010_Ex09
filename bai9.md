# PHẦN 1\. PHÂN TÍCH BÀI TOÁN (I/O)

## 1\. Xác định đầu vào và đầu ra

Đầu vào: Mã tác giả ở C2:C81 và danh mục tác giả. Đầu ra: Tên tác giả, Phòng trình bày, Thù lao tại D2:F81.

Quy ước: Chuyển nguyên sheet DanhMucTacGia từ file riêng vào cùng workbook với DanhSachBaoCao, không chép từng ô. Giả sử dòng 1 là tiêu đề và danh mục nằm trong A2:E1000; thay 1000 bằng dòng cuối phù hợp nếu cần.

| Cột kết quả | Lookup\_value | Table\_array | Col\_index\_num |
| :---- | :---- | :---- | :---- |
| D – Tên tác giả | TRIM(\$C2) | 'DanhMucTacGia'\!\$A\$2:\$E\$1000 | 2 |
| E – Phòng | TRIM(\$C2) | 'DanhMucTacGia'\!\$A\$2:\$E\$1000 | 4 |
| F – Thù lao | TRIM(\$C2) | 'DanhMucTacGia'\!\$A\$2:\$E\$1000 | 5 |

## 2\. Chiến lược xử lý ba bẫy

Khoảng trắng: TRIM(\$C2) loại khoảng trắng thừa ở đầu/cuối mã trước khi tra cứu.

Tác giả chưa có: IFERROR bọc ngoài VLOOKUP để trả “Chưa cập nhật” khi tra cứu lỗi, gồm \#N/A.

Vùng tra bị trôi: Cố định 'DanhMucTacGia'\!\$A\$2:\$E\$1000. Dùng FALSE để tìm mã khớp chính xác; \$C2 giữ cột C nhưng cho số dòng thay đổi.

## 3\. Các bước thực hiện trong Excel

* Đưa hai sheet vào cùng workbook; kiểm tra mã tác giả trong danh mục không trùng và không có khoảng trắng thừa.  
* Xác định vùng tra A2:E1000 và ba cột kết quả D, E, F.  
* Nhập công thức ở D2, E2, F2 theo Phần 2\.  
* Chọn D2:F2 và kéo xuống dòng 81 để xử lý đủ 80 báo cáo.  
* Định dạng F2:F81 theo tiền VNĐ; thử mã có khoảng trắng và rà lại các ô “Chưa cập nhật”.

# PHẦN 2\. TRIỂN KHAI & CHỐNG LỖI

## 4\. Công thức tại D2 – Tên tác giả

\=IFERROR(VLOOKUP(TRIM(\$C2),'DanhMucTacGia'\!\$A\$2:\$E\$1000,2,FALSE),"Chưa cập nhật")

TRIM xử lý mã trước → VLOOKUP tìm trong cột A và trả tên từ cột B → IFERROR thay lỗi bằng thông báo. Dấu \$ giữ nguyên vùng tra khi kéo công thức.

## 5\. Công thức tại E2 và F2

E2 – Phòng trình bày:

\=IFERROR(VLOOKUP(TRIM(\$C2),'DanhMucTacGia'\!\$A\$2:\$E\$1000,4,FALSE),"Chưa cập nhật")

F2 – Thù lao:

\=IFERROR(VLOOKUP(TRIM(\$C2),'DanhMucTacGia'\!\$A\$2:\$E\$1000,5,FALSE),"Chưa cập nhật")

Ba công thức cùng mã tra, vùng tra và cách chống lỗi. Chỉ đổi Col\_index\_num thành 2, 4, 5 vì cần lấy lần lượt cột B, D, E trong cùng bảng; chỉ số được tính từ cột đầu vùng tra.

## 6\. Hiện tượng khi bỏ dấu \$

Nếu D2 dùng 'DanhMucTacGia'\!A2:E1000, khi kéo xuống, vùng tra sẽ dịch theo số dòng:

D2 → A2:E1000; D3 → A3:E1001;

D4 → A4:E1002; D5 → A5:E1003.

Ví dụ TG001 nằm ở A2 của danh mục: công thức tại D3 không còn tìm trong dòng 2 nên có thể báo “Chưa cập nhật” dù mã vẫn tồn tại. Không phải mọi dòng từ D3 đều sai; lỗi xảy ra khi mã cần tìm bị loại khỏi vùng tra. Dùng \$A\$2:\$E\$1000 để tránh tình trạng này.

Lưu ý: Nếu Excel dùng dấu chấm phẩy, thay dấu phẩy ngăn cách đối số bằng “;”. IFERROR bắt mọi lỗi nên cần kiểm tra vùng tra và chỉ số cột nếu nhiều ô cùng báo “Chưa cập nhật”.


