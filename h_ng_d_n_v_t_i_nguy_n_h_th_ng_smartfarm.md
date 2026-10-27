# BÀI LÀM VÀ HƯỚNG DẪN CHI TIẾT: QUẢN TRỊ HẠ TẦNG SMARTFARM

## CHỨC NĂNG 1: KHỞI TẠO CĂN CỨ SỐ (TERMINAL / POWERSHELL)

### 1.1. Lệnh PowerShell 01 dòng duy nhất để khởi tạo cây thư mục

Mở PowerShell và chạy duy nhất lệnh sau để khởi tạo thư mục gốc `smartfarm-system` cùng 4 thư mục con (`docs`, `scripts`, `data`, `backup`):

```
New-Item -Path "smartfarm-system/docs", "smartfarm-system/scripts", "smartfarm-system/data", "smartfarm-system/backup" -ItemType Directory -Force

```

### 1.2. Tạo file `.env` ẩn bằng đường dẫn tương đối

Di chuyển vào thư mục `data/` và dùng đường dẫn tương đối `..` để tạo file `.env` ở thư mục gốc, sau đó ẩn file:

```
# Di chuyển vào thư mục data
cd smartfarm-system/data

# Tạo file .env ở thư mục gốc (..) với nội dung password
Set-Content -Path "..\.env" -Value "DB_PASSWORD=demo_123456"

# Đặt thuộc tính Hidden cho file .env
(Get-Item "..\.env").Attributes = 'Hidden'

```

## CHỨC NĂNG 2: SCRIPT LỌC LOG TỰ ĐỘNG (`scripts/auto_report.ps1`)

Nội dung bên dưới là toàn bộ mã nguồn của file `auto_report.ps1` (được lưu tại `smartfarm-system/scripts/auto_report.ps1`).

```
# AUDIT: Xac nhan da tham dinh script, khong chua lenh xoa nguy hiem (Remove-Item / del)

# Lay thoi gian hien tai
$currentTime = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

# Write-Output tieu de bao cao
Write-Output "===== BAO CAO SU CO HE THONG SMARTFARM =====" > ..\backup\bao-cao-su-co.txt
Write-Output "Thoi gian chay: $currentTime" >> ..\backup\bao-cao-su-co.txt
Write-Output "----------------------------------------" >> ..\backup\bao-cao-su-co.txt
Write-Output "DANH SACH LOI GRAPHIC & CRITICAL:" >> ..\backup\bao-cao-su-co.txt

# Loc cac dong log chua CRITICAL hoac ERROR bang Pipeline
Get-Content -Path ..\data\log-he-thong.log | Select-String -Pattern "CRITICAL", "ERROR" | ForEach-Object { $_.Line } >> ..\backup\bao-cao-su-co.txt

# Dem so luong dong loi
$errorCount = (Get-Content -Path ..\data\log-he-thong.log | Select-String -Pattern "CRITICAL", "ERROR" | Measure-Object).Count

# Ghi tong so su co vao cuoi file
Write-Output "----------------------------------------" >> ..\backup\bao-cao-su-co.txt
Write-Output "TONG SO SU CO: $errorCount" >> ..\backup\bao-cao-su-co.txt

Write-Host "Xuat bao cao thanh cong tai backup/bao-cao-su-co.txt"

```

### Lệnh thực thi Script:

Nếu bị chặn policy, khởi chạy từ thư mục `scripts/`:

```
powershell -ExecutionPolicy Bypass -File .\auto_report.ps1

```

## CHỨC NĂNG 3: ĐỐI SOÁT TÀI NGUYÊN (EXCEL)

Thực hiện trong file `quan-ly-thiet-bi.xlsx`:

### 1. Chuẩn hóa mã linh kiện (Sheet `trang-thai-tram`)

* Thêm cột `ma_linh_kien_sach` tại ô **`C2`**.

* Nhập công thức chuẩn hóa khoảng trắng thừa:

  ```
  =TRIM(B2)
  
  ```

* Sau đó copy/fill công thức xuống toàn bộ các dòng còn lại của cột.

### 2. Tra cứu tên thiết bị (Sheet `trang-thai-tram`)

* Thêm cột `ten_thiet_bi` tại ô **`D2`**.

* Nhập công thức tra cứu chính xác (`FALSE`) từ sheet `danh-muc-linh-kien`:

  ```
  =VLOOKUP(C2, 'danh-muc-linh-kien'!$A$2:$D$100, 2, FALSE)
  
  ```

  *(Lưu ý: Dùng dấu nháy đơn `'danh-muc-linh-kien'` do tên sheet chứa dấu gạch ngang, cố định vùng dữ liệu bằng ký hiệu `$`).*

### 3. Kiểm tra hạ tầng dữ liệu

* Bật **AutoFilter** cho toàn bộ hàng tiêu đề (`Data` -> `Filter`).

* Đảm bảo không còn bất kỳ ô nào bị lỗi `#N/A`.

* Xóa/tránh các hàng trống hoặc tình trạng Merge Cell trong bảng dữ liệu.

## CHỨC NĂNG 4: TÀI LIỆU HƯỚNG DẪN & PHÂN TÍCH (`docs/README.md`)

Dưới đây là tệp Markdown hoàn chỉnh đặt tại `smartfarm-system/docs/README.md`.

````
# He thong quan tri ha tang SmartFarm

## Huong dan chay script
De chay robot loc log tu dong, mo PowerShell va thuc hien cac lenh sau:

```powershell
cd scripts
powershell -ExecutionPolicy Bypass -File .\auto_report.ps1

````

## Checklist kiem tra thiet bi

* \[x\] Kiem tra ket noi mang tai cac tram sensor

* \[ \] Kiem tra dung luong RAM con trong

* \[ \] Sao luu file log he thong vao thu muc backup/

## Bang thong so phan cung

Dưới đây là thông số các thiết bị mẫu lấy từ danh mục linh kiện:

| Ten thiet bi | Don gia (VND) | Loai Storage | 
| ----- | ----- | ----- | 
| Sensor Nhiet do DHT22 | 150,000 | Non-Storage | 
| Module Gateway LoRa | 1,200,000 | Flash | 
| SSD NVMe 256GB Server | 850,000 | SSD | 

## Phan tich he thong

### 1. Luong xu ly thong tin IPO (Input - Process - Output)

* **Input:** Tệp nhật ký máy chủ thô `data/log-he-thong.log` phát sinh liên tục.

* **Process:** Script `scripts/auto_report.ps1` đọc nội dung, dùng `Select-String` lọc từ khóa `CRITICAL` / `ERROR`, đồng thời đếm tổng số sự cố bằng `Measure-Object`.

* **Output:** Tệp báo cáo tổng hợp `backup/bao-cao-su-co.txt` chứa danh sách chi tiết các dòng lỗi kèm tổng số lượng sự cố.

### 2. Phan loai theo mo hinh DIKW

* **(b) Data:** `"2026-10-12 08:15:01 [ERROR] sensor-03 Connection TIMEOUT after 30s"` *(Dữ liệu thô ban đầu chưa qua xử lý).*

* **(c) Information:** `"Hôm nay hệ thống ghi nhận 8 sự cố, trong đó 4 sự cố đến từ trạm sensor-03."` *(Dữ liệu đã qua tổng hợp và thống kê).*

* **(d) Knowledge:** `"Trạm` sensor-03 hay` mất kết nối vào buổi chiều do sóng yếu, nên cần lắp thêm anten."` *(Tri thức rút ra từ việc phân tích nguyên nhân và quy luật sự cố).*

* **(a) Wisdom:** `"Quyết định thay modem trạm sensor-03 trước thứ Sát."` *(Quyết định quản trị/vận hành đưa ra dựa trên tri thức đã đúc kết).*

## Bay phan cung (Tinh toan thoi gian tai log)

### Yêu cầu:

Tính thời gian tải file log dung lượng **500 MB** qua đường truyền băng thông 100 Mbps