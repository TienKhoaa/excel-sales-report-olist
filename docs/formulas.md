# Power Query và công thức chính (project báo cáo kinh doanh trên Excel)

Tài liệu tổng hợp các bước Power Query và công thức dùng trong `Olist_Sales_Report.xlsx`. Excel dùng dấu phẩy giữa các tham số. Nếu máy dùng cài đặt khu vực tiếng Việt, đổi thành dấu chấm phẩy.

## 1. Power Query

6 file CSV nạp bằng **Data** → **Get Data** → **From Text/CSV** (UTF-8, dấu phẩy), đổi tên query và đặt ở chế độ **Only Create Connection**: `Orders`, `OrderItems`, `Customers`, `Products`, `CategoryTranslation`, `Reviews`.

### Orders

Giữ các cột: `order_id`, `customer_id`, `order_status`, `order_purchase_timestamp`, `order_delivered_customer_date`, `order_estimated_delivery_date`. Lọc bỏ `order_status` là `canceled` và `unavailable`. Thêm các cột:

```
OrderDate = Date.From([order_purchase_timestamp])
```
```
YearMonth = Date.ToText([OrderDate], "yyyy-MM")
```
```
DeliveryDays =
if [order_delivered_customer_date] = null then null
else Duration.Days(Date.From([order_delivered_customer_date]) - [OrderDate])
```
```
OnTime =
if [order_delivered_customer_date] = null then null
else if Date.From([order_delivered_customer_date]) <= Date.From([order_estimated_delivery_date]) then 1 else 0
```

### OrderItems

Giữ các cột: `order_id`, `order_item_id`, `product_id`, `price`, `freight_value`.

### ReviewsByOrder (tham chiếu từ Reviews)

Giữ `order_id` và `review_score`, rồi **Group By** `order_id` với phép **Average** của `review_score`, đặt tên cột `AvgReviewScore`. Mỗi đơn còn đúng một dòng.

### CategoryTranslation

Dùng dòng đầu làm tên cột (**Use First Row as Headers**).

### Products

Ghép với `CategoryTranslation` theo `product_category_name` (**Left Outer**), mở rộng cột `product_category_name_english`, rồi thêm cột:

```
Category =
if [product_category_name_english] <> null then [product_category_name_english]
else if [product_category_name] <> null then [product_category_name]
else "unknown"
```
Chỉ giữ `product_id` và `Category`.

### StateRegion (từ bảng tra trong sheet Lookup)

27 bang Brazil ánh xạ sang 5 khu vực: Centro-Oeste, Nordeste, Norte, Sudeste, Sul.

### Customers

Ghép với `StateRegion` theo `customer_state` và `State` (**Left Outer**), lấy cột `Region`. Giữ `customer_id`, `customer_unique_id`, `customer_state`, `Region`.

### Sales_Data (tham chiếu từ OrderItems)

1. Ghép `Orders` theo `order_id` bằng **Inner**, mở rộng: `customer_id`, `order_status`, `OrderDate`, `YearMonth`, `DeliveryDays`, `OnTime`.
2. Ghép `Customers` theo `customer_id` bằng **Left Outer**, mở rộng: `customer_unique_id`, `customer_state`, `Region`.
3. Ghép `Products` theo `product_id` bằng **Left Outer**, mở rộng `Category`.
4. Ghép `ReviewsByOrder` theo `order_id` bằng **Left Outer**, mở rộng `AvgReviewScore`.
5. Kiểm tra số dòng sau từng lần ghép: không tăng (tránh nhân đôi doanh thu).
6. Thêm các cột:

```
OrderFlag = if [order_item_id] = 1 then 1 else 0
```
```
MonthStart = Date.StartOfMonth([OrderDate])
```
```
DeliveredOrder = if [OrderFlag] = 1 and [OnTime] <> null then 1 else 0
```
```
OnTimeOrder = if [OrderFlag] = 1 and [OnTime] = 1 then 1 else 0
```
```
CategoryClean =
if [Category] = null or Text.Trim([Category]) = "" then "unknown" else [Category]
```
Xóa cột `Category` cũ và đổi `CategoryClean` thành `Category`. Đổi tên `price` thành `Revenue`, `freight_value` thành `Freight`. Kiểu dữ liệu: `Revenue` và `Freight` là **Currency** (Fixed decimal), `OrderDate` và `MonthStart` là **Date**.

Nạp thành Excel Table tên **tblSales** (112.101 dòng).

## 2. Tên đã đặt (Named Range)

| Tên | Ô | Ý nghĩa |
|---|---|---|
| StartMonth | Assumptions!B3 | Tháng bắt đầu báo cáo (01/2017) |
| EndMonth | Assumptions!B4 | Tháng kết thúc báo cáo (08/2018) |
| TargetGrowth | Assumptions!B5 | Tăng trưởng mục tiêu (5%) |
| MinOrdersYoY | Assumptions!B6 | Số đơn tối thiểu để tính YoY (500) |
| TargetOK | Assumptions!B7 | Ngưỡng đạt (100%) |
| TargetWarn | Assumptions!B8 | Ngưỡng cảnh báo (90%) |
| SelMonth | Dashboard!C3 | Tháng đang chọn |
| SelRegion | Dashboard!C4 | Khu vực đang chọn |
| RegionCrit | Calc_Dash!B3 | Điều kiện khu vực dùng cho SUMIFS |
| MonthList, RegionList | Lookup / Calc | Danh sách cho ô chọn |

## 3. Sheet Calc: báo cáo theo tháng (toàn công ty)

Danh sách tháng (tự tràn):
```
=SORT(UNIQUE(FILTER(tblSales[MonthStart],(tblSales[MonthStart]>=StartMonth)*(tblSales[MonthStart]<=EndMonth))))
```
Các cột (dòng 5, kéo xuống):
```
Doanh thu = SUMIFS(tblSales[Revenue],tblSales[MonthStart],A5)
Số đơn = SUMIFS(tblSales[OrderFlag],tblSales[MonthStart],A5)
Giá trị đơn TB = IFERROR(B5/C5,0)
MoM % = B6/B5-1
Doanh thu cùng kỳ năm trước = SUMIFS(tblSales[Revenue],tblSales[MonthStart],EDATE(A5,-12))
Số đơn cùng kỳ năm trước = SUMIFS(tblSales[OrderFlag],tblSales[MonthStart],EDATE(A5,-12))
YoY % = IF(G5<MinOrdersYoY,"",B5/F5-1)
```

## 4. Sheet Calc_Region: theo khu vực, mục tiêu, % đạt

Tên khu vực:
```
=TRANSPOSE(SORT(UNIQUE(tblSales[Region])))
```
Doanh thu thực tế (một công thức duy nhất cho cả khối, tự tràn):
```
=SUMIFS(tblSales[Revenue],tblSales[MonthStart],A5#,tblSales[Region],B4#)
```
Mục tiêu = trung bình 3 tháng liền trước x (1 + tăng trưởng), ba tháng đầu để trống:
```
=IF(ROWS(I$5:I5)<=3,"",AVERAGE(B2:B4)*(1+TargetGrowth))
```
% đạt mục tiêu:
```
=IF(I5="","",B5/I5)
```
Tô màu % đạt bằng Conditional Formatting (ba luật dùng công thức):
```
=AND(ISNUMBER(P5),P5>=TargetOK)
=AND(ISNUMBER(P5),P5>=TargetWarn,P5<TargetOK)
=AND(ISNUMBER(P5),P5<TargetWarn)
```

## 5. Sheet Calc_Dash: chỉ số của tháng và khu vực đang chọn

Điều kiện khu vực (dấu `*` nghĩa là mọi khu vực):
```
RegionCrit = IF(SelRegion="All","*",SelRegion)
```
Các chỉ số chính:
```
Doanh thu = SUMIFS(tblSales[Revenue],tblSales[MonthStart],SelMonth,tblSales[Region],RegionCrit)
Số đơn = SUMIFS(tblSales[OrderFlag],tblSales[MonthStart],SelMonth,tblSales[Region],RegionCrit)
Đơn đã giao = SUMIFS(tblSales[OrderFlag],tblSales[MonthStart],SelMonth,tblSales[Region],RegionCrit,tblSales[OnTime],">=0")
Đơn giao đúng hạn = SUMIFS(tblSales[OrderFlag],tblSales[MonthStart],SelMonth,tblSales[Region],RegionCrit,tblSales[OnTime],1)
Số ngày giao TB = IFERROR(AVERAGEIFS(tblSales[DeliveryDays],tblSales[MonthStart],SelMonth,tblSales[Region],RegionCrit,tblSales[OrderFlag],1),"n/a")
Điểm đánh giá TB = IFERROR(AVERAGEIFS(tblSales[AvgReviewScore],tblSales[MonthStart],SelMonth,tblSales[Region],RegionCrit,tblSales[OrderFlag],1),"n/a")
Điểm: đơn đúng hạn = ...AVERAGEIFS(...,tblSales[OrderFlag],1,tblSales[OnTime],1)
Điểm: đơn giao trễ = ...AVERAGEIFS(...,tblSales[OrderFlag],1,tblSales[OnTime],0)
```
Mục tiêu của tháng đang chọn (tổng nếu chọn All, theo cột khu vực nếu chọn một khu vực):
```
=IFERROR(IF(SelRegion="All",
   XLOOKUP(SelMonth,Calc_Region!$A$5:$A$24,Calc_Region!$N$5:$N$24),
   XLOOKUP(SelMonth,Calc_Region!$A$5:$A$24,
      INDEX(Calc_Region!$I$5:$M$24,0,MATCH(SelRegion,Calc_Region!$I$4:$M$4,0)))),"n/a")
```

## 6. Top và Bottom nhóm hàng (hàm mảng động)

Doanh thu từng nhóm hàng của tháng và khu vực đang chọn:
```
=SORT(UNIQUE(tblSales[Category]))
=SUMIFS(tblSales[Revenue],tblSales[MonthStart],SelMonth,tblSales[Region],RegionCrit,tblSales[Category],D6#)
```
Top 10 và Bottom 5 (chỉ tính nhóm có doanh thu):
```
=TAKE(SORTBY(D6#,E6#,-1),10)
=TAKE(SORT(E6#,,-1),10)
=TAKE(SORTBY(FILTER(D6#,E6#>0),FILTER(E6#,E6#>0),1),5)
```

## 7. Ô chọn và biểu đồ

```
Danh sách khu vực = VSTACK("All",SORT(UNIQUE(tblSales[Region])))
MonthList = Calc!$A$5#
```
Ô chọn tháng và khu vực dùng **Data Validation** kiểu **List** với nguồn là các tên trên. Tiêu đề biểu đồ là tiêu đề động: ô chứa chuỗi như `="Doanh thu theo tháng - "&SelRegion`, biểu đồ liên kết bằng `=Calc_Dash!$X$3`. Chấm nổi bật tháng đang chọn trên biểu đồ đường dùng một chuỗi phụ:
```
=IF(N6#=SelMonth,O6#,NA())
```

## 8. Sheet WhatIf

```
Doanh thu mới = Doanh thu cơ sở * (1 + thay đổi giá) * (1 + thay đổi sản lượng)
Chi phí cơ sở = Doanh thu cơ sở * (1 - biên lợi nhuận gộp)
Chi phí mới = Chi phí cơ sở * (1 + thay đổi sản lượng)
Lợi nhuận gộp mới = Doanh thu mới - Chi phí mới
Thay đổi lợi nhuận gộp (%) = (Lợi nhuận gộp mới - Lợi nhuận gộp cơ sở) / Lợi nhuận gộp cơ sở
```
Bảng hai chiều dùng **Data Table** với ô đầu vào hàng là thay đổi giá và ô đầu vào cột là thay đổi sản lượng. **Goal Seek** tìm mức tăng giá để lợi nhuận gộp không đổi khi sản lượng giảm 10% (khoảng 3,33%, biên 30%).

## 9. PivotTable

- PivotTable 1: `Category` ở Rows; `Revenue` ở Values (Top 10 theo doanh thu).
- PivotTable 2: `Region` ở Rows; `Revenue`, `OrderFlag`, `DeliveredOrder`, `OnTimeOrder` ở Values, kèm trường tính toán `OnTimeRate = OnTimeOrder/DeliveredOrder`.
- Slicer `Region` và `order_status`, Timeline `MonthStart`, kết nối với cả hai PivotTable.

## 10. Ghi chú thiết kế

- Mọi công thức dùng tham chiếu có cấu trúc đến bảng `tblSales`, nên khi làm mới dữ liệu công thức tự mở rộng.
- Một đơn có thể có nhiều sản phẩm. Cột `OrderFlag` giúp đếm đơn và tính các chỉ số theo đơn (ngày giao, điểm đánh giá) mà không tính trùng.
- Ghép đơn hàng với sản phẩm bằng Inner join để loại sản phẩm của các đơn bị hủy. Các bảng bổ sung dùng Left join để không mất dòng.
- Dùng cột ngày `MonthStart` làm điều kiện cho SUMIFS (thay cho chữ `YearMonth`), tránh việc Excel hiểu nhầm chuỗi tháng.
