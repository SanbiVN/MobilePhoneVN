# MobilePhoneVN
> Chuyển đổi và lấy số điện thoại từ số cũ hoặc từ chuỗi số hợp lệ (giai đoạn năm 2018 đến nay)
		
> với Hàm S_GetMobileVN và Hàm S_MobileVN				
				
> Sẽ trả lại mảng chứa kết quả chuyển đổi và tìm số điện thoại trong chuỗi gồm các cột: 				
Đánh thứ tự, số mới, số cũ, định dạng tiêu chuẩn, nhà mạng, chuỗi không hợp lệ				
				
## Hướng dẫn sử dụng hàm:				
	
Cả 2 Hàm đều có chức năng như nhau gồm 12 tham số :			
Riêng hàm S_MobileVN có thêm chức năng in ra mảng	
	
Vị trí	|Tham số	|Kiểu	|Chức năng
------------ | -------------| -------------| -------------
1	|Numbers 	|Chuỗi/Mảng	|Chuỗi hoặc mảng chứa số điện thoại để xử lý
2	|Delimiter 	|Chuỗi	|Ký tự nối chuỗi nếu nhiều số cùng chuỗi
3	|includeInvalid 	|Có/Không	|Trả về kết quả gồm chuỗi không hợp lệ
4	|Expand 	|Có/Không	|Mở rộng xuống hàng mới nếu chuỗi có nhiều số ĐT
5	|ZeroFrontNumber 	|Có/Không	|Giữ lại số 0 khi in mảng
6	|Header 	|Có/Không	|Mảng có đầu đề hay không
7	|ReturnOrder 	|Số nguyên	|Đặt vị trí cột, nếu có cột số thứ tự
8	|ReturnNewPhone 	|Số nguyên	|Đặt vị trí cột, nếu có cột số điện thoại mới
9	|ReturnOldPhone 	|Số nguyên	|Đặt vị trí cột, nếu có cột số điện thoại cũ
10	|ReturnStandardsE164 	|Số nguyên	|Đặt vị trí cột, nếu có cột chuẩn hóa số Điện thoại (E164) 
11	|ReturnCompany 	|Số nguyên	|Đặt vị trí cột, nếu có cột tên Nhà Mạng
12	|ReturnInvalid 	|Số nguyên	|Đặt vị trí cột, nếu có cột chuỗi không hợp lệ
13	|Title 	|Chuỗi	|Đặt chuỗi trả về cho riêng hàm S_MobileVN
	
	
## Cách viết hàm nhanh, gõ vào ô chuỗi =S_MobileVN và ấn tổ hợp phím Ctrl+Shift+A			
## Sẽ nhận được: 
> =S_MobileVN(Numbers,Delimiter,includeInvalid,Expand,ZeroFrontNumber,Header,ReturnOrder,ReturnNewPhone,ReturnOldPhone,ReturnStandardsE164,ReturnCompany,ReturnInvalid,Title)			
				
## Ví dụ cách viết hàm:			
> Trả về giá trị chuỗi số mới, từ số cũ, hoặc lấy các số hợp lệ:		
> =S_GetMobileVN("01681234567")		
> =S_GetMobileVN("03812345670321234567", ",")		
## Trả về kết quả mảng:		
> =S_GetMobileVN(A3:A10,CHAR(10),TRUE,TRUE,TRUE,TRUE,1,2,3,4,5,6)		
## In ra mảng:		
> =S_MobileVN(A3:A10,CHAR(10),TRUE,TRUE,TRUE,TRUE,1,2,3,4,5,6,"Sửa đổi đầu số ĐT")		
> Nếu muốn trả về kết quả cột Số mới và nhà mạng thì:		
> =S_GetMobileVN(A3:A10,CHAR(10),TRUE,TRUE,TRUE,TRUE,0,1,0,0,2)		
> Hoặc =S_GetMobileVN(A3:A10,CHAR(10),TRUE,TRUE,TRUE,TRUE,,1,,,2)		
## Ví dụ code sử dụng hàm:
```VBA			
Dim vArray, rArray, V		
V = S_GetMobileVN(vArray,CHAR(10),TRUE,TRUE,TRUE,TRUE,1,2,3,4,5,6)		
rArray = S_GetMobileVN(vArray,CHAR(10),TRUE,TRUE,TRUE,TRUE,1,2,3,4,5,6)		
```
Hàm S_MobileVN chỉ để in mảng, chỉ sử dụng trong code khi trả kết quả	

