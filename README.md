# MobilePhoneVN
> Chuyển đổi và lấy số điện thoại từ số cũ hoặc từ chuỗi số hợp lệ (giai đoạn năm 2018 đến nay)
		
> với Hàm PhoneVN				
				
> Sẽ trả lại mảng chứa kết quả chuyển đổi và tìm số điện thoại trong chuỗi gồm các cột: 				
Đánh thứ tự, số mới, số cũ, định dạng tiêu chuẩn, nhà mạng, chuỗi không hợp lệ


![Image](https://github.com/SanbiVN/MobilePhoneVN/blob/main/mobile_phone_VN.gif)
			
			
## Hướng dẫn sử dụng hàm:				
=PhoneVN([Số/Danh_sách],[Đối_số_cài_đặt]

 Tham số	|Kiểu	|Chức năng
------------- | ------------- | -------------
=epDelimiter(",") 	| 	|Ký tự nối chuỗi nếu nhiều số cùng chuỗi
=epincludeInvalid() 	| 	|Trả về kết quả gồm chuỗi không hợp lệ
=epExpand() 	| 	|Mở rộng xuống hàng mới nếu chuỗi có nhiều số ĐT
=epZeroFrontNumber() 	|	|Giữ lại số 0 khi in mảng
=epHeader() 	|	|Mảng có đầu đề hay không
=epColumns(1,2,3,4,5,6)	| 1 	|Đặt vị trí cột, nếu có cột số thứ tự
2	|  	|Đặt vị trí cột, nếu có cột số điện thoại mới
3 	|  	|Đặt vị trí cột, nếu có cột số điện thoại cũ
4 	|  	|Đặt vị trí cột, nếu có cột chuẩn hóa số Điện thoại (E164) 
5 	|  	|Đặt vị trí cột, nếu có cột tên Nhà Mạng
6 	|  	|Đặt vị trí cột, nếu có cột chuỗi không hợp lệ
	
Tệp MobilePhoneVN.xlsm tách số điện thoại di động.
Tệp HomePhoneVN.xlsm tách số điện thoại cố định.

