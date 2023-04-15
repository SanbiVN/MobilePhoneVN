# MobilePhoneVN
> Chuyển đổi và lấy số điện thoại từ số cũ hoặc từ chuỗi số hợp lệ (giai đoạn năm 2018 đến nay)
		
> với Hàm PhoneVN				
				
> Sẽ trả lại mảng chứa kết quả chuyển đổi và tìm số điện thoại trong chuỗi gồm các cột: 				
Đánh thứ tự, số mới, số cũ, định dạng tiêu chuẩn, nhà mạng, chuỗi không hợp lệ


![Image](https://github.com/SanbiVN/MobilePhoneVN/blob/main/mobile_phone_VN.gif)
			
			
## Hướng dẫn sử dụng hàm:				
=PhoneVN([Số/Danh_sách],[Đối_số_cài_đặt]
		
Tham số	|Kiểu	|Chức năng
-------------| -------------| -------------| -------------
=epDelimiter(",") 	|Chuỗi	|Ký tự nối chuỗi nếu nhiều số cùng chuỗi
=epincludeInvalid() 	|Có/Không	|Trả về kết quả gồm chuỗi không hợp lệ
=epExpand() 	|Có/Không	|Mở rộng xuống hàng mới nếu chuỗi có nhiều số ĐT
=epZeroFrontNumber() 	|Có/Không	|Giữ lại số 0 khi in mảng
=epHeader() 	|Có/Không	|Mảng có đầu đề hay không
epColumns	|ReturnOrder 	|Số nguyên	|Đặt vị trí cột, nếu có cột số thứ tự
	|ReturnNewNumber 	|Số nguyên	|Đặt vị trí cột, nếu có cột số điện thoại mới
	|ReturnOldNumber 	|Số nguyên	|Đặt vị trí cột, nếu có cột số điện thoại cũ
	|ReturnStandardsE164 	|Số nguyên	|Đặt vị trí cột, nếu có cột chuẩn hóa số Điện thoại (E164) 
	|ReturnCompany 	|Số nguyên	|Đặt vị trí cột, nếu có cột tên Nhà Mạng
	|ReturnInvalid 	|Số nguyên	|Đặt vị trí cột, nếu có cột chuỗi không hợp lệ
	


