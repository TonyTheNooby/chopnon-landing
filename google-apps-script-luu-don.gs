function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DonHang');

    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('DonHang');
      sheet.appendRow([
        'Thời gian',
        'Mã đơn',
        'Trạng thái',
        'Phương thức thanh toán',
        'Họ tên khách',
        'Số điện thoại',
        'Email',
        'Địa chỉ',
        'Sản phẩm',
        'Tạm tính',
        'Phí ship',
        'Giảm giá',
        'Tổng tiền'
      ]);
    }

    var data = JSON.parse(e.postData.contents);
    var itemsText = (data.items || []).map(function(item) {
      return (item.name || 'Sản phẩm') + ' x ' + (item.qty || 1) + ' = ' + ((item.price || 0) * (item.qty || 1));
    }).join('\n');

    sheet.appendRow([
      new Date(),
      data.orderId || '',
      data.status || 'Chờ xác nhận',
      data.paymentMethod || 'COD',
      data.customer && data.customer.name || '',
      data.customer && data.customer.phone || '',
      data.customer && data.customer.email || '',
      data.customer && data.customer.address || '',
      itemsText,
      data.subtotal || 0,
      data.shipping || 0,
      data.discount || 0,
      data.total || 0
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
