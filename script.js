// ملف: Code.gs - Google Apps Script Backend (ماركت الخير) - مع دعم الكاشير الكامل

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ماركت - لوحة الإدارة')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// دوال إدارة المنتجات (مع cost و stock)
// ==========================================

function getAllProducts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('القائمة');
    
    if (!sheet) {
      sheet = ss.insertSheet('القائمة');
      sheet.appendRow(['معرف', 'الاسم', 'الوصف', 'السعر', 'الصورة', 'متوفر', 'ترتيب', 'معرف الفئة', 'الوحدة', 'الوزن', 'سعر الشراء', 'المخزون']);
      sheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const costColIndex = headers.indexOf('سعر الشراء');
    const stockColIndex = headers.indexOf('المخزون');
    
    const products = [];
    
    for (let i = 1; i < data.length; i++) {
      products.push({
        id: data[i][0],
        name: data[i][1],
        description: data[i][2],
        price: data[i][3],
        image: data[i][4],
        available: data[i][5] || 'نعم',
        order: data[i][6] || i,
        categoryId: data[i][7] || '',
        unit: data[i][8] || '',
        weight: data[i][9] || '',
        cost: costColIndex !== -1 ? (parseFloat(data[i][costColIndex]) || 0) : 0,
        stock: stockColIndex !== -1 ? (parseInt(data[i][stockColIndex]) || 0) : 0
      });
    }
    
    return { success: true, products: products };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addProduct(productData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('القائمة');
    
    if (!sheet) {
      sheet = ss.insertSheet('القائمة');
      sheet.appendRow(['معرف', 'الاسم', 'الوصف', 'السعر', 'الصورة', 'متوفر', 'ترتيب', 'معرف الفئة', 'الوحدة', 'الوزن', 'سعر الشراء', 'المخزون']);
      sheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
    }
    
    const newId = 'P' + new Date().getTime().toString().slice(-6);
    const headers = sheet.getDataRange().getValues()[0];
    const costColIndex = headers.indexOf('سعر الشراء');
    const stockColIndex = headers.indexOf('المخزون');
    
    const rowData = [
      newId,
      productData.name,
      productData.description,
      productData.price,
      productData.image || 'https://images.unsplash.com/photo-1546069901-ba9599a7e63c?w=300',
      productData.available || 'نعم',
      sheet.getLastRow(),
      productData.categoryId || '',
      productData.unit || '',
      productData.weight || ''
    ];
    
    if (costColIndex !== -1) rowData[costColIndex] = productData.cost || 0;
    if (stockColIndex !== -1) rowData[stockColIndex] = productData.stock || 0;
    
    sheet.appendRow(rowData);
    
    return { success: true, message: 'تم إضافة المنتج بنجاح', productId: newId };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateProduct(productData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('القائمة');
    
    if (!sheet) {
      return { success: false, message: 'القائمة غير موجودة' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const costColIndex = headers.indexOf('سعر الشراء');
    const stockColIndex = headers.indexOf('المخزون');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productData.id) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(productData.name);
        sheet.getRange(row, 3).setValue(productData.description);
        sheet.getRange(row, 4).setValue(productData.price);
        sheet.getRange(row, 5).setValue(productData.image);
        sheet.getRange(row, 6).setValue(productData.available);
        sheet.getRange(row, 8).setValue(productData.categoryId);
        sheet.getRange(row, 9).setValue(productData.unit || '');
        sheet.getRange(row, 10).setValue(productData.weight || '');
        
        if (costColIndex !== -1) sheet.getRange(row, costColIndex+1).setValue(productData.cost || 0);
        if (stockColIndex !== -1) sheet.getRange(row, stockColIndex+1).setValue(productData.stock || 0);
        
        return { success: true, message: 'تم تحديث المنتج بنجاح' };
      }
    }
    
    return { success: false, message: 'المنتج غير موجود' };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteProduct(productId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('القائمة');
    
    if (!sheet) {
      return { success: false, message: 'القائمة غير موجودة' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'تم حذف المنتج بنجاح' };
      }
    }
    
    return { success: false, message: 'المنتج غير موجود' };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// دوال إدارة المخزون
// ==========================================

function updateMultipleStocks(updates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('القائمة');
    if (!sheet) {
      return { success: false, message: 'القائمة غير موجودة' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let stockColIndex = headers.indexOf('المخزون');
    if (stockColIndex === -1) {
      sheet.getRange(1, headers.length+1).setValue('المخزون');
      const newHeaders = sheet.getDataRange().getValues()[0];
      stockColIndex = newHeaders.indexOf('المخزون');
      for (let i = 1; i < data.length; i++) {
        sheet.getRange(i+1, stockColIndex+1).setValue(0);
      }
    }
    
    for (const update of updates) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === update.id) {
          const row = i + 1;
          const currentStock = parseInt(sheet.getRange(row, stockColIndex+1).getValue()) || 0;
          const newStock = Math.max(0, currentStock - update.quantity);
          sheet.getRange(row, stockColIndex+1).setValue(newStock);
          break;
        }
      }
    }
    
    return { success: true, message: 'تم تحديث المخزون بنجاح' };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// دوال التقارير اليومية (للمبيعات والأرباح)
// ==========================================

function getDailySalesReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('الطلبات');
    if (!sheet) {
      return { success: true, totalSales: 0, totalProfit: 0, details: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const today = new Date().toLocaleDateString('en-CA');
    
    let totalSales = 0;
    let totalProfit = 0;
    const productMap = new Map();
    
    for (let i = 1; i < data.length; i++) {
      const orderDateStr = data[i][0];
      let orderDate;
      try {
        const parts = orderDateStr.split('،')[0].split('/');
        if (parts.length === 3) {
          orderDate = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
        } else {
          orderDate = new Date(orderDateStr);
        }
      } catch(e) {
        orderDate = new Date(orderDateStr);
      }
      const orderDateISO = orderDate.toLocaleDateString('en-CA');
      if (orderDateISO !== today) continue;
      
      let items = [];
      const itemsRaw = data[i][6];
      try {
        items = JSON.parse(itemsRaw);
      } catch(e) {
        continue;
      }
      
      for (const item of items) {
        const qty = item.quantity;
        const sellingPrice = item.price;
        const costPrice = item.cost || 0;
        const profit = (sellingPrice - costPrice) * qty;
        totalSales += sellingPrice * qty;
        totalProfit += profit;
        
        if (!productMap.has(item.name)) {
          productMap.set(item.name, {
            quantity: 0,
            sellingPrice: sellingPrice,
            costPrice: costPrice,
            profit: 0
          });
        }
        const p = productMap.get(item.name);
        p.quantity += qty;
        p.profit += profit;
        p.sellingPrice = sellingPrice;
        p.costPrice = costPrice;
      }
    }
    
    const details = Array.from(productMap.entries()).map(([name, d]) => ({
      productName: name,
      quantity: d.quantity,
      sellingPrice: d.sellingPrice,
      costPrice: d.costPrice,
      profit: d.profit
    }));
    
    return { success: true, totalSales, totalProfit, details };
    
  } catch (error) {
    return { success: false, message: error.toString(), totalSales: 0, totalProfit: 0, details: [] };
  }
}

// ==========================================
// دوال الفئات
// ==========================================

function getAllCategories() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('الفئات');
    
    if (!sheet) {
      sheet = ss.insertSheet('الفئات');
      sheet.appendRow(['معرف', 'الاسم', 'الترتيب', 'نشط']);
      sheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
      sheet.appendRow(['C1', 'فواكه', 1, 'نعم']);
      sheet.appendRow(['C2', 'خضروات', 2, 'نعم']);
      sheet.appendRow(['C3', 'ألبان', 3, 'نعم']);
      sheet.appendRow(['C4', 'مخبوزات', 4, 'نعم']);
      sheet.appendRow(['C5', 'لحوم ودواجن', 5, 'نعم']);
    }
    
    const data = sheet.getDataRange().getValues();
    const categories = [];
    for (let i = 1; i < data.length; i++) {
      categories.push({
        id: data[i][0],
        name: data[i][1],
        order: data[i][2] || 0,
        active: data[i][3] || 'نعم'
      });
    }
    return { success: true, categories: categories };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addCategory(categoryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('الفئات');
    if (!sheet) {
      sheet = ss.insertSheet('الفئات');
      sheet.appendRow(['معرف', 'الاسم', 'الترتيب', 'نشط']);
    }
    const newId = 'C' + new Date().getTime().toString().slice(-6);
    sheet.appendRow([newId, categoryData.name, categoryData.order || 0, categoryData.active || 'نعم']);
    return { success: true, message: 'تم إضافة الفئة بنجاح', categoryId: newId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateCategory(categoryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('الفئات');
    if (!sheet) return { success: false, message: 'الفئات غير موجودة' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === categoryData.id) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(categoryData.name);
        sheet.getRange(row, 3).setValue(categoryData.order);
        sheet.getRange(row, 4).setValue(categoryData.active);
        return { success: true, message: 'تم تحديث الفئة بنجاح' };
      }
    }
    return { success: false, message: 'الفئة غير موجودة' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteCategory(categoryId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('الفئات');
    if (!sheet) return { success: false, message: 'الفئات غير موجودة' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === categoryId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'تم حذف الفئة بنجاح' };
      }
    }
    return { success: false, message: 'الفئة غير موجودة' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// دوال صور السلايدر
// ==========================================

function getSliderImages() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('السلايدر');
    if (!sheet) {
      sheet = ss.insertSheet('السلايدر');
      sheet.appendRow(['معرف', 'رابط الصورة', 'الترتيب', 'نشط']);
      sheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
      sheet.appendRow(['S1', 'https://images.unsplash.com/photo-1542838132-92c53300491e?w=1200', 1, 'نعم']);
      sheet.appendRow(['S2', 'https://images.unsplash.com/photo-1556905055-8f358a7a47b2?w=1200', 2, 'نعم']);
    }
    const data = sheet.getDataRange().getValues();
    const images = [];
    for (let i = 1; i < data.length; i++) {
      images.push({
        id: data[i][0],
        image_url: data[i][1],
        order: data[i][2] || 0,
        active: data[i][3] || 'نعم'
      });
    }
    return { success: true, images: images };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addSliderImage(imageData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('السلايدر');
    if (!sheet) {
      sheet = ss.insertSheet('السلايدر');
      sheet.appendRow(['معرف', 'رابط الصورة', 'الترتيب', 'نشط']);
    }
    const newId = 'S' + new Date().getTime().toString().slice(-6);
    sheet.appendRow([newId, imageData.image_url, imageData.order || 0, imageData.active || 'نعم']);
    return { success: true, message: 'تم إضافة الصورة بنجاح', imageId: newId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateSliderImage(imageData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('السلايدر');
    if (!sheet) return { success: false, message: 'السلايدر غير موجود' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === imageData.id) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(imageData.image_url);
        sheet.getRange(row, 3).setValue(imageData.order);
        sheet.getRange(row, 4).setValue(imageData.active);
        return { success: true, message: 'تم تحديث الصورة بنجاح' };
      }
    }
    return { success: false, message: 'الصورة غير موجودة' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteSliderImage(imageId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('السلايدر');
    if (!sheet) return { success: false, message: 'السلايدر غير موجود' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === imageId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'تم حذف الصورة بنجاح' };
      }
    }
    return { success: false, message: 'الصورة غير موجودة' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// دوال إعدادات المتجر
// ==========================================

function getRestaurantSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('الإعدادات');
    if (!sheet) {
      sheet = ss.insertSheet('الإعدادات');
      sheet.appendRow(['الإعداد', 'القيمة']);
      sheet.appendRow(['admin_email', Session.getActiveUser().getEmail()]);
      sheet.appendRow(['restaurant_name', 'سوبر ماركت الخير']);
      sheet.appendRow(['phone', '0627278070']);
      sheet.appendRow(['delivery_fee', '10']);
      sheet.appendRow(['min_order', '50']);
      sheet.appendRow(['free_delivery_threshold', '200']);
      sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
    }
    const data = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
    }
    return { success: true, settings: settings };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateRestaurantSettings(settings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('الإعدادات');
    if (!sheet) {
      sheet = ss.insertSheet('الإعدادات');
      sheet.appendRow(['الإعداد', 'القيمة']);
    }
    const existingSettings = sheet.getDataRange().getValues();
    const existingKeys = existingSettings.map(row => row[0]);
    for (const [key, value] of Object.entries(settings)) {
      const rowIndex = existingKeys.indexOf(key);
      if (rowIndex > 0) {
        sheet.getRange(rowIndex + 1, 2).setValue(value);
      } else {
        sheet.appendRow([key, value]);
      }
    }
    return { success: true, message: 'تم تحديث الإعدادات بنجاح' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// دوال الطلبات (العادية والكاشير)
// ==========================================

function saveOrderToSheet(orderData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('الطلبات');
    
    if (!sheet) {
      sheet = ss.insertSheet('الطلبات');
      sheet.appendRow([
        'تاريخ الطلب', 'العميل', 'رقم الهاتف', 'العنوان', 
        'وقت التوصيل', 'طريقة الدفع', 'الأصناف', 'مجموع الأصناف',
        'مبلغ التوصيل', 'المجموع الكلي', 'ملاحظات', 'حالة الطلب', 'معرف الطلب'
      ]);
      sheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
    }
    
    const settings = getRestaurantSettings().settings;
    const deliveryFee = parseFloat(settings.delivery_fee) || 10;
    const freeDeliveryThreshold = parseFloat(settings.free_delivery_threshold) || 200;
    
    const finalDeliveryFee = orderData.subtotal >= freeDeliveryThreshold ? 0 : deliveryFee;
    const totalWithDelivery = orderData.subtotal + finalDeliveryFee;
    
    const itemsForStorage = orderData.items.map(item => ({
      name: item.name,
      quantity: item.quantity,
      price: item.price,
      cost: item.cost || 0
    }));
    const itemsText = JSON.stringify(itemsForStorage);
    
    const orderId = 'ORD-' + new Date().getTime().toString().slice(-8) + '-' + 
                    Math.random().toString(36).substring(2, 6).toUpperCase();
    
    sheet.appendRow([
      orderData.timestamp,
      orderData.customerName,
      orderData.phone,
      orderData.address,
      orderData.deliveryTime,
      orderData.paymentMethod,
      itemsText,
      orderData.subtotal + ' درهم',
      finalDeliveryFee + ' درهم' + (finalDeliveryFee === 0 ? ' (مجاني)' : ''),
      totalWithDelivery + ' درهم',
      orderData.notes || 'لا يوجد',
      'جديد',
      orderId
    ]);
    
    sendEmailNotification(orderData, orderId, finalDeliveryFee, totalWithDelivery, itemsForStorage);
    
    return { 
      success: true, 
      message: 'تم حفظ الطلب بنجاح',
      orderId: orderId 
    };
    
  } catch (error) {
    return { 
      success: false, 
      message: 'حدث خطأ في حفظ الطلب: ' + error.toString() 
    };
  }
}

function saveCashierOrder(orderData) {
  // حفظ طلب الكاشير (بدون توصيل، مع خصم وطريقة دفع)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('الطلبات');
    
    if (!sheet) {
      sheet = ss.insertSheet('الطلبات');
      sheet.appendRow([
        'تاريخ الطلب', 'العميل', 'رقم الهاتف', 'العنوان', 
        'وقت التوصيل', 'طريقة الدفع', 'الأصناف', 'مجموع الأصناف',
        'مبلغ التوصيل', 'المجموع الكلي', 'ملاحظات', 'حالة الطلب', 'معرف الطلب'
      ]);
      sheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
    }
    
    const itemsForStorage = orderData.items.map(item => ({
      name: item.name,
      quantity: item.quantity,
      price: item.price,
      cost: item.cost || 0
    }));
    const itemsText = JSON.stringify(itemsForStorage);
    
    const orderId = 'CSH-' + new Date().getTime().toString().slice(-8) + '-' + 
                    Math.random().toString(36).substring(2, 6).toUpperCase();
    
    // المجموع الكلي بعد الخصم
    const finalTotal = orderData.totalAfterDiscount || orderData.subtotal;
    const notes = `خصم: ${orderData.discount} ${orderData.discountType === 'percent' ? '%' : 'درهم'} | ${orderData.notes || ''}`;
    
    sheet.appendRow([
      orderData.timestamp,
      orderData.customerName || 'كاشير',
      orderData.phone || 'نقدي',
      orderData.address || 'نقطة البيع',
      'فوري',
      orderData.paymentMethod,
      itemsText,
      orderData.subtotal + ' درهم',
      '0 درهم',
      finalTotal + ' درهم',
      notes,
      'مكتمل',
      orderId
    ]);
    
    return { 
      success: true, 
      message: 'تم حفظ طلب الكاشير بنجاح',
      orderId: orderId 
    };
    
  } catch (error) {
    return { 
      success: false, 
      message: 'حدث خطأ في حفظ طلب الكاشير: ' + error.toString() 
    };
  }
}

function getAllOrders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('الطلبات');
    if (!sheet) return { success: false, message: 'لا توجد طلبات بعد' };
    const data = sheet.getDataRange().getValues();
    const orders = [];
    for (let i = 1; i < data.length; i++) {
      orders.push({
        timestamp: data[i][0],
        customerName: data[i][1],
        phone: data[i][2],
        address: data[i][3],
        deliveryTime: data[i][4],
        paymentMethod: data[i][5],
        items: data[i][6],
        subtotal: data[i][7],
        deliveryFee: data[i][8],
        total: data[i][9],
        notes: data[i][10],
        status: data[i][11],
        orderId: data[i][12]
      });
    }
    return { success: true, orders: orders };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getOrderById(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('الطلبات');
    if (!sheet) return { success: false, message: 'لا توجد طلبات' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][12] === orderId) {
        let items = [];
        try { items = JSON.parse(data[i][6]); } catch(e) { items = []; }
        return { 
          success: true, 
          order: {
            orderId: data[i][12],
            timestamp: data[i][0],
            customerName: data[i][1],
            phone: data[i][2],
            address: data[i][3],
            deliveryTime: data[i][4],
            paymentMethod: data[i][5],
            items: items,
            subtotal: parseFloat(data[i][7]) || 0,
            deliveryFee: parseFloat(data[i][8]) || 0,
            total: parseFloat(data[i][9]) || 0,
            notes: data[i][10],
            status: data[i][11]
          }
        };
      }
    }
    return { success: false, message: 'الطلب غير موجود' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateOrderStatus(orderId, newStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('الطلبات');
    if (!sheet) return { success: false, message: 'ورقة الطلبات غير موجودة' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][12] === orderId) {
        sheet.getRange(i + 1, 12).setValue(newStatus);
        return { success: true, message: 'تم تحديث الحالة بنجاح' };
      }
    }
    return { success: false, message: 'الطلب غير موجود' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteAllOrders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('الطلبات');
    if (!sheet) return { success: false, message: 'لا توجد طلبات' };
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return { success: true, message: 'تم مسح جميع الطلبات بنجاح' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function sendEmailNotification(orderData, orderId, deliveryFee, totalWithDelivery, itemsForStorage) {
  try {
    const settings = getRestaurantSettings().settings;
    const adminEmail = settings.admin_email || Session.getActiveUser().getEmail();
    const restaurantName = settings.restaurant_name || 'سوبر ماركت الخير';
    
    let itemsList = '';
    itemsForStorage.forEach(item => {
      itemsList += `
        <tr>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">${item.name}</td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;">${item.quantity}</td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;">${item.price} درهم</td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;">${item.price * item.quantity} درهم</td>
        </tr>
      `;
    });
    
    const htmlBody = `
      <div dir="rtl" style="font-family: 'Tajawal', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #2e7d32; color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
          <h2 style="margin: 0;">🛒 طلب جديد من ${restaurantName}</h2>
          <p style="margin: 5px 0 0;">رقم الطلب: <strong>${orderId}</strong></p>
        </div>
        <div style="background: #f8f9fa; padding: 20px; border: 1px solid #ddd;">
          <h3 style="color: #2e7d32; margin-top: 0;">معلومات العميل</h3>
          <table style="width: 100%; border-collapse: collapse;">
            <tr><td style="padding: 8px; background: #ecf0f1; width: 30%;"><strong>الاسم:</strong></td><td style="padding: 8px;">${orderData.customerName}</td></tr>
            <tr><td style="padding: 8px; background: #ecf0f1;"><strong>رقم الهاتف:</strong></td><td style="padding: 8px;">${orderData.phone}</td></tr>
            <tr><td style="padding: 8px; background: #ecf0f1;"><strong>العنوان:</strong></td><td style="padding: 8px;">${orderData.address}</td></tr>
            <tr><td style="padding: 8px; background: #ecf0f1;"><strong>وقت التوصيل:</strong></td><td style="padding: 8px;">${orderData.deliveryTime}</td></tr>
            <tr><td style="padding: 8px; background: #ecf0f1;"><strong>طريقة الدفع:</strong></td><td style="padding: 8px;">${orderData.paymentMethod}</td></tr>
           </table>
          <h3 style="color: #2e7d32; margin-top: 20px;">تفاصيل الطلب</h3>
          <table style="width: 100%; border-collapse: collapse;">
            <thead><tr style="background: #2e7d32; color: white;"><th style="padding: 8px;">الصنف</th><th style="padding: 8px;">الكمية</th><th style="padding: 8px;">السعر</th><th style="padding: 8px;">المجموع</th></tr></thead>
            <tbody>${itemsList}</tbody>
          </table>
          <div style="margin-top: 15px; padding: 15px; background: white; border-radius: 8px;">
            <table style="width: 100%;">
              <tr><td>مجموع الأصناف:</td><td style="text-align: left;"><strong>${orderData.subtotal} درهم</strong></td></tr>
              <tr><td>مبلغ التوصيل:</td><td style="text-align: left;"><strong>${deliveryFee} درهم ${deliveryFee === 0 ? '(مجاني)' : ''}</strong></td></tr>
              <tr style="font-size: 18px; color: #2e7d32;"><td>المجموع الكلي:</td><td style="text-align: left;"><strong>${totalWithDelivery} درهم</strong></td></tr>
            </table>
          </div>
          ${orderData.notes ? `<div style="margin-top: 20px; padding: 10px; background: #fff3cd; border: 1px solid #ffeeba; border-radius: 5px;"><strong>📝 ملاحظات:</strong><p style="margin: 5px 0 0;">${orderData.notes}</p></div>` : ''}
        </div>
        <div style="background: #ecf0f1; padding: 15px; text-align: center; border-radius: 0 0 10px 10px;">
          <p style="margin: 0;">تم استلام الطلب في: ${orderData.timestamp}</p>
          <p style="margin: 5px 0 0; color: #2e7d32;">يرجى تجهيز الطلب والتواصل مع العميل</p>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: adminEmail,
      subject: `🛒 طلب جديد - ${orderId}`,
      htmlBody: htmlBody
    });
    
  } catch (error) {
    console.error('خطأ في إرسال الإشعار:', error.toString());
  }
}

// ==========================================
// دوال العملاء (جديدة)
// ==========================================

function saveCustomerIfNew(customer) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('العملاء');
    if (!sheet) {
      sheet = ss.insertSheet('العملاء');
      sheet.appendRow(['رقم الهاتف', 'الاسم', 'العنوان', 'آخر طلب', 'عدد الطلبات']);
      sheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
    }
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === customer.phone) {
        found = true;
        const row = i + 1;
        const currentOrders = parseInt(sheet.getRange(row, 5).getValue()) || 0;
        sheet.getRange(row, 5).setValue(currentOrders + 1);
        sheet.getRange(row, 4).setValue(new Date().toLocaleString());
        break;
      }
    }
    
    if (!found && customer.phone && customer.phone !== 'غير مدخل') {
      sheet.appendRow([
        customer.phone,
        customer.name || 'زائر',
        customer.address || '',
        new Date().toLocaleString(),
        1
      ]);
    }
    
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// دوال التحقق والتهيئة
// ==========================================

function initialize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // الإعدادات
  if (!ss.getSheetByName('الإعدادات')) {
    const settingsSheet = ss.insertSheet('الإعدادات');
    settingsSheet.appendRow(['الإعداد', 'القيمة']);
    settingsSheet.appendRow(['admin_email', Session.getActiveUser().getEmail()]);
    settingsSheet.appendRow(['restaurant_name', 'ماركت الخير']);
    settingsSheet.appendRow(['phone', '0627278070']);
    settingsSheet.appendRow(['delivery_fee', '10']);
    settingsSheet.appendRow(['min_order', '50']);
    settingsSheet.appendRow(['free_delivery_threshold', '200']);
    settingsSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
  }
  
  // القائمة مع الأعمدة الجديدة
  let menuSheet = ss.getSheetByName('القائمة');
  if (!menuSheet) {
    menuSheet = ss.insertSheet('القائمة');
    menuSheet.appendRow(['معرف', 'الاسم', 'الوصف', 'السعر', 'الصورة', 'متوفر', 'ترتيب', 'معرف الفئة', 'الوحدة', 'الوزن', 'سعر الشراء', 'المخزون']);
    menuSheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
    
    const defaultProducts = [
      ['P1001', 'تفاح أحمر', 'طازج مصري', 12, 'https://images.unsplash.com/photo-1567306226416-28f0efdc88ce?w=300', 'نعم', 1, 'C1', 'كيلو', '1 كجم', 8, 50],
      ['P1002', 'موز', 'درجة أولى', 8, 'https://images.unsplash.com/photo-1603833665858-e61d17a86224?w=300', 'نعم', 2, 'C1', 'كيلو', '1 كجم', 5, 40],
      ['P1003', 'حليب طازج', 'قليل الدسم', 6, 'https://images.unsplash.com/photo-1563636619-e9143da4abd4?w=300', 'نعم', 3, 'C3', 'علبة', '1 لتر', 4, 30],
      ['P1004', 'خبز عربي', 'طازج يومياً', 2, 'https://images.unsplash.com/photo-1509440159596-0249088772ff?w=300', 'نعم', 4, 'C4', 'ربطة', '4 قطع', 1.2, 100],
      ['P1005', 'طماطم', 'عضوي', 5, 'https://images.unsplash.com/photo-1592924357228-91a4daadcfea?w=300', 'نعم', 5, 'C2', 'كيلو', '1 كجم', 3, 60],
      ['P1006', 'دجاج مجمد', 'طازج', 25, 'https://images.unsplash.com/photo-1587593810167-a84920fde1a1?w=300', 'نعم', 6, 'C5', 'كيلو', '1 كجم', 18, 20]
    ];
    defaultProducts.forEach(product => menuSheet.appendRow(product));
  } else {
    // التأكد من وجود عمودي سعر الشراء والمخزون
    const headers = menuSheet.getRange(1, 1, 1, menuSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('سعر الشراء') === -1) {
      menuSheet.getRange(1, headers.length+1).setValue('سعر الشراء');
      const lastRow = menuSheet.getLastRow();
      if (lastRow > 1) menuSheet.getRange(2, headers.length+1, lastRow-1, 1).setValue(0);
    }
    if (headers.indexOf('المخزون') === -1) {
      const newHeaders = menuSheet.getRange(1, 1, 1, menuSheet.getLastColumn()).getValues()[0];
      menuSheet.getRange(1, newHeaders.length+1).setValue('المخزون');
      const lastRow = menuSheet.getLastRow();
      if (lastRow > 1) menuSheet.getRange(2, newHeaders.length+1, lastRow-1, 1).setValue(0);
    }
  }
  
  // الفئات
  if (!ss.getSheetByName('الفئات')) {
    const catSheet = ss.insertSheet('الفئات');
    catSheet.appendRow(['معرف', 'الاسم', 'الترتيب', 'نشط']);
    const defaultCategories = [
      ['C1', 'فواكه', 1, 'نعم'],
      ['C2', 'خضروات', 2, 'نعم'],
      ['C3', 'ألبان', 3, 'نعم'],
      ['C4', 'مخبوزات', 4, 'نعم'],
      ['C5', 'لحوم ودواجن', 5, 'نعم']
    ];
    defaultCategories.forEach(cat => catSheet.appendRow(cat));
    catSheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
  }
  
  // السلايدر
  if (!ss.getSheetByName('السلايدر')) {
    const sliderSheet = ss.insertSheet('السلايدر');
    sliderSheet.appendRow(['معرف', 'رابط الصورة', 'الترتيب', 'نشط']);
    sliderSheet.appendRow(['S1', 'https://images.unsplash.com/photo-1542838132-92c53300491e?w=1200', 1, 'نعم']);
    sliderSheet.appendRow(['S2', 'https://images.unsplash.com/photo-1556905055-8f358a7a47b2?w=1200', 2, 'نعم']);
    sliderSheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
  }
  
  // الطلبات
  if (!ss.getSheetByName('الطلبات')) {
    const ordersSheet = ss.insertSheet('الطلبات');
    ordersSheet.appendRow([
      'تاريخ الطلب', 'العميل', 'رقم الهاتف', 'العنوان', 
      'وقت التوصيل', 'طريقة الدفع', 'الأصناف', 'مجموع الأصناف',
      'مبلغ التوصيل', 'المجموع الكلي', 'ملاحظات', 'حالة الطلب', 'معرف الطلب'
    ]);
    ordersSheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
  }
  
  // العملاء
  if (!ss.getSheetByName('العملاء')) {
    const customersSheet = ss.insertSheet('العملاء');
    customersSheet.appendRow(['رقم الهاتف', 'الاسم', 'العنوان', 'آخر طلب', 'عدد الطلبات']);
    customersSheet.getRange('1:1').setFontWeight('bold').setBackground('#2e7d32').setFontColor('#ffffff');
  }
  
  return 'تم تهيئة جميع الأوراق بنجاح';
}

function verifyAdminPassword(password) {
  const ADMIN_PASSWORD = 'admin1041';
  return password === ADMIN_PASSWORD;
}
