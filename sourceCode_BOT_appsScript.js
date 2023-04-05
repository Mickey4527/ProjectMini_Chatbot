var ss = SpreadsheetApp.openByUrl("Web Google Sheet URL");// ใส่ URL ของ Google Sheet ที่ต้องการใช้งาน

function doPost(e) { // ส่วนของการรับค่าจาก Line และแยกคำสั่ง
  var data = JSON.parse(e.postData.contents); // รับค่าจาก Line และแปลงเป็น JSON
  var userMsg = data.originalDetectIntentRequest.payload.data.message.text; // รับค่าข้อความจากผู้ใช้

  switch(true){ // ส่วนของการแยกคำสั่ง
    case(userMsg === "เนื้อสัตว์" ||userMsg ===  "ผักสด" ||userMsg ===  "ผลไม้"): // ถ้าผู้ใช้พิมพ์เป็นคำสั่งเดียวกัน
    return doPostMenu(userMsg); // ให้เรียกใช้ฟังก์ชัน doPostMenu และส่งค่า userMsg ไปด้วย
    default: // ถ้าไม่ใช่คำสั่งที่กำหนด
    return AddQuetoTable(userMsg); // ให้เรียกใช้ฟังก์ชัน AddQuetoTable และส่งค่า userMsg ไปด้วย
  }
}


function AddQuetoTable(userMsg){ // ฟังก์ชันสำหรับเพิ่มข้อมูลลงใน Google Sheet
  var userCommand = userMsg.split(",");  // แยกคำสั่งด้วยเครื่องหมาย ,
  var sheet = ss.getSheetByName("จองคิว");  // เลือกชีทที่ต้องการใช้งาน
  var values = sheet.getRange(1, 1, sheet.getLastRow(),sheet.getLastColumn()).getValues(); // ดึงข้อมูลทั้งหมดในชีทมาเก็บไว้ในตัวแปร values 
  
    
      sheet.getRange(values.length+1,1).setValue(values.length+1); // ใส่ลำดับลงในช่องที่ 1
      sheet.getRange(values.length+1,2).setValue(userCommand[1]); // ใส่ชื่อลงในช่องที่ 2
      sheet.getRange(values.length+1,3).setValue(userCommand[2]); // ใส่เบอร์โทรลงในช่องที่ 3
      sheet.getRange(values.length+1,4).setValue(Math.ceil(Math.random(8)* userCommand[2])); // ใส่เลขสุ่มลงในช่องที่ 4
      if(sheet.getRange(values.length,5).getValue() === 100){ // ถ้าค่าในช่องที่ 5 เท่ากับ 100
        sheet.getRange(values.length+1,5).setValue('A' + 000); // ให้เริ่มใส่ค่า A000 ลงในช่องที่ 5
      }
      else if(sheet.getRange(values.length,5).getValue() === 'Que'){ // ถ้าค่าในช่องที่ 5 เท่ากับ Que
        sheet.getRange(values.length+1,5).setValue('A' + 000); // ให้เริ่มใส่ค่า A000 ลงในช่องที่ 5
      }
      else{
        var number = parseInt(sheet.getRange(values.length,5).getValue().slice(1)); // แยกตัวเลขออกจากตัวอักษร
        sheet.getRange(values.length+1,5).setValue('A'+ Number(number+1)); // ใส่ค่าลำดับลงในช่องที่ 5
      }
      sheet.getRange(values.length+1,6).setValue("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" + sheet.getRange(values.length+1,4).getValue()); // ใส่ลิงค์ QR Code ลงในช่องที่ 6

// ส่วนของการส่งข้อความกลับไปยัง Line
var result = { 
        "fulfillmentMessages": [{
          "platform": "line",
          "type": 4,
          "payload" : {
            "line": {
                "altText": "This is a Flex message",
                "type": "flex",
                "contents": {
            "type": "bubble",
            "direction": "ltr",
            "header": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                {
                    "type": "text",
                    "text": "จองคิวสำเร็จ !",
                    "align": "center",
                    "contents": []
                }
                ]
            },
            "body": {
                "type": "box",
                "layout": "vertical",
                "backgroundColor": "#F7F7F7FF",
                "contents": [
                {
                    "type": "text",
                    "text": "ลำคับคิวของคุณคือ",
                    "size": "sm",
                    "align": "center",
                    "contents": []
                },
                {
                    "type": "text",
                    "text": sheet.getRange(values.length+1,5).getValue(),
                    "weight": "bold",
                    "size": "3xl",
                    "align": "center",
                    "margin": "xl",
                    "contents": []
                },
                {
                    "type": "image",
                    "url": sheet.getRange(values.length+1,6).getValue(),
                    "margin": "lg"
                },
                {
                    "type": "text",
                    "text": "ชื่อผู้จอง : "+sheet.getRange(values.length+1,2).getValue(),
                    "size": "sm",
                    "margin": "xxl",
                    "align": "start",
                    "wrap": true,
                    "contents": []
                },
                {
                    "type": "text",
                    "text": "เบอร์ผู้จอง : "+sheet.getRange(values.length+1,3).getValue(),
                    "size": "sm",
                    "margin": "md",
                    "contents": []
                }
                ]
            }
                }
            }           
          
            }
          }]
        }


      var replyJSON = ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON); // ส่งข้อความกลับไปยัง Line
      return replyJSON;   // ส่งข้อความกลับไปยัง Line
}


function doPostMenu(userMsg) {     // ส่วนของการแสดงรายการสินค้า
  let sheet = ss.getSheetByName("รายการสินค้า"); // ดึงข้อมูลจาก Sheet ชื่อ รายการสินค้า

  var values = sheet.getRange(2, 1, sheet.getLastRow(),sheet.getLastColumn()).getValues(); // ดึงข้อมูลทั้งหมดใน Sheet รายการสินค้า
  let menu = []; // สร้างตัวแปรเก็บรายการสินค้า
  let menuLink = []; // สร้างตัวแปรเก็บลิงค์รูปสินค้า
  var j = 2 // ตัวแปรเก็บจำนวนแถวของ Sheet รายการสินค้า


  for(var i = 0;i<values.length; i++){ // วนลูปเพื่อดึงข้อมูลรายการสินค้า
      if(values[i][4] == userMsg ){      // ถ้าข้อความที่ผู้ใช้พิมพ์มาตรงกับข้อมูลในช่องที่ 5 ให้ดึงข้อมูลมาแสดง
        menu.push(sheet.getRange(j,2).getValue()); // ดึงข้อมูลชื่อสินค้ามาเก็บในตัวแปร menu
        menuLink.push(sheet.getRange(j,4).getValue());   // ดึงข้อมูลลิงค์รูปสินค้ามาเก็บในตัวแปร menuLink
     }
    j++;
  }
    var result = basechat(cardChat(menu,menuLink)); // ส่งข้อมูลไปยังฟังก์ชัน cardChat และส่งข้อมูลกลับมาเก็บในตัวแปร result
  

    var replyJSON = ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON); // ส่งข้อความกลับไปยัง Line    
    return replyJSON; // ส่งข้อความกลับไปยัง Line
}

function basechat(Data){ // ฟังก์ชันสร้างข้อความแบบ Flex Message
   var result = {    
        "fulfillmentMessages": [  
          {    
          "platform": "line",
          "type": 4,       
          "payload" : {    
            "line": {
               "altText": "รายการอาหาร",
               "quickReply": {
                  "items": [
                  {
                    "type": "action",
                    "action": {
                        "label": "กลับเมนูหลัก",
                        "text": "เมนูหลัก",
                        "type": "message"
                      }
                  },
                  {
                    "type": "action",
                    "action": {
                        "label": "ดูรายการอาหาร",
                        "text": "ดูรายการอาหาร",
                        "type": "message"
                      }
                  },
                  ]
                },
                "type": "flex",
                "contents": {
                  "type": "carousel",
                  "contents":
                    Data 
              }
            }
            }
          }]
      }
      return result; // ส่งข้อมูลกลับไปยังฟังก์ชัน doPostMenu
}

function cardChat(menu,menuLink){ // ฟังก์ชันสร้างข้อความแบบ Flex Message
  var temp;
  var data = [];
    if(menu.length == 0){
            data = {      
              "title":"text",
              "text": "ไม่มีสินค้าในขณะนี้"
          }
        }
    else{
      for(let i = 0; i < menu.length;i++){
        temp = {
       "hero": {
            "url": menuLink[i],
            "aspectRatio": "1:1",
            "size": "full",
            "type": "image",
            "aspectMode": "cover"
          },
          "type": "bubble",
          "direction": "ltr",
          "footer": {
            "contents": [
              {
                "color": "#193C40",
                "type": "button",
                "action": {
                  "text": menu[i],
                  "label": menu[i],
                  "type": "message"
                }
              }
            ],
            "layout": "horizontal",
            "type": "box"
          }
        }
        data.push(temp);
     
  }
}
  return data;
}
