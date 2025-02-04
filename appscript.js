function doGet() {
  var sheet = SpreadsheetApp.openById("1HWk8s__aXVkSLs4iShYuAdaiI231ze0-GDN52dcNN_A").getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var jsonData = {
    left_table: [],
    right_table: {
      table1: [],
      table2: []
    }
  };

  for (var i = 1; i < data.length; i++) {
    if (data[i][8] == "") break;
    data[i][8] = 0;
    sheet.getRange(i + 1, 9).setValue(0);
  }

  // 讀取上一次的最後索引 (J2)
  var lastIndexCell = Number(sheet.getRange("J2").getValue()) || 0;
  var ii = lastIndexCell + 1;
  if (ii >= data.length) ii = 1; // 避免超出範圍

  var sum = 0;
  var previousResults = [];

  while (sum < 10) {
    if (ii > Number(sheet.getRange("H1").getValue())) ii = 1;
    var x = data[ii] ? data[ii][7] : "";
    data[ii][8] = 1;
    sheet.getRange(ii + 1, 9).setValue(1);
    ii++;

    if (ii > Number(sheet.getRange("H1").getValue())) ii = 1;
    var y = data[ii] ? data[ii][7] : "";
    data[ii][8] = 1;
    sheet.getRange(ii + 1, 9).setValue(1);
    ii++;

    jsonData.left_table.push({ col1: x, col2: y });

    previousResults.push([x, y]); // 記錄此次安排結果

    sum += 2;
  }

  // **將上一次的 A1~A5, B1~B5 移動到 A10~A14, B10~B14**
  var lastResults = sheet.getRange("A1:B5").getValues();
  sheet.getRange("A10:B14").setValues(lastResults);

  // **將新結果存到 A1~A5, B1~B5**
  sheet.getRange("A1:B5").setValues(previousResults);

  sheet.getRange("J2").setValue(ii - 1);
  sheet.getRange("J3").setValue(data[ii - 1][7]);




  // 右邊值日  

  // 取得試算表資料
  var previousRightTable = sheet.getRange("D1:E2").getValues();
  sheet.getRange("D10:E11").setValues(previousRightTable); // 上一次

  var last = Number(sheet.getRange("L2").getValue()) || 0;
  sum = 0;
  if (last == Number(sheet.getRange("O5").getValue())) last = 0; // 防止無限循環
  last++;

  var cellMapping = [
    { row: 1, col: 4 },
    { row: 2, col: 4 },
    { row: 1, col: 5 },
    { row: 2, col: 5 }
  ];

  // 取得 Q 欄現有值，並同步更新 data 陣列
  var qValues = sheet.getRange("Q1:Q" + sheet.getLastRow()).getValues().filter(row => row[0] !== "");
  var qlength = qValues.length;

  while (sum < 4) {
    var t = 1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][15] == "") break;
      if (data[i][15] == last) {  
        last++;
        t = 0;
        break;
      }
    }

    if (t == 1) {
      qValues = sheet.getRange("Q1:Q" + sheet.getLastRow()).getValues().filter(row => row[0] !== "");
      if (qlength > 0 && sum < 4) {
        while (qlength > 0 && sum < 4) {
          var valueToMove = qValues[0][0]; // 取得 Q 欄第一個元素

          // 取得左側(A1:A5)和右側(B1:B5)的值
          var leftTableValues = sheet.getRange("A1:A5").getValues().flat().filter(v => v !== "" && v !== undefined);
          var rightTableValues = sheet.getRange("B1:B5").getValues().flat().filter(v => v !== "" && v !== undefined);

          // 檢查 valueToMove 是否已經存在於 A1:A5 或 B1:B5
          if (!leftTableValues.includes(valueToMove) && 
              !rightTableValues.includes(valueToMove) && 
              valueToMove != Number(sheet.getRange("D1").getValue()) && 
              valueToMove != Number(sheet.getRange("D2").getValue()) &&
              valueToMove != Number(sheet.getRange("E1").getValue()) &&
              valueToMove != Number(sheet.getRange("E2").getValue())) {
            // 若可移動，則設定到 cellMapping 對應位置
            if (cellMapping[sum]) {
              sheet.getRange(cellMapping[sum].row, cellMapping[sum].col).setValue(valueToMove);
              data[cellMapping[sum].row - 1][cellMapping[sum].col - 1] = valueToMove; // 更新 data 陣列
              sum++; 

              // 更新 qValues 陣列，並同步更新 data 陣列
              qValues = qValues.filter(row => row[0] !== valueToMove);
            }
          } 
          qlength--;
        }

        // 清空 Q 欄，並同步更新 data 陣列
        sheet.getRange("Q1:Q" + sheet.getLastRow()).clearContent();
        for (var i = 0; i < data.length; i++) {
          data[i][16] = ""; // 清空 Q 欄 (Q = 第 17 欄，索引 16)
        }

        // 如果 qValues 還有剩餘的元素，則寫回 Q 欄並同步 data 陣列
        if (qValues.length > 0) {
          sheet.getRange("Q1:Q" + qValues.length).setValues(qValues.map(value => [value[0]]));
          for (var i = 0; i < qValues.length; i++) {
            data[i][16] = qValues[i][0]; // 更新 data 陣列
          }
        }

      } else {
        var leftTableValues = sheet.getRange("A1:A5").getValues().flat().filter(v => v !== "" && v !== undefined);
        var rightTableValues = sheet.getRange("B1:B5").getValues().flat().filter(v => v !== "" && v !== undefined);

        if ((leftTableValues.includes(last) || rightTableValues.includes(last)) && 
            last != Number(sheet.getRange("D1").getValue()) && 
            last != Number(sheet.getRange("D2").getValue()) && 
            last != Number(sheet.getRange("E1").getValue()) && 
            last != Number(sheet.getRange("E2").getValue())) {
          sheet.getRange("Q" + (qValues.length + 1)).setValue(last); // 放入 Q 欄空欄
          data[qValues.length][16] = last; // 更新 data 陣列

        } 
        else if (!leftTableValues.includes(last) && 
                !rightTableValues.includes(last) && 
                last != Number(sheet.getRange("D1").getValue()) && 
                last != Number(sheet.getRange("D2").getValue()) && 
                last != Number(sheet.getRange("E1").getValue()) && 
                last != Number(sheet.getRange("E2").getValue())) {
          if (cellMapping[sum]) {
            sheet.getRange(cellMapping[sum].row, cellMapping[sum].col).setValue(last);
            data[cellMapping[sum].row - 1][cellMapping[sum].col - 1] = last; // 更新 data 陣列
            sum++;
            sheet.getRange("L2").setValue(last);
          }
        }
        last++;
      }
    }
  }











  /*
  // 取得試算表資料
var previousRightTable = sheet.getRange("D1:E2").getValues();
sheet.getRange("D10:E11").setValues(previousRightTable); // 上一次

var last = Number(sheet.getRange("L2").getValue()) || 0;
sum = 0;
if (last == Number(sheet.getRange("O5").getValue())) last = 0; // 防止無限循環
last++;

var cellMapping = [
  { row: 1, col: 4 },
  { row: 2, col: 4 },
  { row: 1, col: 5 },
  { row: 2, col: 5 }
];

// 取得 Q 欄現有值，並同步更新 data 陣列
var qValues = sheet.getRange("Q1:Q" + sheet.getLastRow()).getValues().filter(row => row[0] !== "");
var qlength = qValues.length;

while (sum < 4) {
  var t = 1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][15] == "") break;
    if (data[i][15] == last) {  
      last++;
      t = 0;
      break;
    }
  }

  if (t == 1) {
    if (qlength > 0 && sum < 4) {
      while (qlength > 0 && sum < 4) {
        var valueToMove = qValues[0][0]; // 取得 Q 欄第一個元素

        // 取得左側(A1:A5)和右側(B1:B5)的值
        var leftTableValues = sheet.getRange("A1:A5").getValues().flat().filter(v => v !== "" && v !== undefined);
        var rightTableValues = sheet.getRange("B1:B5").getValues().flat().filter(v => v !== "" && v !== undefined);

        // 檢查 valueToMove 是否已經存在於 A1:A5 或 B1:B5
        if (!leftTableValues.includes(valueToMove) && 
            !rightTableValues.includes(valueToMove) && 
            valueToMove != Number(sheet.getRange("D1").getValue()) && 
            valueToMove != Number(sheet.getRange("D2").getValue()) &&
            valueToMove != Number(sheet.getRange("E1").getValue()) &&
            valueToMove != Number(sheet.getRange("E2").getValue())) {
          // 若可移動，則設定到 cellMapping 對應位置
          if (cellMapping[sum]) {
            sheet.getRange(cellMapping[sum].row, cellMapping[sum].col).setValue(valueToMove);
            data[cellMapping[sum].row - 1][cellMapping[sum].col - 1] = valueToMove; // 更新 data 陣列
            sum++; 

            // 更新 qValues 陣列，並同步更新 data 陣列
            qValues = qValues.filter(row => row[0] !== valueToMove);
          }
        } 
      }

      // 清空 Q 欄，並同步更新 data 陣列
      sheet.getRange("Q1:Q" + sheet.getLastRow()).clearContent();
      for (var i = 0; i < data.length; i++) {
        data[i][16] = ""; // 清空 Q 欄 (Q = 第 17 欄，索引 16)
      }

      // 如果 qValues 還有剩餘的元素，則寫回 Q 欄並同步 data 陣列
      if (qValues.length > 0) {
        sheet.getRange("Q1:Q" + qValues.length).setValues(qValues.map(value => [value[0]]));
        for (var i = 0; i < qValues.length; i++) {
          data[i][16] = qValues[i][0]; // 更新 data 陣列
        }
      }

      qlength--;
    } else {
      var leftTableValues = sheet.getRange("A1:A5").getValues().flat().filter(v => v !== "" && v !== undefined);
      var rightTableValues = sheet.getRange("B1:B5").getValues().flat().filter(v => v !== "" && v !== undefined);

      if ((leftTableValues.includes(last) || rightTableValues.includes(last)) && 
          last != Number(sheet.getRange("D1").getValue()) && 
          last != Number(sheet.getRange("D2").getValue()) && 
          last != Number(sheet.getRange("E1").getValue()) && 
          last != Number(sheet.getRange("E2").getValue())) {
        sheet.getRange("Q" + (qValues.length + 1)).setValue(last); // 放入 Q 欄空欄
        data[qValues.length][16] = last; // 更新 data 陣列
      } 
      else if (!leftTableValues.includes(last) && 
               !rightTableValues.includes(last) && 
               last != Number(sheet.getRange("D1").getValue()) && 
               last != Number(sheet.getRange("D2").getValue()) && 
               last != Number(sheet.getRange("E1").getValue()) && 
               last != Number(sheet.getRange("E2").getValue())) {
        if (cellMapping[sum]) {
          sheet.getRange(cellMapping[sum].row, cellMapping[sum].col).setValue(last);
          data[cellMapping[sum].row - 1][cellMapping[sum].col - 1] = last; // 更新 data 陣列
          sum++;
        }
      }
      last++;
    }
  }
}


  /*
  var previousRightTable = sheet.getRange("D1:E2").getValues();
  sheet.getRange("D10:E11").setValues(previousRightTable);// 上一次

  var last = Number(sheet.getRange("L2").getValue()) || 0;
  sum=0;
  if (last == Number(sheet.getRange("O5").getValue())) last = 0;  // 防止無限循環
  last++;

  var cellMapping = [
    { row: 1, col: 4 },
    { row: 2, col: 4 },
    { row: 1, col: 5 },
    { row: 2, col: 5 }
  ];
  var qValues = sheet.getRange("Q1:Q" + sheet.getLastRow()).getValues().filter(row => row[0] !== "");
  var qlength = qValues.length;

  while(sum <4 ){ 
    var t = 1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][15] == "") break;
      if (data[i][15] == last) {  
        last++;
        t = 0;
        break;
      }
    }

    if (t == 1) {

      if (qlength > 0 && sum < 4) {
        // 使用 while 進行處理，每次從 Q 欄讀取一個值
        while (qlength > 0 && sum < 4) {
          var valueToMove = qValues[0][0]; // 取得 Q 欄的第一個元素

          // 取得左邊(A1:A5)和右邊(B1:B5)的值
          var leftTableValues = sheet.getRange("A1:A5").getValues().flat().filter(v => v !== "" && v !== undefined);
          var rightTableValues = sheet.getRange("B1:B5").getValues().flat().filter(v => v !== "" && v !== undefined);

          // 檢查 valueToMove 是否已經存在於 A1:A5 或 B1:B5
          if (!leftTableValues.includes(valueToMove) && !rightTableValues.includes(valueToMove) && valueToMove!=Number(sheet.getRange("D1").getValue()) && valueToMove!=Number(sheet.getRange("D2").getValue())  && valueToMove!=Number(sheet.getRange("E1").getValue())  && valueToMove!=Number(sheet.getRange("E2").getValue()) ) {
            // 如果 valueToMove 不在 A1:A5 和 B1:B5 中，則可以將其移動
            if (cellMapping[sum]) {
              sheet.getRange(cellMapping[sum].row, cellMapping[sum].col).setValue(valueToMove);
              sum++; 
              qValues = qValues.filter(row => row[0] !== valueToMove);// 移除已處理的 Q 欄元素
            }
          } 
        }

        // 清空 Q 欄
        sheet.getRange("Q1:Q" + sheet.getLastRow()).clearContent();

        // 如果 qValues 還有剩餘的元素，則將剩餘的 Q 欄資料放回
        if (qValues.length > 0) {
          sheet.getRange("Q1:Q" + qValues.length).setValues(qValues.map(value => [value[0]]));
        }

        qlength--;
      } else {
        var leftTableValues = sheet.getRange("A1:A5").getValues().flat().filter(v => v !== "" && v !== undefined);
        var rightTableValues = sheet.getRange("B1:B5").getValues().flat().filter(v => v !== "" && v !== undefined);

        if ((leftTableValues.includes(last) || rightTableValues.includes(last)) && last!=Number(sheet.getRange("D1").getValue()) && last!=Number(sheet.getRange("D2").getValue())  && last!=Number(sheet.getRange("E1").getValue())  && last!=Number(sheet.getRange("E2").getValue())) {
          sheet.getRange("Q" + qValues.length +1).setValue(last);  // 放入 Q 欄空欄
        }
        // 檢查號碼是否存在於 A1:A5 或 B1:B5
        else if (!leftTableValues.includes(last) && !rightTableValues.includes(last) && last!=Number(sheet.getRange("D1").getValue()) && last!=Number(sheet.getRange("D2").getValue())  && last!=Number(sheet.getRange("E1").getValue())  && last!=Number(sheet.getRange("E2").getValue())) {
          if (cellMapping[sum]) {
            sheet.getRange(cellMapping[sum].row, cellMapping[sum].col).setValue(last);
            sum++;
          }
        }

        last++;
      }

    }

  }
*/








  // **讀取右半邊表格**
  for (var j = 0; j < Math.min(2, data.length); j++) {
    if (data[j]) {
      jsonData.right_table.table1.push(data[j][3] || "");
      jsonData.right_table.table2.push(data[j][4] || "");
    }
  }

  var output = ContentService.createTextOutput(JSON.stringify(jsonData));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
