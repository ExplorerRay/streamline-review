function createNewForm2(){
  let yr = '112'; //年度記得修改
  // 下學期
  let new_form_2 = FormApp.create(yr + "-2 課程查核"); 
  let new_form_2_id = new_form_2.getId();

  let destinationFolderURL = "https://drive.google.com/drive/folders/1JbrHbkUuiF3wOYXufaDkC3nTRzgWvQfD"; //須根據情況修改
  var tmp = destinationFolderURL.split("/");
  let destinationFolderId = tmp[tmp.length - 1];
  moveFile(new_form_2_id, destinationFolderId);
  
  SetForm2(new_form_2, yr)
}

function moveFile(fileId, destinationFolderId) {
  let destinationFolder = DriveApp.getFolderById(destinationFolderId);
  DriveApp.getFileById(fileId).moveTo(destinationFolder);
}

function SetForm2(form, yr){
  var cbValidation = FormApp.createCheckboxValidation()
  .requireSelectAtLeast(1)
  .build();

  // 需修改description，成果報告之雲端連結
  //   let des = yr+'-2課程成果報告：\n\
  //   https://drive.google.com/drive/folders/1LLBxWHjNT9wq6ryua6LDEiFmeY4lxlyR\n\
  // \n\
  // 說明如下：\n\
  // \n\
  // 一、 考評重點：\n\
  // \n\
  //     1.課程的備課狀況\n\
  // \n\
  //     2.與推廣模組整合的規劃\n\
  // \n\
  //     3.教學設備採購情況\n\
  // \n\
  // \n\
  // 二、 評等分數：優(4分)、佳(3分)、可(2分)、差(1分)';
  //   form.setDescription(des);
  form.setCollectEmail(true);

  let md_num = ['A-4', 'A-5', 'A-6', 'A-7', 'A-8', 'B-2', 'B-3', 'B-4', 'B-5', 'B-6', 'B-7', 'B-8', 'B-9', 'C-1', 'C-2', 'C-3', 'C-4', 'C-6', 'C-7', 'C-8', 'C-9', 'C-10', 'C-11', 'C-12', 'D-1', 'D-2', 'D-3', 'D-4', 'D-5', 'D-6', 'E-4', 'E-5', 'E-6'];
  let md_nam = ['晶片及硬體之邏輯暨架構層次的資安防護設計', '晶片及硬體之供應鏈層次的進階資安防護設計課程', '智慧晶片系統生醫領域應用之安全性規範簡介模組教材開發', '機器學習預測IR電壓降', 'Handling Placement Constraints in Analog Layout Synthesis', '醫用智慧系統與電子感測晶片整合 \
  設計', '智慧健康微感測系統', '低功耗線性及切換式穩壓器設計', '能源擷取電路設計', '智慧感測晶片之類比數位轉換電路', '健康促進應用開發專題', '基因資訊探勘與序列比對晶片設計', '硬體計算在生物 \
  資訊學上的應用', '數位系統的高階合成設計方法', 'AI加速器設計概論與實務', '智慧影像處理AI加速器 \
  設計', '智慧終端裝置影像處理晶片設計', '智慧型自走載具系統與晶片設計模組', '人體活動辨識和非接 \
  觸式體溫量測模組', '近記憶體運算及記憶體內運算電路設計', '語音辨識系統', '加速TinyML模型於微控 \
  制器之方法設計與實作', '邊緣AI加速器設計與實作', '軟硬體協同設計之人工智慧晶片設計', '微型環境感測介面電路設計與應用', ' \
  環境能量擷取電路晶片設計', '功率管理模組', '空品與水質感測晶片技術', '低功耗無線感控節點', '應 \
  用於土壤成分監測之感測介面電路設計', '多元自駕車空間感知技術與實作', '模式預測控制技術於自動駕 \
  駛系統之應用', '室外定位融合系統模擬與實作'];
  let md=[];
  for (let l=0; l<md_nam.length; l++) {
    md.push(md_num[l].concat(' ', md_nam[l]));
  }

  let sht = SpreadsheetApp.getActiveSheet();
  let startRow = 5;
  let numRows = sht.getLastRow() - startRow +1;
  let startCol = 1;
  let numCols = sht.getLastColumn() - startCol +1;
  let rg = sht.getRange(startRow, startCol, numRows, numCols);
  let dt = rg.getValues();

  let crs = [];
  for (let i in dt) {
    if(dt[i][16]==(yr+'-2')){ // 若是下學期課程才會執行
      crs.push(String(dt[i][1]).concat(' ', dt[i][14]))
    }
  }
  let list_item = form.addListItem();
  list_item.setTitle('請選擇您的課程編號')
          .setChoiceValues(crs)
          .setRequired(true);

  var numValidation = FormApp.createTextValidation()
          .setHelpText('請輸入數字')
          .requireNumber()
          .build();

  // 第一點
  form.addPageBreakItem().setTitle('1. 經費執行情形(含自籌款)');
  form.addSectionHeaderItem().setTitle('原核定計畫金額');
  form.addTextItem().setTitle('人事費').setRequired(true).setValidation(numValidation);
  form.addTextItem().setTitle('業務費').setRequired(true).setValidation(numValidation);
  form.addTextItem().setTitle('設備費').setRequired(true).setValidation(numValidation);

  form.addSectionHeaderItem().setTitle('目前實支數');
  form.addTextItem().setTitle('人事費').setRequired(true).setValidation(numValidation);
  form.addTextItem().setTitle('業務費').setRequired(true).setValidation(numValidation);
  form.addTextItem().setTitle('設備費').setRequired(true).setValidation(numValidation);

  // 第二點
  form.addPageBreakItem().setTitle('2. 教學設備採購進度(填設備費之設備)')
                        .setHelpText('格式:品項*個數/金額\nEX: 伺服器*1/NT$20,000');
  form.addParagraphTextItem().setTitle('預計購買項目(含金額)').setRequired(true);
  form.addParagraphTextItem().setTitle('已完成招標/完成請購之項目(含金額)').setRequired(true);

  // 第三點
  form.addPageBreakItem().setTitle('3. 課程與模組結合使用情形')
  form.addCheckboxItem().setTitle('課程之採用模組')
                        .setChoiceValues(md)
                        .setRequired(true);
  form.addTextItem().setTitle('採用模組時數\n例如: A-1 (12小時)、C-2 (9小時)').setRequired(true);
  var img = UrlFetchApp.fetch('https://www.google.com/images/srpr/logo4w.png')
  form.addImageItem().setTitle('範例\n\
          (請將完成的"課程與模組結合使用情形"檔案上傳至\n\
          https://drive.google.com/drive/folders/12E-3wG8AzMo5u5YFkycrWI0WoCyWNZG4 )').setImage(img);

  // 第四點
  form.addPageBreakItem().setTitle('4. 預計業界或校外講師參與教學情形');
  form.addParagraphTextItem().setTitle('業界專家學者及其他跨領域教師實際參與計畫情形');
}
