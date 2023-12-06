function createNewForm(){
  let new_form = FormApp.create("期末書面審查意見表test");
  let new_form_id = new_form.getId();
  let destinationFolderId = "14rwuyU6AOO0ed-0SA7nwOUHkn5D_1X1N"; //須根據情況修改
  moveFile(new_form_id, destinationFolderId);
  
  SetForm(new_form)
}

function moveFile(fileId, destinationFolderId) {
  let destinationFolder = DriveApp.getFolderById(destinationFolderId);
  DriveApp.getFileById(fileId).moveTo(destinationFolder);
}

function SetForm(form){
  var cbValidation = FormApp.createCheckboxValidation()
  .requireSelectAtLeast(1)
  .build();

  let des = '111年度課程期末成果報告：\n\
https://drive.google.com/drive/folders/197RBItVngoSUAyTbgK7JTGbl4_XBT1wO?usp=sharing\n\
\n\
說明如下：\n\
\n\
一、 考評重點：\n\
\n\
    1.課程設立\n\
        補助課程與模組間的關聯度與整合程度\n\
\n\
    2.課程開授\n\
­	 課程內容、授課師資、實驗/實作課程/相關活動\n\
\n\
    3.課程推動\n\
­	實施面：修習課程學生人數\n\
­	管理面：開課規劃\n\
\n\
    4.人力與經費使用\n\
­	投入之人力、經費支用情形、執行率\n\
\n\
\n\
二、 評等分數：特優(4分)、優(3分)、良(2分)、差(1分)';
  form.setDescription(des);
  form.setCollectEmail(true);

  let ls = ['洪士灝 委員', '馬席彬 委員', '黃世旭 委員', '吳文慶 委員'];
  let list_item = form.addListItem();
  list_item.setTitle('請選擇您的身分')
          .setChoiceValues(ls)
          .setRequired(true);
  
  let sht = SpreadsheetApp.getActiveSheet();
  let startRow = 4;
  let numRows = sht.getLastRow() - startRow +1;
  let startCol = 2;
  let numCols = sht.getLastColumn() - startCol +1;
  let rg = sht.getRange(startRow, startCol, numRows, numCols);
  let dt = rg.getValues();

  let cb_ls = ['建議鼓勵學生多參與相關競賽', '設備採購執行率偏低', '業務費執行率偏低', '建議提早規畫期末專題與討論時間', '建議列舉及詳細說明學生專題作品', '修課人數低於預期', '報告內容說明不夠詳盡', '未明確說明實際教材數量及進度百分比', '無(於下一題說明)'];

  for (let i in dt) {
    //dt[i] is a row
    add_pages(form, dt[i])
    for (let j = 1; j <= 4; j++){ // 增加審查重點， 因為有四項，所以j<=4
      add_list(form, dt[i], j)
    }

    let cb = form.addCheckboxItem();
    cb.setTitle(dt[i][0]+' 綜合審查意見(可複選或在下欄中填寫補充意見)')
        .setChoiceValues(cb_ls)
        .setValidation(cbValidation)
        .setRequired(true);

    let dt_qus = form.addParagraphTextItem();
    dt_qus.setTitle(dt[i][0]+' 審查意見補充說明 (含以上文字至少撰寫100字)')
          .setRequired(true);

    let mc_ls = [4,3,2,1]
    let mc = form.addMultipleChoiceItem();
    mc.setTitle(dt[i][0]+' 綜合評分')
      .setHelpText('特優(4分)、優(3分)、良(2分)、差(1分)\n若分數有小數，可選擇其他欄位填寫')
      .setChoiceValues(mc_ls)
      .showOtherOption(true)
      .setRequired(true);
  }
  
}

function add_pages(form, row){
  let pg = form.addPageBreakItem().setTitle('智慧終端裝置晶片系統與應用聯盟');
  pg.setHelpText('課程名稱: '+ row[6] 
  +'\n計畫編號: '+ row[0] 
  +'\n使用重點模組: '+ row[13] 
  +'\n課程教師: '+ row[5] 
  +'\n學校/系所(服務單位): '+row[1]+'/'+row[4])
}

function add_list(form, row, sel){
  let lst = form.addListItem();
  if(sel==1){
    lst.setTitle(row[0]+' 審查重點- 第一項：課程設立')
    lst.setHelpText('說明：補助課程與模組間的關聯度與整合程度')
  }
  else if(sel==2){
    lst.setTitle(row[0]+' 審查重點- 第二項：課程開授')
    lst.setHelpText('說明：課程內容、授課師資、實驗/實作課程/相關活動')
  }
  else if(sel==3){
    lst.setTitle(row[0]+' 審查重點- 第三項：課程推動')
    lst.setHelpText('說明：1.實施面：修習課程學生人數。  2.-管理面：開課規劃')
  }
  else{
    lst.setTitle(row[0]+' 審查重點- 第四項：人力與經費使用')
    lst.setHelpText('說明：投入之人力、經費支用情形、執行率')
  }

  let ls = ['特優', '優', '良', '差'];
  lst.setChoiceValues(ls)
      .setRequired(true);
}

// function anotherSetForm(){
//   let form = FormApp.openById('1nQIUsIas8vWDxLM6WVec3CtPbRxFUFobHiXrhREN6cI');

//   let des = '111年度課程期末成果報告：\n\
// https://drive.google.com/drive/folders/197RBItVngoSUAyTbgK7JTGbl4_XBT1wO?usp=sharing\n\
// \n\
// 說明如下：\n\
// \n\
// 一、 考評重點：\n\
// \n\
//     1.課程設立\n\
//         補助課程與模組間的關聯度與整合程度\n\
// \n\
//     2.課程開授\n\
// ­	 課程內容、授課師資、實驗/實作課程/相關活動\n\
// \n\
//     3.課程推動\n\
// ­	實施面：修習課程學生人數\n\
// ­	管理面：開課規劃\n\
// \n\
//     4.人力與經費使用\n\
// ­	投入之人力、經費支用情形、執行率\n\
// \n\
// \n\
// 二、 評等分數：特優(4分)、優(3分)、良(2分)、差(1分)';
//   form.setDescription(des);
//   form.setCollectEmail(true);

//   let ls = ['洪士灝 委員', '馬席彬 委員', '黃世旭 委員', '吳文慶 委員'];
//   let list_item = form.addListItem();
//   list_item.setTitle('請選擇您的身分')
//           .setChoiceValues(ls)
  
//   let sht = SpreadsheetApp.getActiveSheet();
//   let startRow = 4;
//   let numRows = sht.getLastRow() - startRow +1;
//   let startCol = 2;
//   let numCols = sht.getLastColumn() - startCol +1;
//   let rg = sht.getRange(startRow, startCol, numRows-1, numCols);
//   let dt = rg.getValues();

//   for (let i in dt) {
//     //dt[i] is a row
//     add_pages(form, dt[i])
//     for (let j = 1; j <= 4; j++){ // 增加審查重點， 因為有四項，所以j<=4
//       add_list(form, dt[i], j)
//     }
//   }
  
// }
