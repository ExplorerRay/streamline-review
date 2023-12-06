function createNewForm2(){
  let yr = '111'; //年度記得修改
  // 下學期
  let new_form_2 = FormApp.create(yr + "-2 課程書面審查意見表test"); 
  let new_form_2_id = new_form_2.getId();

  let destinationFolderId = "1r_cIrF6-v94fRg2Jj5QfKlPHiJEorrs8"; //須根據情況修改
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
  let des = yr+'-2課程成果報告：\n\
  https://drive.google.com/drive/folders/1LLBxWHjNT9wq6ryua6LDEiFmeY4lxlyR\n\
\n\
說明如下：\n\
\n\
一、 考評重點：\n\
\n\
    1.課程的備課狀況\n\
\n\
    2.與推廣模組整合的規劃\n\
\n\
    3.教學設備採購情況\n\
\n\
\n\
二、 評等分數：優(4分)、佳(3分)、可(2分)、差(1分)';
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

  let cb_ls = ['建議鼓勵學生多參與相關競賽', '設備採購須加快進行', '經費執行率偏低', '建議提早規畫期末專題與討論時間', '建議列舉及詳細說明學生專題作品', '須留意學生對於課程的吸收', '報告內容說明不夠詳盡', '未明確說明實際教材數量及進度百分比', '無(於下一題說明)'];

  for (let i in dt) {
    //dt[i] is a row
    if(dt[i][8]==(yr+'-2')){ // 若是下學期課程才會執行
      add_pages2(form, dt[i])
      for (let j = 1; j <= 3; j++){ // 增加審查重點， 因為有3項，所以j<=3
        add_list2(form, dt[i], j)
      }

      let cb = form.addCheckboxItem();
      cb.setTitle(dt[i][0]+' 綜合審查意見(可複選或在下欄中填寫補充意見)')
          .setChoiceValues(cb_ls)
          .setValidation(cbValidation)
          .setRequired(true);

      let dt_qus = form.addParagraphTextItem();
      dt_qus.setTitle(dt[i][0]+' 審查意見補充說明')
            .setRequired(true);

      let mc_ls = [4,3,2,1]
      let mc = form.addMultipleChoiceItem();
      mc.setTitle(dt[i][0]+' 綜合評分')
        .setHelpText('優(4分)、佳(3分)、可(2分)、差(1分)\n若分數有小數，可選擇其他欄位填寫')
        .setChoiceValues(mc_ls)
        .showOtherOption(true)
        .setRequired(true);
    }
  }
  
}

function add_pages2(form, row){
  let pg = form.addPageBreakItem().setTitle('智慧終端裝置晶片系統與應用聯盟');
  pg.setHelpText('課程名稱: '+ row[6] 
  +'\n計畫編號: '+ row[0] 
  +'\n使用重點模組: '+ row[13] 
  +'\n課程教師: '+ row[5] 
  +'\n學校/系所(服務單位): '+row[1]+'/'+row[4])
}

function add_list2(form, row, sel){
  let lst = form.addListItem();
  if(sel==1){
    lst.setTitle(row[0]+' 審查重點- 第一項：課程的備課狀況')
  }
  else if(sel==2){
    lst.setTitle(row[0]+' 審查重點- 第二項：與推廣模組整合的規劃')
  }
  else{
    lst.setTitle(row[0]+' 審查重點- 第三項：教學設備採購情況')
  }

  let ls = ['優', '佳', '可', '差'];
  lst.setChoiceValues(ls)
      //.setRequired(true);
}
  
  