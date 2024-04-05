function creatForm() {
 var formId = "15oEQ0YnovTNXeZaZLDMu2rTe21EjnW6NcU9M13K_dW8";
 var form = FormApp.openById(formId);
 Logger.log(form);


 deleteAllItems(form); 
 

  form.addListItem()
    .setTitle("模考類型")
    .setChoiceValues(simulateTestBasicInfo());


  form.addListItem()
    .setTitle("科目")
    .setChoiceValues(["數甲", "物理", "化學"]);


  var textValidation = FormApp.createTextValidation()
    .requireNumber()
    .requireNumberBetween(0,60)
    .setHelpText("請輸入0到60之間的數字")
    .build();

  form.addTextItem()
    .setTitle("級分")
    .setValidation(textValidation);

  form.addTextItem()
    .setTitle("錯題編號");

  form.addTextItem()
    .setTitle("備註");
}

function deleteAllItems (form) {
  var items = form.getItems();
  for( let i = 0; i < items.length; ++i ) {
    form.deleteItem(items[i]);
  }
}

function simulateTestBasicInfo() {
  return [
    "110-2-中",
    "110-1-北",
    "110-2-北",
    "110-1-全",
    "110-1-全-1",
    "110-2-全",
    "110-2-全-1",
    "110-3-全",
    "111-1-中",
    "111-3-北",
    "111-1-全",
    "111-2-全",
    "111-3-全"
  ];
}

