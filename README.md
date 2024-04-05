# 專案目標
- 可以透過填表單自動將表單資料填入`google sheet`



# 檔案架構
- `createForm.gs`
- `form.gs`
- `sheet.gs`

## `createForm.gs`
- 製作`google sheet`

### **使用時需要修改的地方**
- `function simulateTestBasicInfo() => return` 應該跟 `form.gs => function tableNameWithUrl() => return` 有關係
- `function createForm() => formId` 請修改成自己`google form`的`ID`

### `createForm()`
- 動態新建表單

### `simulateTestBasicInfo()`
- 這裡應該會有檔案夾的名稱

## `form.gs`
- 清加在`google sheet`裡面的程式碼編輯器
###  **使用時需要修改的地方**
- `function onFormSubmit(e) => url ` 修改成自己表單的`url`
- `function tableNameWithUrl() => return`  如果有需要的話請自行增加


### `tableNamWithUrl()`
- 回傳檔案夾名稱跟連結 

### `findOrCreateTableName(tableName, sheet)`
- 找到`A`列表單起始位置，並且找到第一個空白的位置放入`tableName`

### `columnLetterToNumber(columnLeeter)`
-  將`列`轉換成數字
- `A` $\to$ `1`, `E` $\to$ `5`, `AB` $\to$ `28`

### `onFormSubmit(e)`
-   在`formSubmit`  的時候更新`sheet`裡面的單元格


## `sheet.gs`
- 請在`google sheet`裡面的程式碼編輯器
- 在選擇下拉式選單的時候可以動態將表單值抓到選單後面

### `checkIsChooseType(range)`
-  確認修改的是不是下拉式選單


### `columnLetterToNumber(columnLetter)`
-  同上

### `getStartPositoin(sheet)`
- 拿到 `數`, `物`, `化` 科目的欄位起始位置(儲存在`Y5`, `Y7`, `Y9`內)

### `onEdit(e)`
- 在修改的時候確認修該倒的是不是下拉式選單

