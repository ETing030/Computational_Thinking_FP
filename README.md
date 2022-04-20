# Computational_Thinking_FP
運算思維之期末小組報告 :D

## 關於此 repo
以「如何幫助對轉系、雙主修和輔系感興趣的同學，可以釐清自己想要走的方向，並統整相關資訊（申請條件、所修課程和學生心得）以方便大家查詢」，這個問題作為初衷
而最終成果只著重於資料上的整理，以及 GUI 的設計，學生心得這一部分也尚未納入統整中  
程式相關說明將會在之後補上，估計會在 4/25 前完成  
目前僅附上最終 GUI 的成果，但仍有需改進的地方：
1. 申請條件中的部分欄位無法顯示所有內容（如先修課程），必須勾選「顯示較原始條件」才可以顯示出來
2. 課程中的欄位需再精簡，若改成以課表形式表示會更好

![image](https://user-images.githubusercontent.com/39528069/162632428-0a12c84e-4d67-491d-9a8d-a4fae7d8f731.png)

![image](https://user-images.githubusercontent.com/39528069/164186875-08993e65-8897-46b1-9c34-ebde098aa892.png)

整體來說，這次的成果並非有效地解決我們所提出的問題，算得上是一次失敗的經驗，但有磨練到程式能力 :)

---
## 關於報告題目（有點過度修飾）
有著轉系經驗的資管系學生、已自由選修外系多堂課的化工系學生，以及曾試圖雙主修的材料系的學生，在一堂跨領域的課程上相遇。系上的課程不足以滿足我們的「胃口」，而這種現象在學生中也不少見，但是，怎麼樣的課程，怎麼樣的方式，才能找到符合我們「口味的食物」呢？因此，我們題目訂為「如何幫助對轉系、雙主修和輔系感興趣的同學，可以釐清自己想要走的方向，並統整相關資訊以方便大家查詢」，試圖讓更多同學找到適合自己的課程。  


## 流程
1.~ 4. 為相關資料的收集與整理  
5. 為主要程式，說明針對這次報告題目，我們所採取的方式  
6. 為以 5. 形式為基礎所建立之使用者介面   

1. 以人工的方式（複製和打字），搜集申請條件，包含以下項目。除了最後一項，其餘皆先由 Google sheet 存取，再下載至電腦中
   - 109 學年輔系條件、109 學年雙主條件、 110 學年轉系條件
   - 109 學年上學期和 108 學年下學期課程（依星期幾分成多個工作表），即 `20201203課程.xlsx` 和 `20201206 108學年下學期課程.xlsx`
   - 各系輔系和雙主修之修課要求，即 `輔系課程.xlsx` 和 `雙主修課程.xlsx`
![image](https://user-images.githubusercontent.com/39528069/163972410-3d83f238-b615-4580-8880-d4df331fa063.png)  
▲ `雙主修課程.xlsx` 整理形式，工作表為各個系所的英文縮寫
   
2. 人工初步分類相關條件與修課要求
   - 將輔系和雙主修之原文分段，如「申請條件」、「申請時間」、「需繳交文件」等，即 `條件一覽 - 109學年輔系條件.csv`、`條件一覽 - 109學年雙主條件.csv`
   - 轉系則是分成「平轉」和「降轉」，即 `條件一覽 - 110學年轉系條件.csv`

![image](https://user-images.githubusercontent.com/39528069/163945333-9fc73f67-0598-41e4-8fbf-51670b34e485.png)
▲ 輔系
![image](https://user-images.githubusercontent.com/39528069/163944466-0be0b7b5-dbcd-4622-bd11-001a27fd9349.png)
▲ 轉系

3. 將初步整理所得到的那三個 csv 檔案，先合併成一個 xlsx 檔案，其中各類別各佔一個工作表，並對內容微幅調整。此合併微調之檔案即 `條件一覽(整理).xlsx`。微調內容如下，但檔案中仍保留部分微調前的版本
   - 輔系和雙主修：「申請條件」更改為「成績要求」、「先修課程」等更細微類別
   - 轉系：由於降轉要求都同於平轉，因此捨去「平轉」和「降轉」，並和輔系／雙主修一樣，細分段落
 
4. `運算思維2.rmd` ，用 R 統整資料，包含
   - 將 `條件一覽(整理).xlsx` 中分於不同段落的類別，各自獨立成 column，並以 Dataframe 的形式表示
   - 針對 `輔系課程.xlsx` 和 `雙主修課程.xlsx`，分別將課程資訊串接為 Vector，並以 Dataframe 的形式表示
   - 合併條件和課程之 Dataframe （輔系和雙主修）
   - 分別匯出檔案，即 `輔系.xlsx`、`雙主.xlsx` 和 `轉系.xlsx`  
   註：此 repo 中所附的 `條件一覽 - 輔系統整.csv`、`條件一覽 - 雙主統整.csv` 和 `條件一覽 - 轉系統整.csv` 為上述三個 xlsx 檔之 csv 檔
   註：由於 rmd 檔案中路徑設定沒更動，因此要執行時須設定讀取路徑
   
![image](https://user-images.githubusercontent.com/39528069/163963207-5205c7eb-668d-48b1-9f7f-235e34ae69fe.png)  
▲ `雙主修課程.xlsx` 之課程名稱與備註 


5. 對於「釐清想要走的方向」，我們在這邊並沒有太大著墨，反而著重於「統整相關資訊」這塊，而主要程式為 `project.py`，涵蓋以下內容
   - 確認學生心中是否有明確目標（想雙主修／輔系／轉系的系所），若沒有，則另外透過外部職涯測驗來辨別自身較適合那一種類型的工作類別，再以測驗結果作為選擇科系的依據
   - 確認學生想去之科系和辦法（雙主修／輔系／轉系／自由選修）
   - 顯示對應該科系該辦法之原始資料（資料為 `條件一覽 - 109學年輔系條件.csv`、`條件一覽 - 109學年雙主條件.csv`、`條件一覽 - 110學年轉系條件.csv` 中之內容）
   - 顯示對應該科系該辦法之分類後的條件（資料為 `條件一覽 - 輔系統整.csv`、`條件一覽 - 雙主統整.csv`、`條件一覽 - 轉系統整.csv` 中之內容）
   - 若為雙主修／輔系，則會顯示課程名稱和備註（資料為 `條件一覽 - 輔系統整.csv`、`條件一覽 - 雙主統整.csv` 中之內容）
![image](https://user-images.githubusercontent.com/39528069/164173761-9f200d78-29c0-4a99-b8a7-09fd6d48abf7.png)
▲ 以建築系轉系為例
![image](https://user-images.githubusercontent.com/39528069/164184891-b666b292-1960-45d5-9e06-22f096d86c2e.png)
▲ 以雙主修數學系為例


6. 設計 GUI 版面，`projectuitry.ui`、`projectuitry.py` 分別為 PyQt5 的 GUI 版面設計與該版面對應之 py 檔，並以 `project.py` 為參考，製作具有該程式中所提及之功能，即 `projectpyqt.py`
![image](https://user-images.githubusercontent.com/39528069/162632428-0a12c84e-4d67-491d-9a8d-a4fae7d8f731.png)
▲ 以材料系轉系為例
![image](https://user-images.githubusercontent.com/39528069/164186890-abcfe40c-aaf9-455d-ac7f-42b9acfd65c9.png)
▲ 以雙主修資工系為例

---
4/20 晚上 或　4/21 補足
- GUI排版, 功能說明(不確定必要與否)
- 可改善點, 應補足地方
- 心得
- :)






