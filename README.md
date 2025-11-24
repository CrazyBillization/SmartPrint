# 發票 PDF 重新排序工具

以 WPF/.NET 6 建立的 Windows 桌面應用程式，可依指定規則重新排列每頁四等份的發票 PDF，輸出新的檔案。

## 專案結構
- `InvoiceReorderTool/InvoiceReorderTool.csproj`：WPF 專案與 NuGet 套件參考（iText7）。
- `InvoiceReorderTool/MainWindow.xaml`：主要 UI 配置與繁體中文介面文字。
- `InvoiceReorderTool/MainWindow.xaml.cs`：檔案讀取、重新排序與 PDF 輸出邏輯。
- `InvoiceReorderTool/App.xaml`、`App.xaml.cs`：應用程式進入點。

## 主要流程
1. 以 `OpenFileDialog` 選擇原始 PDF，偵測頁數 P 與總發票數 4P。
2. 依規則 `invoiceNumber = page + (pos - 1) * P` 重新定位來源頁與區塊：
   - `srcPage = ceil(invoiceNumber / 4)`
   - `srcPos = ((invoiceNumber - 1) mod 4) + 1`
3. 每頁垂直切成四等份，使用 iText7 `PdfCanvas` 將來源區塊平移到新頁對應等份。
4. 以 `SaveFileDialog` 輸出新 PDF。

## 建置與單檔發佈
1. 安裝 .NET 6 SDK。
2. 還原套件與建置：
   ```bash
   dotnet restore InvoiceReorderTool/InvoiceReorderTool.csproj
   dotnet build InvoiceReorderTool/InvoiceReorderTool.csproj -c Release
   ```
3. 產生單一可執行檔（win-x64 範例）：
   ```bash
   dotnet publish InvoiceReorderTool/InvoiceReorderTool.csproj -c Release -r win-x64 -p:PublishSingleFile=true --self-contained true
   ```
   輸出會出現在 `bin/Release/net6.0-windows/win-x64/publish/`。
