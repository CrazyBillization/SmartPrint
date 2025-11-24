using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;

namespace InvoiceReorderTool
{
    public partial class MainWindow : Window
    {
        private string? _selectedFilePath;
        private int _totalPages;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnSelectFileClicked(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "PDF 檔案 (*.pdf)|*.pdf",
                Title = "選擇原始 PDF 檔案"
            };

            if (dialog.ShowDialog() == true)
            {
                _selectedFilePath = dialog.FileName;
                FilePathText.Text = $"已選擇：{_selectedFilePath}";

                try
                {
                    using var reader = new PdfReader(_selectedFilePath);
                    using var pdfDoc = new PdfDocument(reader);
                    _totalPages = pdfDoc.GetNumberOfPages();
                    PageInfoText.Text = $"偵測到共 {_totalPages} 頁，總共 {_totalPages * 4} 張發票";
                    StatusText.Text = "狀態：準備開始";
                    StartButton.IsEnabled = _totalPages > 0;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"無法讀取 PDF：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                    _totalPages = 0;
                    StartButton.IsEnabled = false;
                }
            }
        }

        private async void OnStartClicked(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedFilePath))
            {
                MessageBox.Show("請先選擇 PDF 檔案。", "提醒", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "PDF 檔案 (*.pdf)|*.pdf",
                Title = "選擇儲存位置",
                FileName = "重新排序.pdf"
            };

            if (saveDialog.ShowDialog() != true)
            {
                return;
            }

            Progress.Value = 0;
            StatusText.Text = "處理中，請稍候…";
            StartButton.IsEnabled = false;
            SelectFileButton.IsEnabled = false;

            try
            {
                await Task.Run(() => ReorderPdf(_selectedFilePath!, saveDialog.FileName, ReportProgress));
                StatusText.Text = "處理完成，請選擇儲存位置";
                MessageBox.Show("發票重新排序的 PDF 已建立完成。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"處理時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "狀態：處理失敗";
            }
            finally
            {
                StartButton.IsEnabled = _totalPages > 0;
                SelectFileButton.IsEnabled = true;
            }
        }

        private void ReportProgress(int pageNumber)
        {
            Dispatcher.Invoke(() =>
            {
                double percent = _totalPages == 0 ? 0 : (pageNumber * 100.0 / _totalPages);
                Progress.Value = percent;
                StatusText.Text = $"正在處理第 {pageNumber} / {_totalPages} 頁…";
            });
        }

        private void ReorderPdf(string srcPath, string destPath, Action<int> reportProgress)
        {
            using var reader = new PdfReader(srcPath);
            using var srcDoc = new PdfDocument(reader);
            using var writer = new PdfWriter(destPath);
            using var destDoc = new PdfDocument(writer);

            int P = srcDoc.GetNumberOfPages();
            PageSize pageSize = PageSize.A4;
            float pageWidth = pageSize.GetWidth();
            float pageHeight = pageSize.GetHeight();
            float regionHeight = pageHeight / 4f;

            // 逐頁建立新文件
            for (int page = 1; page <= P; page++)
            {
                destDoc.AddNewPage(pageSize);
                PdfPage targetPage = destDoc.GetPage(page);

                for (int pos = 1; pos <= 4; pos++)
                {
                    // 依規則計算新頁面上 pos 的發票編號：invoiceNumber = page + (pos - 1) * P
                    int invoiceNumber = page + (pos - 1) * P;

                    // 由發票編號反推原始頁碼與位置
                    int srcPage = (int)Math.Ceiling(invoiceNumber / 4.0);
                    int srcPos = ((invoiceNumber - 1) % 4) + 1;

                    PdfPage originalPage = srcDoc.GetPage(srcPage);
                    var xObject = originalPage.CopyAsFormXObject(destDoc);

                    // 計算來源區塊與目標區塊的座標
                    float srcY = pageHeight - srcPos * regionHeight; // 來源區塊左下角 Y 座標
                    float destY = pageHeight - pos * regionHeight;   // 目標區塊左下角 Y 座標

                    PdfCanvas canvas = new PdfCanvas(targetPage);
                    canvas.SaveState();

                    // 只在目標等份範圍內繪製
                    canvas.Rectangle(0, destY, pageWidth, regionHeight);
                    canvas.Clip();
                    canvas.EndPath();

                    // 將來源區塊平移到目標區塊
                    canvas.AddXObject(xObject, 0, destY - srcY);
                    canvas.RestoreState();
                }

                reportProgress(page);
            }
        }
    }
}
