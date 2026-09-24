using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Windows.Storage;
using Windows.Storage.Streams;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using WinPdfDocument = Windows.Data.Pdf.PdfDocument;
using WinPdfPageRenderOptions = Windows.Data.Pdf.PdfPageRenderOptions;
using WpfMessageBox = System.Windows.MessageBox;
using WpfOpenFileDialog = Microsoft.Win32.OpenFileDialog;
using WpfPoint = System.Windows.Point;
using WinForms = System.Windows.Forms;

namespace TachFilePdf;

public partial class MainWindow : Window
{
    private const double RenderedPreviewPageWidth = 900;
    private const double PreviewPageHorizontalChrome = 54;
    private const double PreviewZoomMin = 0.35;
    private const double PreviewZoomMax = 2.5;
    private const double PreviewZoomStep = 0.1;

    private readonly ObservableCollection<PageRangeItem> _ranges = [];
    private readonly ObservableCollection<PdfPagePreviewItem> _pagePreviews = [];
    private readonly ObservableCollection<MergePdfItem> _mergeFiles = [];
    private string? _inputFilePath;
    private string? _outputFolderPath;
    private int _pageCount;
    private int _currentPreviewPage = 1;
    private double _previewZoom = 1;
    private bool _isFitWidthZoom;
    private bool _isSelectingCrop;
    private WpfPoint _cropStartPoint;
    private CropArea? _cropArea;

    public MainWindow()
    {
        InitializeComponent();

        _ranges.Add(new PageRangeItem { Index = 1, FromPage = "1", ToPage = "1" });
        RangesItemsControl.ItemsSource = _ranges;
        PdfPagesItemsControl.ItemsSource = _pagePreviews;
        MergeFilesListBox.ItemsSource = _mergeFiles;
        ApplyCropScopeChanged(this, new RoutedEventArgs());
        ApplyPreviewZoom();
        UpdatePreviewInteractionLayers();
        UpdatePreviewPageHeader();
        UpdatePrimaryActionButton();
    }

    private async void ChoosePdfButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new WpfOpenFileDialog
        {
            Title = "Chọn tệp PDF",
            Filter = "PDF files (*.pdf)|*.pdf",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        _inputFilePath = dialog.FileName;
        InputFileTextBox.Text = _inputFilePath;

        if (string.IsNullOrWhiteSpace(_outputFolderPath))
        {
            _outputFolderPath = Path.Combine(
                Path.GetDirectoryName(_inputFilePath)!,
                $"{Path.GetFileNameWithoutExtension(_inputFilePath)}_tach");
            OutputFolderTextBox.Text = _outputFolderPath;
        }

        try
        {
            using var document = PdfReader.Open(_inputFilePath, PdfDocumentOpenMode.Import);
            _pageCount = document.PageCount;
            _currentPreviewPage = 1;
            PreviewPageTextBox.Text = "1";
            CropPageTextBox.Text = "1";
            ClearCropSelectionVisual();
            _pagePreviews.Clear();
            PreviewPlaceholderPanel.Visibility = Visibility.Collapsed;
            UpdatePreviewPageHeader();
            UpdatePreviewInteractionLayers();

            if (_ranges.Count == 1)
            {
                _ranges[0].FromPage = "1";
                _ranges[0].ToPage = _pageCount.ToString(CultureInfo.InvariantCulture);
            }

            DocumentInfoTextBlock.Text = $"{_pageCount} trang";
            StatusTextBlock.Text = $"Đang render {_pageCount} trang PDF...";
            await LoadPdfPreviewPagesAsync(_inputFilePath);
            FitPreviewToWidth(allowZoomIn: false);
            ScrollToPreviewPage(_currentPreviewPage);
            StatusTextBlock.Text = $"Đã chọn PDF có {_pageCount} trang.";
        }
        catch (Exception ex)
        {
            _pageCount = 0;
            _pagePreviews.Clear();
            DocumentInfoTextBlock.Text = "Không đọc được PDF";
            StatusTextBlock.Text = "Không đọc được tệp PDF.";
            UpdatePreviewPageHeader();
            UpdatePreviewInteractionLayers();
            WpfMessageBox.Show(this, ex.Message, "Lỗi đọc PDF", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ChooseOutputFolderButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new WinForms.FolderBrowserDialog
        {
            Description = "Chọn thư mục xuất file PDF",
            UseDescriptionForTitle = true
        };

        if (dialog.ShowDialog() != WinForms.DialogResult.OK)
        {
            return;
        }

        _outputFolderPath = dialog.SelectedPath;
        OutputFolderTextBox.Text = _outputFolderPath;
    }

    private void ModeRadio_Checked(object sender, RoutedEventArgs e)
    {
        if (CustomRangesPanel is null || FixedPanel is null || CustomModeRadio is null)
        {
            return;
        }

        CustomRangesPanel.Visibility = CustomModeRadio.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        FixedPanel.Visibility = CustomModeRadio.IsChecked == true ? Visibility.Collapsed : Visibility.Visible;
        UpdatePreviewInteractionLayers();
    }

    private void AddRangeButton_Click(object sender, RoutedEventArgs e)
    {
        _ranges.Add(new PageRangeItem
        {
            Index = _ranges.Count + 1,
            FromPage = "1",
            ToPage = _pageCount > 0 ? _pageCount.ToString(CultureInfo.InvariantCulture) : "1"
        });
        RefreshRangeLabels();
    }

    private void RemoveRangeButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: PageRangeItem item })
        {
            return;
        }

        if (_ranges.Count == 1)
        {
            item.FromPage = "1";
            item.ToPage = _pageCount > 0 ? _pageCount.ToString(CultureInfo.InvariantCulture) : "1";
            return;
        }

        _ranges.Remove(item);
        RefreshRangeLabels();
    }

    private void PreviousPageButton_Click(object sender, RoutedEventArgs e)
    {
        if (_pageCount <= 0 || _currentPreviewPage <= 1)
        {
            return;
        }

        _currentPreviewPage--;
        SyncCurrentPageFields();
        ScrollToPreviewPage(_currentPreviewPage);
    }

    private void NextPageButton_Click(object sender, RoutedEventArgs e)
    {
        if (_pageCount <= 0 || _currentPreviewPage >= _pageCount)
        {
            return;
        }

        _currentPreviewPage++;
        SyncCurrentPageFields();
        ScrollToPreviewPage(_currentPreviewPage);
    }

    private void ZoomOutButton_Click(object sender, RoutedEventArgs e)
    {
        _isFitWidthZoom = false;
        SetPreviewZoom(_previewZoom - PreviewZoomStep);
    }

    private void ZoomInButton_Click(object sender, RoutedEventArgs e)
    {
        _isFitWidthZoom = false;
        SetPreviewZoom(_previewZoom + PreviewZoomStep);
    }

    private void FitWidthButton_Click(object sender, RoutedEventArgs e)
    {
        FitPreviewToWidth(allowZoomIn: true);
    }

    private async void PreviewPageTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        await NavigateToPreviewPageFromTextBoxAsync();
    }

    private async void PreviewPageTextBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
        {
            return;
        }

        await NavigateToPreviewPageFromTextBoxAsync();
        e.Handled = true;
    }

    private async Task NavigateToPreviewPageFromTextBoxAsync()
    {
        if (_pageCount <= 0)
        {
            PreviewPageTextBox.Text = "1";
            return;
        }

        if (!int.TryParse(PreviewPageTextBox.Text, out var page))
        {
            page = _currentPreviewPage;
        }

        _currentPreviewPage = Math.Clamp(page, 1, _pageCount);
        SyncCurrentPageFields();
        ScrollToPreviewPage(_currentPreviewPage);
    }

    private void SetSplitPointButton_Click(object sender, RoutedEventArgs e)
    {
        SetCurrentPageAsSplitPoint();
    }

    private void SetCurrentPageAsSplitPoint()
    {
        if (_pageCount <= 0 || MainActionTabControl.SelectedIndex != 0 || CustomModeRadio.IsChecked != true)
        {
            return;
        }

        var startPage = ReadCurrentPreviewPage();
        if (startPage <= 1)
        {
            StatusTextBlock.Text = "Trang 1 đã là trang đầu tiên của PDF.";
            return;
        }

        var endPage = startPage - 1;
        ApplyCustomSplitEndPage(endPage);
        StatusTextBlock.Text = $"Đã đặt điểm tách trước trang {startPage}. File mới bắt đầu từ trang {startPage}.";
    }

    private void ApplyCropScopeChanged(object sender, RoutedEventArgs e)
    {
        if (CropPageTextBox is null || ApplyCropAllPagesCheckBox is null)
        {
            return;
        }

        CropPageTextBox.IsEnabled = ApplyCropAllPagesCheckBox.IsChecked != true;
    }

    private void MainActionTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (e.Source == MainActionTabControl)
        {
            UpdatePreviewInteractionLayers();
            UpdatePrimaryActionButton();
        }
    }

    private void CropOverlayCanvas_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_pageCount <= 0)
        {
            return;
        }

        _isSelectingCrop = true;
        _cropStartPoint = e.GetPosition(CropOverlayCanvas);
        CropSelectionRectangle.Visibility = Visibility.Visible;
        Canvas.SetLeft(CropSelectionRectangle, _cropStartPoint.X);
        Canvas.SetTop(CropSelectionRectangle, _cropStartPoint.Y);
        CropSelectionRectangle.Width = 0;
        CropSelectionRectangle.Height = 0;
        CropOverlayCanvas.CaptureMouse();
        e.Handled = true;
    }

    private void CropOverlayCanvas_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isSelectingCrop)
        {
            return;
        }

        var currentPoint = e.GetPosition(CropOverlayCanvas);
        var left = Math.Min(_cropStartPoint.X, currentPoint.X);
        var top = Math.Min(_cropStartPoint.Y, currentPoint.Y);
        var width = Math.Abs(currentPoint.X - _cropStartPoint.X);
        var height = Math.Abs(currentPoint.Y - _cropStartPoint.Y);

        Canvas.SetLeft(CropSelectionRectangle, left);
        Canvas.SetTop(CropSelectionRectangle, top);
        CropSelectionRectangle.Width = width;
        CropSelectionRectangle.Height = height;
    }

    private void CropOverlayCanvas_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isSelectingCrop)
        {
            return;
        }

        _isSelectingCrop = false;
        CropOverlayCanvas.ReleaseMouseCapture();

        if (CropSelectionRectangle.Width < 12 || CropSelectionRectangle.Height < 12)
        {
            ClearCropSelectionVisual();
            e.Handled = true;
            return;
        }

        var canvasWidth = Math.Max(CropOverlayCanvas.ActualWidth, 1);
        var canvasHeight = Math.Max(CropOverlayCanvas.ActualHeight, 1);
        var left = Math.Clamp(Canvas.GetLeft(CropSelectionRectangle), 0, canvasWidth);
        var top = Math.Clamp(Canvas.GetTop(CropSelectionRectangle), 0, canvasHeight);
        var right = Math.Clamp(left + CropSelectionRectangle.Width, 0, canvasWidth);
        var bottom = Math.Clamp(top + CropSelectionRectangle.Height, 0, canvasHeight);

        _cropArea = new CropArea(
            left / canvasWidth,
            top / canvasHeight,
            right / canvasWidth,
            bottom / canvasHeight);

        CropSelectionInfoTextBlock.Text = $"Đã chọn vùng giữ lại: {Math.Round(_cropArea.Value.WidthPercent)}% x {Math.Round(_cropArea.Value.HeightPercent)}%";
        StatusTextBlock.Text = "Đã chọn vùng crop. Có thể lưu PDF crop hoặc dùng vùng này khi tách file.";
        e.Handled = true;
    }

    private void ClearCropSelectionButton_Click(object sender, RoutedEventArgs e)
    {
        ClearCropSelectionVisual();
        StatusTextBlock.Text = "Đã xóa vùng crop.";
    }

    private void PrimaryActionButton_Click(object sender, RoutedEventArgs e)
    {
        switch (MainActionTabControl.SelectedIndex)
        {
            case 1:
                CropPdfButton_Click(sender, e);
                break;
            case 2:
                MergePdfButton_Click(sender, e);
                break;
            default:
                SplitPdfButton_Click(sender, e);
                break;
        }
    }

    private void SplitPdfButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateInputAndOutput())
        {
            return;
        }

        try
        {
            Directory.CreateDirectory(_outputFolderPath!);

            using var inputDocument = PdfReader.Open(_inputFilePath!, PdfDocumentOpenMode.Import);
            var ranges = CustomModeRadio.IsChecked == true
                ? BuildCustomRanges(inputDocument.PageCount)
                : BuildFixedRanges(inputDocument.PageCount);
            var cropOptions = UseCropWhenSplitCheckBox.IsChecked == true
                ? ReadCropOptions(inputDocument.PageCount)
                : null;

            var baseName = SanitizeFileName(Path.GetFileNameWithoutExtension(_inputFilePath!));
            var outputCount = 0;
            foreach (var range in ranges)
            {
                outputCount++;
                var outputFile = Path.Combine(
                    _outputFolderPath!,
                    $"{baseName}_trang_{range.From:000}_den_{range.To:000}.pdf");
                SaveRange(inputDocument, range.From, range.To, outputFile, cropOptions);
            }

            StatusTextBlock.Text = $"Hoàn tất: đã tách {outputCount} file vào {_outputFolderPath}.";
            WpfMessageBox.Show(this, $"Đã tách xong {outputCount} file PDF.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Tách PDF thất bại.";
            WpfMessageBox.Show(this, ex.Message, "Lỗi tách PDF", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CropPdfButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateInputAndOutput())
        {
            return;
        }

        try
        {
            Directory.CreateDirectory(_outputFolderPath!);

            using var inputDocument = PdfReader.Open(_inputFilePath!, PdfDocumentOpenMode.Import);
            var cropOptions = ReadCropOptions(inputDocument.PageCount);
            var baseName = SanitizeFileName(Path.GetFileNameWithoutExtension(_inputFilePath!));
            var outputFile = GetUniqueOutputPath(Path.Combine(_outputFolderPath!, $"{baseName}_crop.pdf"));
            SaveCroppedPdf(inputDocument, outputFile, cropOptions);

            StatusTextBlock.Text = $"Đã lưu file crop: {outputFile}";
            WpfMessageBox.Show(this, "Đã lưu PDF đã crop.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Crop PDF thất bại.";
            WpfMessageBox.Show(this, ex.Message, "Lỗi crop PDF", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void AddMergeFilesButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new WpfOpenFileDialog
        {
            Title = "Chọn các tệp PDF cần gộp",
            Filter = "PDF files (*.pdf)|*.pdf",
            CheckFileExists = true,
            Multiselect = true
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var addedCount = 0;
        foreach (var filePath in dialog.FileNames)
        {
            if (_mergeFiles.Any(item => string.Equals(item.FilePath, filePath, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            using var document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
            var fileInfo = new FileInfo(filePath);
            _mergeFiles.Add(new MergePdfItem
            {
                Index = _mergeFiles.Count + 1,
                FilePath = filePath,
                PageCount = document.PageCount,
                FileSizeBytes = fileInfo.Length
            });
            addedCount++;
        }

        if (string.IsNullOrWhiteSpace(_outputFolderPath) && _mergeFiles.Count > 0)
        {
            _outputFolderPath = Path.GetDirectoryName(_mergeFiles[0].FilePath);
            OutputFolderTextBox.Text = _outputFolderPath;
        }

        if (MergeFilesListBox.SelectedIndex < 0 && _mergeFiles.Count > 0)
        {
            MergeFilesListBox.SelectedIndex = 0;
        }

        RefreshMergeFileIndexes();
        StatusTextBlock.Text = addedCount > 0
            ? $"Đã thêm {addedCount} file PDF vào danh sách gộp."
            : "Các file đã có trong danh sách gộp.";
    }

    private void ClearMergeFilesButton_Click(object sender, RoutedEventArgs e)
    {
        _mergeFiles.Clear();
        StatusTextBlock.Text = "Đã xóa danh sách file gộp.";
    }

    private void RemoveMergeFileButton_Click(object sender, RoutedEventArgs e)
    {
        if (MergeFilesListBox.SelectedItem is not MergePdfItem item)
        {
            StatusTextBlock.Text = "Vui lòng chọn file cần xóa khỏi danh sách gộp.";
            return;
        }

        var selectedIndex = MergeFilesListBox.SelectedIndex;
        _mergeFiles.Remove(item);
        RefreshMergeFileIndexes();
        if (_mergeFiles.Count > 0)
        {
            MergeFilesListBox.SelectedIndex = Math.Min(selectedIndex, _mergeFiles.Count - 1);
        }

        StatusTextBlock.Text = "Đã xóa file khỏi danh sách gộp.";
    }

    private void MoveMergeFileUpButton_Click(object sender, RoutedEventArgs e)
    {
        MoveSelectedMergeFile(-1);
    }

    private void MoveMergeFileDownButton_Click(object sender, RoutedEventArgs e)
    {
        MoveSelectedMergeFile(1);
    }

    private void MergeFilesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (MainActionTabControl.SelectedIndex != 2 || MergeFilesListBox.SelectedItem is not MergePdfItem item)
        {
            return;
        }

        StatusTextBlock.Text = $"Đã chọn trong danh sách gộp: {item.FileName}";
    }

    private void MergePdfButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateMergeInput())
        {
            return;
        }

        try
        {
            Directory.CreateDirectory(_outputFolderPath!);

            var outputFile = GetUniqueOutputPath(Path.Combine(_outputFolderPath!, BuildMergeOutputFileName()));
            using var outputDocument = new PdfDocument();
            foreach (var item in _mergeFiles)
            {
                using var inputDocument = PdfReader.Open(item.FilePath, PdfDocumentOpenMode.Import);
                for (var pageIndex = 0; pageIndex < inputDocument.PageCount; pageIndex++)
                {
                    outputDocument.AddPage(inputDocument.Pages[pageIndex]);
                }
            }

            outputDocument.Save(outputFile);
            StatusTextBlock.Text = $"Hoàn tất: đã gộp {_mergeFiles.Count} file vào {outputFile}.";
            WpfMessageBox.Show(this, "Đã gộp xong các file PDF.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Gộp PDF thất bại.";
            WpfMessageBox.Show(this, ex.Message, "Lỗi gộp PDF", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private bool ValidateMergeInput()
    {
        if (_mergeFiles.Count < 2)
        {
            WpfMessageBox.Show(this, "Vui lòng chọn ít nhất 2 file PDF để gộp.", "Thiếu file PDF", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        foreach (var item in _mergeFiles)
        {
            if (!File.Exists(item.FilePath))
            {
                WpfMessageBox.Show(this, $"Không tìm thấy file: {item.FilePath}", "Thiếu file PDF", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
        }

        if (string.IsNullOrWhiteSpace(_outputFolderPath))
        {
            WpfMessageBox.Show(this, "Vui lòng chọn thư mục xuất.", "Thiếu thư mục xuất", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    private string BuildMergeOutputFileName()
    {
        var rawName = MergeOutputNameTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(rawName))
        {
            rawName = "PDF_da_gop.pdf";
        }

        var extension = Path.GetExtension(rawName);
        var nameWithoutExtension = string.IsNullOrWhiteSpace(extension)
            ? rawName
            : Path.GetFileNameWithoutExtension(rawName);
        var safeName = SanitizeFileName(nameWithoutExtension);
        return $"{safeName}.pdf";
    }

    private void MoveSelectedMergeFile(int direction)
    {
        var oldIndex = MergeFilesListBox.SelectedIndex;
        if (oldIndex < 0)
        {
            StatusTextBlock.Text = "Vui lòng chọn file cần đổi thứ tự.";
            return;
        }

        var newIndex = oldIndex + direction;
        if (newIndex < 0 || newIndex >= _mergeFiles.Count)
        {
            return;
        }

        _mergeFiles.Move(oldIndex, newIndex);
        RefreshMergeFileIndexes();
        MergeFilesListBox.SelectedIndex = newIndex;
        StatusTextBlock.Text = "Đã cập nhật thứ tự gộp PDF.";
    }

    private void RefreshMergeFileIndexes()
    {
        for (var i = 0; i < _mergeFiles.Count; i++)
        {
            _mergeFiles[i].Index = i + 1;
        }
    }

    private void UpdatePrimaryActionButton()
    {
        if (PrimaryActionButton is null || MainActionTabControl is null)
        {
            return;
        }

        PrimaryActionButton.Content = MainActionTabControl.SelectedIndex switch
        {
            1 => "Crop PDF",
            2 => "Gộp PDF",
            _ => "Tách PDF"
        };
    }

    private async Task LoadPdfPreviewPagesAsync(string filePath)
    {
        _pagePreviews.Clear();

        var storageFile = await StorageFile.GetFileFromPathAsync(filePath);
        var pdfDocument = await WinPdfDocument.LoadFromFileAsync(storageFile);
        for (uint pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            using var page = pdfDocument.GetPage(pageIndex);
            var image = await RenderPdfPageAsync(page);
            _pagePreviews.Add(new PdfPagePreviewItem((int)pageIndex + 1, image));

            if (pageIndex % 5 == 0 || pageIndex + 1 == pdfDocument.PageCount)
            {
                StatusTextBlock.Text = $"Đang render trang {pageIndex + 1}/{pdfDocument.PageCount}...";
                await Task.Yield();
            }
        }
    }

    private static async Task<BitmapImage> RenderPdfPageAsync(Windows.Data.Pdf.PdfPage page)
    {
        var scale = RenderedPreviewPageWidth / Math.Max(page.Size.Width, 1);
        var options = new WinPdfPageRenderOptions
        {
            DestinationWidth = (uint)Math.Round(page.Size.Width * scale),
            DestinationHeight = (uint)Math.Round(page.Size.Height * scale)
        };

        using var randomStream = new InMemoryRandomAccessStream();
        await page.RenderToStreamAsync(randomStream, options);
        randomStream.Seek(0);

        using var reader = new DataReader(randomStream.GetInputStreamAt(0));
        var loaded = await reader.LoadAsync((uint)randomStream.Size);
        var bytes = new byte[loaded];
        reader.ReadBytes(bytes);

        var bitmap = new BitmapImage();
        using var memoryStream = new MemoryStream(bytes);
        bitmap.BeginInit();
        bitmap.CacheOption = BitmapCacheOption.OnLoad;
        bitmap.StreamSource = memoryStream;
        bitmap.EndInit();
        bitmap.Freeze();
        return bitmap;
    }

    private void PreviewPageCard_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement { DataContext: PdfPagePreviewItem pageItem })
        {
            return;
        }

        _currentPreviewPage = pageItem.PageNumber;
        SyncCurrentPageFields();

        if (e.ClickCount >= 2)
        {
            SetCurrentPageAsSplitPoint();
            e.Handled = true;
        }
    }

    private void PdfPagesScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        if (_pageCount <= 0 || PreviewPageTextBox.IsKeyboardFocusWithin)
        {
            return;
        }

        UpdateCurrentPageFromVisiblePreview();
    }

    private void PdfPagesScrollViewer_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (_isFitWidthZoom)
        {
            FitPreviewToWidth(allowZoomIn: true);
        }
    }

    private void ScrollToPreviewPage(int page)
    {
        if (_pagePreviews.Count == 0)
        {
            return;
        }

        PdfPagesItemsControl.UpdateLayout();
        var index = Math.Clamp(page, 1, _pagePreviews.Count) - 1;
        if (PdfPagesItemsControl.ItemContainerGenerator.ContainerFromIndex(index) is FrameworkElement container)
        {
            container.BringIntoView();
        }
    }

    private void UpdateCurrentPageFromVisiblePreview()
    {
        var viewportCenterY = PdfPagesScrollViewer.ViewportHeight / 2;
        var bestPage = 0;
        var bestDistance = double.MaxValue;

        for (var index = 0; index < _pagePreviews.Count; index++)
        {
            if (PdfPagesItemsControl.ItemContainerGenerator.ContainerFromIndex(index) is not FrameworkElement container)
            {
                continue;
            }

            var bounds = container
                .TransformToAncestor(PdfPagesScrollViewer)
                .TransformBounds(new Rect(0, 0, container.ActualWidth, container.ActualHeight));

            if (bounds.Bottom < 0 || bounds.Top > PdfPagesScrollViewer.ViewportHeight)
            {
                continue;
            }

            var distance = Math.Abs((bounds.Top + bounds.Bottom) / 2 - viewportCenterY);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                bestPage = _pagePreviews[index].PageNumber;
            }
        }

        if (bestPage > 0 && bestPage != _currentPreviewPage)
        {
            _currentPreviewPage = bestPage;
            SyncCurrentPageFields();
        }
    }

    private bool ValidateInputAndOutput()
    {
        if (string.IsNullOrWhiteSpace(_inputFilePath) || !File.Exists(_inputFilePath))
        {
            WpfMessageBox.Show(this, "Vui lòng chọn tệp PDF đầu vào.", "Thiếu tệp PDF", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        if (string.IsNullOrWhiteSpace(_outputFolderPath))
        {
            WpfMessageBox.Show(this, "Vui lòng chọn thư mục xuất.", "Thiếu thư mục xuất", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    private List<PageRange> BuildCustomRanges(int pageCount)
    {
        var result = new List<PageRange>();
        foreach (var item in _ranges)
        {
            if (!int.TryParse(item.FromPage, out var from) || !int.TryParse(item.ToPage, out var to))
            {
                throw new InvalidOperationException($"{item.Label}: số trang không hợp lệ.");
            }

            ValidateRange(from, to, pageCount, item.Label);
            result.Add(new PageRange(from, to));
        }

        return result;
    }

    private List<PageRange> BuildFixedRanges(int pageCount)
    {
        if (!int.TryParse(FixedPageCountTextBox.Text, out var pagesPerFile) || pagesPerFile <= 0)
        {
            throw new InvalidOperationException("Số trang mỗi file phải lớn hơn 0.");
        }

        var result = new List<PageRange>();
        for (var from = 1; from <= pageCount; from += pagesPerFile)
        {
            result.Add(new PageRange(from, Math.Min(from + pagesPerFile - 1, pageCount)));
        }

        return result;
    }

    private CropOptions ReadCropOptions(int pageCount)
    {
        if (_cropArea is null)
        {
            throw new InvalidOperationException("Vui lòng giữ chuột phải và kéo trên preview để chọn vùng cần giữ lại.");
        }

        var cropAllPages = ApplyCropAllPagesCheckBox.IsChecked == true;
        int? targetPage = null;

        if (!cropAllPages)
        {
            if (!int.TryParse(CropPageTextBox.Text, out var page))
            {
                throw new InvalidOperationException("Trang áp dụng crop không hợp lệ.");
            }

            if (page < 1 || page > pageCount)
            {
                throw new InvalidOperationException($"Trang áp dụng crop phải nằm trong khoảng 1 đến {pageCount}.");
            }

            targetPage = page;
        }

        return new CropOptions(cropAllPages, targetPage, _cropArea.Value);
    }

    private static void ValidateRange(int from, int to, int pageCount, string label)
    {
        if (from < 1 || to < 1)
        {
            throw new InvalidOperationException($"{label}: trang bắt đầu và kết thúc phải lớn hơn 0.");
        }

        if (from > to)
        {
            throw new InvalidOperationException($"{label}: trang bắt đầu không được lớn hơn trang kết thúc.");
        }

        if (to > pageCount)
        {
            throw new InvalidOperationException($"{label}: trang kết thúc vượt quá tổng số {pageCount} trang.");
        }
    }

    private static void SaveRange(PdfDocument inputDocument, int from, int to, string outputFile, CropOptions? cropOptions)
    {
        using var outputDocument = new PdfDocument();
        for (var pageNumber = from; pageNumber <= to; pageNumber++)
        {
            var outputPage = outputDocument.AddPage(inputDocument.Pages[pageNumber - 1]);
            ApplyCropIfNeeded(outputPage, pageNumber, cropOptions);
        }

        outputDocument.Save(outputFile);
    }

    private static void SaveCroppedPdf(PdfDocument inputDocument, string outputFile, CropOptions cropOptions)
    {
        using var outputDocument = new PdfDocument();
        for (var pageNumber = 1; pageNumber <= inputDocument.PageCount; pageNumber++)
        {
            var outputPage = outputDocument.AddPage(inputDocument.Pages[pageNumber - 1]);
            ApplyCropIfNeeded(outputPage, pageNumber, cropOptions);
        }

        outputDocument.Save(outputFile);
    }

    private static void ApplyCropIfNeeded(PdfPage page, int originalPageNumber, CropOptions? cropOptions)
    {
        if (cropOptions is null)
        {
            return;
        }

        if (!cropOptions.ApplyAllPages && cropOptions.TargetPage != originalPageNumber)
        {
            return;
        }

        ApplyCrop(page, cropOptions.Area);
    }

    private static void ApplyCrop(PdfPage page, CropArea area)
    {
        var mediaBox = page.MediaBox;
        var width = mediaBox.Width;
        var height = mediaBox.Height;
        var x1 = mediaBox.X1 + area.Left * width;
        var y1 = mediaBox.Y2 - area.Bottom * height;
        var x2 = mediaBox.X1 + area.Right * width;
        var y2 = mediaBox.Y2 - area.Top * height;

        if (x2 <= x1 || y2 <= y1)
        {
            throw new InvalidOperationException("Vùng crop không hợp lệ so với kích thước trang PDF.");
        }

        page.CropBox = new PdfRectangle(new XPoint(x1, y1), new XPoint(x2, y2));
    }

    private void SyncCurrentPageFields()
    {
        var pageText = _currentPreviewPage.ToString(CultureInfo.InvariantCulture);
        PreviewPageTextBox.Text = pageText;
        CropPageTextBox.Text = pageText;
        UpdatePreviewPageHeader();
    }

    private int ReadCurrentPreviewPage()
    {
        if (_pageCount <= 0)
        {
            return 1;
        }

        if (!int.TryParse(PreviewPageTextBox.Text, out var page))
        {
            page = _currentPreviewPage;
        }

        _currentPreviewPage = Math.Clamp(page, 1, _pageCount);
        SyncCurrentPageFields();
        return _currentPreviewPage;
    }

    private void ApplyCustomSplitEndPage(int endPage)
    {
        var endPages = new SortedSet<int>();
        foreach (var item in _ranges)
        {
            if (int.TryParse(item.ToPage, out var existingEndPage) && existingEndPage >= 1 && existingEndPage < _pageCount)
            {
                endPages.Add(existingEndPage);
            }
        }

        endPages.Add(endPage);

        _ranges.Clear();
        var fromPage = 1;
        foreach (var toPage in endPages)
        {
            if (toPage < fromPage)
            {
                continue;
            }

            _ranges.Add(new PageRangeItem
            {
                Index = _ranges.Count + 1,
                FromPage = fromPage.ToString(CultureInfo.InvariantCulture),
                ToPage = toPage.ToString(CultureInfo.InvariantCulture)
            });
            fromPage = toPage + 1;
        }

        if (fromPage <= _pageCount)
        {
            _ranges.Add(new PageRangeItem
            {
                Index = _ranges.Count + 1,
                FromPage = fromPage.ToString(CultureInfo.InvariantCulture),
                ToPage = _pageCount.ToString(CultureInfo.InvariantCulture)
            });
        }

        RefreshRangeLabels();
    }

    private void UpdatePreviewPageHeader()
    {
        if (PreviewTotalPagesTextBlock is null)
        {
            return;
        }

        if (_pageCount <= 0)
        {
            PreviewPageTextBox.Text = "1";
            PreviewTotalPagesTextBlock.Text = "/ 0";
            return;
        }

        PreviewTotalPagesTextBlock.Text = $"/ {_pageCount}";
    }

    private void FitPreviewToWidth(bool allowZoomIn)
    {
        if (PdfPagesScrollViewer is null)
        {
            return;
        }

        var availableWidth = PdfPagesScrollViewer.ViewportWidth > 0
            ? PdfPagesScrollViewer.ViewportWidth
            : PdfPagesScrollViewer.ActualWidth;

        if (availableWidth <= 0)
        {
            return;
        }

        var targetZoom = availableWidth / (RenderedPreviewPageWidth + PreviewPageHorizontalChrome);
        if (!allowZoomIn)
        {
            targetZoom = Math.Min(1, targetZoom);
        }

        _isFitWidthZoom = true;
        SetPreviewZoom(targetZoom);
    }

    private void SetPreviewZoom(double zoom)
    {
        _previewZoom = Math.Clamp(zoom, PreviewZoomMin, PreviewZoomMax);
        ApplyPreviewZoom();
    }

    private void ApplyPreviewZoom()
    {
        if (PdfPreviewScaleTransform is not null)
        {
            PdfPreviewScaleTransform.ScaleX = _previewZoom;
            PdfPreviewScaleTransform.ScaleY = _previewZoom;
        }

        if (PreviewZoomTextBlock is not null)
        {
            PreviewZoomTextBlock.Text = $"{Math.Round(_previewZoom * 100)}%";
        }
    }

    private void RefreshRangeLabels()
    {
        for (var i = 0; i < _ranges.Count; i++)
        {
            _ranges[i].Index = i + 1;
        }
    }

    private void UpdatePreviewInteractionLayers()
    {
        if (CropOverlayCanvas is null || SetSplitPointButton is null || MainActionTabControl is null)
        {
            return;
        }

        CropOverlayCanvas.Visibility = MainActionTabControl.SelectedIndex == 1 ? Visibility.Visible : Visibility.Collapsed;
        SetSplitPointButton.IsEnabled =
            _pageCount > 0 && MainActionTabControl.SelectedIndex == 0 && CustomModeRadio.IsChecked == true;
    }

    private void ClearCropSelectionVisual()
    {
        _cropArea = null;
        CropSelectionRectangle.Visibility = Visibility.Collapsed;
        CropSelectionRectangle.Width = 0;
        CropSelectionRectangle.Height = 0;
        CropSelectionInfoTextBlock.Text = "Chưa chọn vùng";
    }

    private static string SanitizeFileName(string fileName)
    {
        var invalidChars = Regex.Escape(new string(Path.GetInvalidFileNameChars()));
        return Regex.Replace(fileName, $"[{invalidChars}]", "_");
    }

    private static string GetUniqueOutputPath(string path)
    {
        if (!File.Exists(path))
        {
            return path;
        }

        var folder = Path.GetDirectoryName(path)!;
        var name = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);

        for (var i = 1; i < 1000; i++)
        {
            var candidate = Path.Combine(folder, $"{name}_{i:000}{extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        return Path.Combine(folder, $"{name}_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
    }
}

public sealed record PdfPagePreviewItem(int PageNumber, ImageSource Image)
{
    public string PageLabel => $"Trang {PageNumber}";
}

public sealed class PageRangeItem : INotifyPropertyChanged
{
    private int _index;
    private string _fromPage = "";
    private string _toPage = "";

    public event PropertyChangedEventHandler? PropertyChanged;

    public int Index
    {
        get => _index;
        set
        {
            if (_index == value)
            {
                return;
            }

            _index = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(Label));
        }
    }

    public string Label => $"Khoảng {Index}";

    public string FromPage
    {
        get => _fromPage;
        set
        {
            if (_fromPage == value)
            {
                return;
            }

            _fromPage = value;
            OnPropertyChanged();
        }
    }

    public string ToPage
    {
        get => _toPage;
        set
        {
            if (_toPage == value)
            {
                return;
            }

            _toPage = value;
            OnPropertyChanged();
        }
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public readonly record struct PageRange(int From, int To);

public sealed class MergePdfItem : INotifyPropertyChanged
{
    private int _index;

    public event PropertyChangedEventHandler? PropertyChanged;

    public int Index
    {
        get => _index;
        set
        {
            if (_index == value)
            {
                return;
            }

            _index = value;
            OnPropertyChanged();
        }
    }

    public required string FilePath { get; init; }

    public required int PageCount { get; init; }

    public required long FileSizeBytes { get; init; }

    public string FileName => Path.GetFileName(FilePath);

    public string Detail => $"{PageCount} trang - {FormatFileSize(FileSizeBytes)} - {FilePath}";

    private static string FormatFileSize(long bytes)
    {
        if (bytes >= 1024 * 1024)
        {
            return $"{bytes / 1024d / 1024d:0.##} MB";
        }

        return $"{Math.Max(1, bytes / 1024d):0.##} KB";
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public readonly record struct CropArea(double Left, double Top, double Right, double Bottom)
{
    public double WidthPercent => Math.Max(0, Right - Left) * 100;

    public double HeightPercent => Math.Max(0, Bottom - Top) * 100;
}

public sealed record CropOptions(bool ApplyAllPages, int? TargetPage, CropArea Area);
