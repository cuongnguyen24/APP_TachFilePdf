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
using Microsoft.Web.WebView2.Core;
using Windows.Storage;
using Windows.Storage.Streams;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using WinPdfDocument = Windows.Data.Pdf.PdfDocument;
using WinPdfPageRenderOptions = Windows.Data.Pdf.PdfPageRenderOptions;
using WpfMessageBox = System.Windows.MessageBox;
using WpfOpenFileDialog = Microsoft.Win32.OpenFileDialog;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfColor = System.Windows.Media.Color;
using WpfPen = System.Windows.Media.Pen;
using WpfPoint = System.Windows.Point;
using WinForms = System.Windows.Forms;

namespace TachFilePdf;

public partial class MainWindow : Window
{
    private const double RenderedPreviewPageWidth = 900;
    private const double PreviewPageHorizontalChrome = 54;
    private const double PreviewZoomMin = 0.18;
    private const double PreviewZoomMax = 2.5;
    private const double PreviewZoomStep = 0.1;
    private const double MultiPagePreviewZoom = 0.22;

    private readonly ObservableCollection<PageRangeItem> _ranges = [];
    private readonly ObservableCollection<PdfPagePreviewItem> _pagePreviews = [];
    private readonly ObservableCollection<MergePdfItem> _mergeFiles = [];
    private string? _inputFilePath;
    private string? _workingPdfPath;
    private string? _outputFolderPath;
    private MergePdfItem? _activeMergeItem;
    private int _pageCount;
    private int _currentPreviewPage = 1;
    private double _previewZoom = 1;
    private bool _isFitWidthZoom;
    private PreviewLayoutMode _previewLayoutMode = PreviewLayoutMode.Vertical;
    private CoreWebView2Environment? _webViewEnvironment;
    private bool _isSelectingCrop;
    private WpfPoint _cropStartPoint;
    private CropArea? _cropArea;
    private readonly List<CoverPatch> _coverPatches = [];
    private bool _isSelectingCoverArea;
    private bool _isPickingCoverColor;
    private WpfPoint _coverStartPoint;
    private WpfColor _coverColor = WpfColor.FromRgb(255, 247, 233);

    public MainWindow()
    {
        InitializeComponent();

        _ranges.Add(new PageRangeItem { Index = 1, FromPage = "1", ToPage = "1" });
        RangesItemsControl.ItemsSource = _ranges;
        PdfPagesItemsControl.ItemsSource = _pagePreviews;
        MergeFilesListBox.ItemsSource = _mergeFiles;
        ApplyCropScopeChanged(this, new RoutedEventArgs());
        ApplyPreviewLayout();
        ApplyPreviewZoom();
        UpdatePreviewInteractionLayers();
        UpdatePreviewPageHeader();
        UpdatePrimaryActionButton();
        UpdateTabWorkspaceState();
    }

    private async void ChoosePdfButton_Click(object sender, RoutedEventArgs e)
    {
        if (IsMergeTabSelected())
        {
            return;
        }

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

        _activeMergeItem = null;
        _inputFilePath = dialog.FileName;
        _workingPdfPath = null;
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
            ClearCoverAreasVisual();
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
            ApplyZoomForCurrentPreviewLayout();
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
        if (IsMergeTabSelected())
        {
            return;
        }

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

    private async void PreviewLayoutButton_Click(object sender, RoutedEventArgs e)
    {
        var previousLayoutMode = _previewLayoutMode;
        _previewLayoutMode = _previewLayoutMode switch
        {
            PreviewLayoutMode.Vertical => PreviewLayoutMode.Grid,
            PreviewLayoutMode.Grid => PreviewLayoutMode.TextSelection,
            _ => PreviewLayoutMode.Vertical
        };

        ApplyPreviewLayout();

        if (_previewLayoutMode == PreviewLayoutMode.TextSelection && !string.IsNullOrWhiteSpace(_inputFilePath))
        {
            if (!await NavigateTextSelectionPreviewAsync(GetActivePdfPath(), _currentPreviewPage))
            {
                ApplyZoomForCurrentPreviewLayout();
                return;
            }
        }
        else if (previousLayoutMode == PreviewLayoutMode.TextSelection && !string.IsNullOrWhiteSpace(_inputFilePath))
        {
            await LoadPdfPreviewPagesAsync(GetActivePdfPath());
        }

        if (previousLayoutMode == PreviewLayoutMode.Compatibility && !string.IsNullOrWhiteSpace(_inputFilePath))
        {
            await LoadPdfPreviewPagesAsync(GetActivePdfPath());
        }

        ApplyZoomForCurrentPreviewLayout();
        StatusTextBlock.Text = _previewLayoutMode switch
        {
            PreviewLayoutMode.Grid => "Đã đổi sang chế độ xem lưới nhiều trang.",
            PreviewLayoutMode.TextSelection => "Đã đổi sang chế độ chọn/copy text bằng WebView2.",
            PreviewLayoutMode.Compatibility => "Renderer ảnh không mở được file này. Đang dùng chế độ tương thích PDF/A.",
            _ => "Đã đổi sang chế độ xem dọc từng trang."
        };

        if (_previewLayoutMode != PreviewLayoutMode.TextSelection)
        {
            ScrollToPreviewPage(_currentPreviewPage);
        }
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

    private void SetCurrentPageAsSplitPoint()
    {
        if (_pageCount <= 0 || !IsSplitTabSelected() || CustomModeRadio.IsChecked != true)
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
            UpdateTabWorkspaceState();
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

    private void EditOverlayCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_pageCount <= 0)
        {
            return;
        }

        if (_isPickingCoverColor)
        {
            PickCoverColorAtMouse(e);
            e.Handled = true;
            return;
        }

        var point = e.GetPosition(EditOverlayCanvas);
        var pageBounds = GetCurrentPageImageBoundsOnOverlay(EditOverlayCanvas);
        if (pageBounds.IsEmpty || !pageBounds.Contains(point))
        {
            StatusTextBlock.Text = "Hãy kéo trong vùng trang PDF để thêm vùng bù màu.";
            return;
        }

        _isSelectingCoverArea = true;
        _coverStartPoint = ClampPointToRect(point, pageBounds);
        EditSelectionRectangle.Visibility = Visibility.Visible;
        Canvas.SetLeft(EditSelectionRectangle, _coverStartPoint.X);
        Canvas.SetTop(EditSelectionRectangle, _coverStartPoint.Y);
        EditSelectionRectangle.Width = 0;
        EditSelectionRectangle.Height = 0;
        EditOverlayCanvas.CaptureMouse();
        e.Handled = true;
    }

    private void EditOverlayCanvas_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isSelectingCoverArea)
        {
            return;
        }

        var pageBounds = GetCurrentPageImageBoundsOnOverlay(EditOverlayCanvas);
        if (pageBounds.IsEmpty)
        {
            return;
        }

        var currentPoint = ClampPointToRect(e.GetPosition(EditOverlayCanvas), pageBounds);
        var left = Math.Min(_coverStartPoint.X, currentPoint.X);
        var top = Math.Min(_coverStartPoint.Y, currentPoint.Y);
        var width = Math.Abs(currentPoint.X - _coverStartPoint.X);
        var height = Math.Abs(currentPoint.Y - _coverStartPoint.Y);

        Canvas.SetLeft(EditSelectionRectangle, left);
        Canvas.SetTop(EditSelectionRectangle, top);
        EditSelectionRectangle.Width = width;
        EditSelectionRectangle.Height = height;
    }

    private void EditOverlayCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_isSelectingCoverArea)
        {
            return;
        }

        _isSelectingCoverArea = false;
        EditOverlayCanvas.ReleaseMouseCapture();

        if (EditSelectionRectangle.Width < 8 || EditSelectionRectangle.Height < 8)
        {
            EditSelectionRectangle.Visibility = Visibility.Collapsed;
            e.Handled = true;
            return;
        }

        var pageBounds = GetCurrentPageImageBoundsOnOverlay(EditOverlayCanvas);
        if (pageBounds.IsEmpty)
        {
            EditSelectionRectangle.Visibility = Visibility.Collapsed;
            e.Handled = true;
            return;
        }

        var left = Canvas.GetLeft(EditSelectionRectangle);
        var top = Canvas.GetTop(EditSelectionRectangle);
        var right = left + EditSelectionRectangle.Width;
        var bottom = top + EditSelectionRectangle.Height;
        var area = new CropArea(
            Math.Clamp((left - pageBounds.Left) / pageBounds.Width, 0, 1),
            Math.Clamp((top - pageBounds.Top) / pageBounds.Height, 0, 1),
            Math.Clamp((right - pageBounds.Left) / pageBounds.Width, 0, 1),
            Math.Clamp((bottom - pageBounds.Top) / pageBounds.Height, 0, 1));

        _coverPatches.Add(new CoverPatch(
            _currentPreviewPage,
            ApplyCoverAllPagesCheckBox.IsChecked == true,
            area,
            _coverColor,
            CoverOpacitySlider.Value / 100.0));

        EditSelectionRectangle.Visibility = Visibility.Collapsed;
        RefreshCoverOverlayVisuals();
        UpdateCoverAreasInfo();
        StatusTextBlock.Text = "Đã thêm vùng bù màu. Có thể kéo thêm vùng khác hoặc lưu PDF.";
        e.Handled = true;
    }

    private void ChooseCoverColorButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new WinForms.ColorDialog
        {
            FullOpen = true,
            Color = System.Drawing.Color.FromArgb(_coverColor.R, _coverColor.G, _coverColor.B)
        };

        if (dialog.ShowDialog() != WinForms.DialogResult.OK)
        {
            return;
        }

        SetCoverColor(WpfColor.FromRgb(dialog.Color.R, dialog.Color.G, dialog.Color.B));
    }

    private void PickCoverColorButton_Click(object sender, RoutedEventArgs e)
    {
        _isPickingCoverColor = true;
        StatusTextBlock.Text = "Bút lấy màu đang bật: bấm chuột trái lên nền trong preview để lấy màu.";
    }

    private void UndoCoverAreaButton_Click(object sender, RoutedEventArgs e)
    {
        if (_coverPatches.Count == 0)
        {
            StatusTextBlock.Text = "Chưa có vùng bù màu để hoàn tác.";
            return;
        }

        _coverPatches.RemoveAt(_coverPatches.Count - 1);
        RefreshCoverOverlayVisuals();
        UpdateCoverAreasInfo();
        StatusTextBlock.Text = "Đã hoàn tác vùng bù màu cuối.";
    }

    private void ClearCoverAreasButton_Click(object sender, RoutedEventArgs e)
    {
        ClearCoverAreasVisual();
        StatusTextBlock.Text = "Đã xóa tất cả vùng bù màu.";
    }

    private async void SaveEditedPdfButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateInputAndOutput())
        {
            return;
        }

        if (_coverPatches.Count == 0)
        {
            WpfMessageBox.Show(this, "Vui lòng kéo ít nhất một vùng bù màu trên preview.", "Chưa có vùng bù màu", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            Directory.CreateDirectory(_outputFolderPath!);
            using var inputDocument = PdfReader.Open(GetActivePdfPath(), PdfDocumentOpenMode.Import);
            var baseName = SanitizeFileName(Path.GetFileNameWithoutExtension(_inputFilePath!));
            var outputFile = GetUniqueOutputPath(Path.Combine(_outputFolderPath!, $"{baseName}_bu_mau.pdf"));
            SaveCoveredPdf(inputDocument, outputFile, _coverPatches);

            var workingCopy = CreateWorkingPdfPath();
            File.Copy(outputFile, workingCopy, overwrite: true);
            SetWorkingPdf(workingCopy);
            ClearCoverAreasVisual();
            await ReloadActivePdfPreviewAsync(_currentPreviewPage);

            StatusTextBlock.Text = $"Đã lưu PDF bù màu: {outputFile}";
            WpfMessageBox.Show(this, "Đã lưu PDF đã bù màu.", "Hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Lưu PDF bù màu thất bại.";
            WpfMessageBox.Show(this, ex.Message, "Lỗi bù màu", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void PrimaryActionButton_Click(object sender, RoutedEventArgs e)
    {
        if (IsMergeTabSelected())
        {
            MergePdfButton_Click(sender, e);
            return;
        }

        if (IsEditTabSelected())
        {
            SaveEditedPdfButton_Click(sender, e);
            return;
        }

        SplitPdfButton_Click(sender, e);
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

            using var inputDocument = PdfReader.Open(GetActivePdfPath(), PdfDocumentOpenMode.Import);
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
                    $"{baseName}_{range.From:000}-{range.To:000}.pdf");
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

            using var inputDocument = PdfReader.Open(GetActivePdfPath(), PdfDocumentOpenMode.Import);
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

    private async void MergeFilesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!IsMergeTabSelected() || MergeFilesListBox.SelectedItem is not MergePdfItem item)
        {
            return;
        }

        await Task.Yield();
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
                using var inputDocument = PdfReader.Open(item.ActiveFilePath, PdfDocumentOpenMode.Import);
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
            if (!File.Exists(item.ActiveFilePath))
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

        PrimaryActionButton.Content = IsMergeTabSelected()
            ? "Gộp PDF"
            : IsEditTabSelected()
                ? "Lưu PDF đã bù màu"
                : "Tách PDF";
    }

    private void UpdateTabWorkspaceState()
    {
        if (PreviewDisabledOverlay is null ||
            PreviewToolbarPanel is null ||
            DocumentPickerPanel is null ||
            PdfPagesScrollViewer is null ||
            CropOverlayCanvas is null ||
            EditOverlayCanvas is null ||
            StatusTextBlock is null)
        {
            return;
        }

        var isMergeTab = IsMergeTabSelected();
        PreviewDisabledOverlay.Visibility = isMergeTab ? Visibility.Visible : Visibility.Collapsed;
        PreviewToolbarPanel.IsEnabled = !isMergeTab;
        PreviewToolbarPanel.Opacity = isMergeTab ? 0.42 : 1;
        PdfPagesScrollViewer.IsEnabled = !isMergeTab;
        if (TextSelectionPdfWebView is not null)
        {
            TextSelectionPdfWebView.IsEnabled = !isMergeTab;
        }
        DocumentPickerPanel.IsEnabled = !isMergeTab;
        DocumentPickerPanel.Opacity = isMergeTab ? 0.46 : 1;

        if (isMergeTab)
        {
            CropOverlayCanvas.Visibility = Visibility.Collapsed;
            EditOverlayCanvas.Visibility = Visibility.Collapsed;
            _activeMergeItem = null;
            StatusTextBlock.Text = "Tab Gộp PDF đang hoạt động. Preview tách file đã tạm ẩn.";
        }
        else
        {
            _activeMergeItem = null;
            StatusTextBlock.Text = _pageCount > 0 ? $"Đã chọn PDF có {_pageCount} trang." : "Sẵn sàng.";
        }
    }

    private bool IsSplitTabSelected()
    {
        return MainActionTabControl?.SelectedItem == SplitTabItem;
    }

    private bool IsMergeTabSelected()
    {
        return MainActionTabControl?.SelectedItem == MergeTabItem;
    }

    private bool IsEditTabSelected()
    {
        return MainActionTabControl?.SelectedItem == EditTabItem;
    }

    private async Task LoadPdfPreviewPagesAsync(string filePath)
    {
        _pagePreviews.Clear();

        if (_previewLayoutMode == PreviewLayoutMode.Compatibility)
        {
            await LoadCompatibilityPreviewPagesAsync(filePath);
            return;
        }

        if (_previewLayoutMode == PreviewLayoutMode.TextSelection)
        {
            await NavigateTextSelectionPreviewAsync(filePath, _currentPreviewPage);
            return;
        }

        try
        {
            var storageFile = await StorageFile.GetFileFromPathAsync(filePath);
            var pdfDocument = await WinPdfDocument.LoadFromFileAsync(storageFile);
            for (uint pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
            {
                using var page = pdfDocument.GetPage(pageIndex);
                var image = await RenderPdfPageAsync(page);
                _pagePreviews.Add(CreatePagePreviewItem((int)pageIndex + 1, image));

                if (pageIndex % 5 == 0 || pageIndex + 1 == pdfDocument.PageCount)
                {
                    StatusTextBlock.Text = $"Đang render trang {pageIndex + 1}/{pdfDocument.PageCount}...";
                    await Task.Yield();
                }
            }
        }
        catch when (_pageCount > 0)
        {
            _previewLayoutMode = PreviewLayoutMode.Compatibility;
            ApplyPreviewLayout();
            await LoadCompatibilityPreviewPagesAsync(filePath);
            StatusTextBlock.Text = "Renderer ảnh không mở được file này. Đã chuyển sang chế độ tương thích PDF/A.";
        }
    }

    private async Task LoadCompatibilityPreviewPagesAsync(string filePath)
    {
        _pagePreviews.Clear();
        using var document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
        for (var pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
        {
            var page = document.Pages[pageIndex];
            var image = RenderCompatibilityPagePreview(page, pageIndex + 1, document.PageCount);
            _pagePreviews.Add(CreatePagePreviewItem(pageIndex + 1, image, true));

            if (pageIndex % 20 == 0 || pageIndex + 1 == document.PageCount)
            {
                StatusTextBlock.Text = $"Đang tạo view tương thích trang {pageIndex + 1}/{document.PageCount}...";
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

    private async Task<bool> NavigateTextSelectionPreviewAsync(string filePath, int pageNumber)
    {
        if (TextSelectionPdfWebView is null || !File.Exists(filePath))
        {
            return false;
        }

        try
        {
            await TextSelectionPdfWebView.EnsureCoreWebView2Async(await GetWebViewEnvironmentAsync());
            TextSelectionPdfWebView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            TextSelectionPdfWebView.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = true;

            var page = Math.Clamp(pageNumber, 1, Math.Max(_pageCount, 1));
            TextSelectionPdfWebView.CoreWebView2.Navigate($"{new Uri(filePath).AbsoluteUri}#page={page}");
            return true;
        }
        catch (Exception ex) when (ex is WebView2RuntimeNotFoundException or InvalidOperationException or UnauthorizedAccessException)
        {
            _previewLayoutMode = PreviewLayoutMode.Vertical;
            ApplyPreviewLayout();
            await LoadPdfPreviewPagesAsync(filePath);
            StatusTextBlock.Text = $"Không mở được chế độ chọn/copy text ({ex.Message}). Đã quay lại chế độ xem dọc.";
            return false;
        }
    }

    private async Task<CoreWebView2Environment> GetWebViewEnvironmentAsync()
    {
        if (_webViewEnvironment is not null)
        {
            return _webViewEnvironment;
        }

        var userDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ToolAXE",
            "TachFilePdf",
            "WebView2");
        Directory.CreateDirectory(userDataFolder);
        _webViewEnvironment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
        return _webViewEnvironment;
    }

    private static ImageSource RenderCompatibilityPagePreview(PdfPage page, int pageNumber, int pageCount)
    {
        const int width = 620;
        const int height = 860;
        var pageWidthPoints = Math.Max(1, page.MediaBox.Width);
        var pageHeightPoints = Math.Max(1, page.MediaBox.Height);
        var pageWidthInches = pageWidthPoints / 72d;
        var pageHeightInches = pageHeightPoints / 72d;
        var pageWidthMm = pageWidthInches * 25.4;
        var pageHeightMm = pageHeightInches * 25.4;

        var visual = new DrawingVisual();
        using (var context = visual.RenderOpen())
        {
            context.DrawRectangle(WpfBrushes.White, new WpfPen(new SolidColorBrush(WpfColor.FromRgb(221, 226, 234)), 2), new Rect(0, 0, width, height));
            context.DrawRectangle(new SolidColorBrush(WpfColor.FromRgb(248, 250, 252)), null, new Rect(28, 28, width - 56, height - 56));
            DrawCenteredText(context, $"Trang {pageNumber}/{pageCount}", 34, 28, width - 56, 28, WpfBrushes.Black);
            DrawCenteredText(context, "Chế độ tương thích PDF/A", 28, 60, width - 120, 28, new SolidColorBrush(WpfColor.FromRgb(239, 35, 41)));
            DrawCenteredText(context, "Không dùng renderer ảnh của Windows", 22, 92, width - 120, 24, new SolidColorBrush(WpfColor.FromRgb(100, 116, 139)));

            var infoTop = 190;
            DrawCenteredText(context, $"{pageWidthInches:0.##}\" x {pageHeightInches:0.##}\"", 42, infoTop, width - 80, 48, WpfBrushes.Black);
            DrawCenteredText(context, $"{pageWidthMm:0.#} x {pageHeightMm:0.#} mm", 28, infoTop + 58, width - 80, 34, new SolidColorBrush(WpfColor.FromRgb(51, 65, 85)));
            DrawCenteredText(context, $"Rotation: {page.Rotate}°", 24, infoTop + 110, width - 80, 30, new SolidColorBrush(WpfColor.FromRgb(100, 116, 139)));

            context.DrawRectangle(null, new WpfPen(new SolidColorBrush(WpfColor.FromRgb(239, 35, 41)), 2), new Rect(98, 410, width - 196, 250));
            DrawCenteredText(context, "Right-click để chỉnh sửa,", 26, 475, width - 120, 34, new SolidColorBrush(WpfColor.FromRgb(15, 23, 42)));
            DrawCenteredText(context, "kiểm tra DPI/PDF-A,", 26, 512, width - 120, 34, new SolidColorBrush(WpfColor.FromRgb(15, 23, 42)));
            DrawCenteredText(context, "hoặc tách theo khoảng.", 26, 549, width - 120, 34, new SolidColorBrush(WpfColor.FromRgb(15, 23, 42)));
        }

        var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(visual);
        bitmap.Freeze();
        return bitmap;
    }

    private static void DrawCenteredText(DrawingContext context, string text, double fontSize, double top, double width, double height, System.Windows.Media.Brush brush)
    {
        var formattedText = new FormattedText(
            text,
            CultureInfo.CurrentCulture,
            System.Windows.FlowDirection.LeftToRight,
            new Typeface("Segoe UI"),
            fontSize,
            brush,
            1.0)
        {
            TextAlignment = TextAlignment.Center,
            MaxTextWidth = width,
            MaxTextHeight = height
        };

        context.DrawText(formattedText, new WpfPoint((620 - width) / 2, top));
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

    private async void RotatePageLeftMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await EditPreviewPageAsync(sender, PageEditAction.RotateLeft);
    }

    private async void RotatePageRightMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await EditPreviewPageAsync(sender, PageEditAction.RotateRight);
    }

    private void CheckPageDpiMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (GetPageItemFromMenuSender(sender) is not { } pageItem)
        {
            return;
        }

        try
        {
            using var document = PdfReader.Open(GetActivePdfPath(), PdfDocumentOpenMode.Import);
            var result = AnalyzePageDpi(document.Pages[pageItem.PageNumber - 1], pageItem.PageNumber, document.PageCount);
            StatusTextBlock.Text = result.StatusText;
            WpfMessageBox.Show(this, result.Message, "Kiểm tra DPI", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Kiểm tra DPI thất bại.";
            WpfMessageBox.Show(this, ex.Message, "Lỗi kiểm tra DPI", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CheckPdfStandardMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var result = AnalyzePdfStandard(GetActivePdfPath());
            StatusTextBlock.Text = result.StatusText;
            WpfMessageBox.Show(this, result.Message, "Kiểm tra chuẩn PDF/A/PDF", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Kiểm tra chuẩn PDF thất bại.";
            WpfMessageBox.Show(this, ex.Message, "Lỗi kiểm tra chuẩn PDF", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void DeletePageMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await EditPreviewPageAsync(sender, PageEditAction.Delete);
    }

    private async void DuplicatePageMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await EditPreviewPageAsync(sender, PageEditAction.Duplicate);
    }

    private async void MovePagePreviousMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await EditPreviewPageAsync(sender, PageEditAction.MovePrevious);
    }

    private async void MovePageNextMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await EditPreviewPageAsync(sender, PageEditAction.MoveNext);
    }

    private async void MovePageFirstMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await EditPreviewPageAsync(sender, PageEditAction.MoveFirst);
    }

    private async void MovePageLastMenuItem_Click(object sender, RoutedEventArgs e)
    {
        await EditPreviewPageAsync(sender, PageEditAction.MoveLast);
    }

    private async void ResetPdfEditsMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (_inputFilePath is null)
        {
            return;
        }

        ResetWorkingPdf();
        await ReloadActivePdfPreviewAsync(Math.Clamp(_currentPreviewPage, 1, Math.Max(_pageCount, 1)));
        StatusTextBlock.Text = "Đã hoàn tác chỉnh sửa và quay lại file PDF gốc.";
    }

    private void PdfPagesScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        if (_pageCount <= 0 || PreviewPageTextBox.IsKeyboardFocusWithin || _previewLayoutMode == PreviewLayoutMode.TextSelection)
        {
            return;
        }

        UpdateCurrentPageFromVisiblePreview();
        RefreshCoverOverlayVisuals();
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
        if (_previewLayoutMode == PreviewLayoutMode.TextSelection)
        {
            _ = NavigateTextSelectionPreviewAsync(GetActivePdfPath(), page);
            return;
        }

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
        if (_previewLayoutMode != PreviewLayoutMode.Vertical)
        {
            UpdateCurrentPageFromMultiPagePreview();
            return;
        }

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

    private void UpdateCurrentPageFromMultiPagePreview()
    {
        var viewportCenterX = PdfPagesScrollViewer.ViewportWidth / 2;
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

            if (bounds.Right < 0 ||
                bounds.Left > PdfPagesScrollViewer.ViewportWidth ||
                bounds.Bottom < 0 ||
                bounds.Top > PdfPagesScrollViewer.ViewportHeight)
            {
                continue;
            }

            var centerX = (bounds.Left + bounds.Right) / 2;
            var centerY = (bounds.Top + bounds.Bottom) / 2;
            var distance = Math.Abs(centerX - viewportCenterX) + Math.Abs(centerY - viewportCenterY);
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

    private async Task EditPreviewPageAsync(object sender, PageEditAction action)
    {
        if (GetPageItemFromMenuSender(sender) is not { } pageItem || _inputFilePath is null)
        {
            return;
        }

        try
        {
            var requestedPage = pageItem.PageNumber;
            if (action == PageEditAction.Delete && _pageCount <= 1)
            {
                WpfMessageBox.Show(this, "PDF phải còn ít nhất 1 trang.", "Không thể xóa trang", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var sourcePath = GetActivePdfPath();
            var editedPath = CreateWorkingPdfPath();
            var newCurrentPage = ApplyPageEdit(sourcePath, editedPath, requestedPage, action);
            SetWorkingPdf(editedPath);
            UpdatePreviewAfterPageEdit(action, requestedPage, newCurrentPage);
            await Task.Yield();
            StatusTextBlock.Text = BuildPageEditStatus(action, requestedPage);
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = "Chỉnh sửa PDF thất bại.";
            WpfMessageBox.Show(this, ex.Message, "Lỗi chỉnh sửa PDF", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static PdfPagePreviewItem? GetPageItemFromMenuSender(object sender)
    {
        if (sender is FrameworkElement { DataContext: PdfPagePreviewItem directItem })
        {
            return directItem;
        }

        if (sender is MenuItem menuItem)
        {
            var current = menuItem.Parent;
            while (current is FrameworkElement element)
            {
                if (element.DataContext is PdfPagePreviewItem item)
                {
                    return item;
                }

                current = element.Parent;
            }
        }

        return null;
    }

    private async Task ReloadActivePdfPreviewAsync(int preferredPage)
    {
        var activePath = GetActivePdfPath();
        using (var document = PdfReader.Open(activePath, PdfDocumentOpenMode.Import))
        {
            _pageCount = document.PageCount;
        }

        if (_activeMergeItem is not null)
        {
            _activeMergeItem.PageCount = _pageCount;
        }

        _currentPreviewPage = Math.Clamp(preferredPage, 1, Math.Max(_pageCount, 1));
        SyncCurrentPageFields();
        ClampSplitRangesToPageCount();
        UpdatePreviewPageHeader();
        PreviewPlaceholderPanel.Visibility = Visibility.Collapsed;
        _pagePreviews.Clear();
        StatusTextBlock.Text = _previewLayoutMode == PreviewLayoutMode.TextSelection
            ? "Đang mở PDF ở chế độ chọn/copy text..."
            : $"Đang render {_pageCount} trang PDF...";
        await LoadPdfPreviewPagesAsync(activePath);
        ApplyZoomForCurrentPreviewLayout();

        ScrollToPreviewPage(_currentPreviewPage);
        DocumentInfoTextBlock.Text = $"{_pageCount} trang";
    }

    private int ApplyPageEdit(string sourcePath, string editedPath, int pageNumber, PageEditAction action)
    {
        using var inputDocument = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Import);
        using var outputDocument = new PdfDocument();

        var pageCount = inputDocument.PageCount;
        var pageIndex = Math.Clamp(pageNumber, 1, pageCount) - 1;
        var order = Enumerable.Range(0, pageCount).Select(index => new PageCopyPlan(index)).ToList();

        switch (action)
        {
            case PageEditAction.RotateLeft:
                order[pageIndex] = order[pageIndex] with { RotateDelta = -90 };
                break;
            case PageEditAction.RotateRight:
                order[pageIndex] = order[pageIndex] with { RotateDelta = 90 };
                break;
            case PageEditAction.Delete:
                order.RemoveAt(pageIndex);
                break;
            case PageEditAction.Duplicate:
                order.Insert(pageIndex + 1, order[pageIndex]);
                break;
            case PageEditAction.MovePrevious:
                if (pageIndex > 0)
                {
                    (order[pageIndex - 1], order[pageIndex]) = (order[pageIndex], order[pageIndex - 1]);
                    pageIndex--;
                }
                break;
            case PageEditAction.MoveNext:
                if (pageIndex < order.Count - 1)
                {
                    (order[pageIndex], order[pageIndex + 1]) = (order[pageIndex + 1], order[pageIndex]);
                    pageIndex++;
                }
                break;
            case PageEditAction.MoveFirst:
                if (pageIndex > 0)
                {
                    var item = order[pageIndex];
                    order.RemoveAt(pageIndex);
                    order.Insert(0, item);
                    pageIndex = 0;
                }
                break;
            case PageEditAction.MoveLast:
                if (pageIndex < order.Count - 1)
                {
                    var item = order[pageIndex];
                    order.RemoveAt(pageIndex);
                    order.Add(item);
                    pageIndex = order.Count - 1;
                }
                break;
        }

        foreach (var plan in order)
        {
            var outputPage = outputDocument.AddPage(inputDocument.Pages[plan.SourceIndex]);
            if (plan.RotateDelta != 0)
            {
                outputPage.Rotate = NormalizePdfRotation(outputPage.Rotate + plan.RotateDelta);
            }
        }

        outputDocument.Save(editedPath);
        return action == PageEditAction.Delete
            ? Math.Clamp(pageNumber, 1, Math.Max(order.Count, 1))
            : pageIndex + 1;
    }

    private void UpdatePreviewAfterPageEdit(PageEditAction action, int pageNumber, int newCurrentPage)
    {
        var index = Math.Clamp(pageNumber, 1, Math.Max(_pagePreviews.Count, 1)) - 1;
        switch (action)
        {
            case PageEditAction.RotateLeft:
                ReplacePreviewImage(index, RotateImageSource(_pagePreviews[index].Image, -90));
                break;
            case PageEditAction.RotateRight:
                ReplacePreviewImage(index, RotateImageSource(_pagePreviews[index].Image, 90));
                break;
            case PageEditAction.Delete:
                if (_pagePreviews.Count > 1)
                {
                    _pagePreviews.RemoveAt(index);
                }
                break;
            case PageEditAction.Duplicate:
                _pagePreviews.Insert(index + 1, CreatePagePreviewItem(index + 2, _pagePreviews[index].Image));
                break;
            case PageEditAction.MovePrevious:
                if (index > 0)
                {
                    _pagePreviews.Move(index, index - 1);
                }
                break;
            case PageEditAction.MoveNext:
                if (index < _pagePreviews.Count - 1)
                {
                    _pagePreviews.Move(index, index + 1);
                }
                break;
            case PageEditAction.MoveFirst:
                if (index > 0)
                {
                    _pagePreviews.Move(index, 0);
                }
                break;
            case PageEditAction.MoveLast:
                if (index < _pagePreviews.Count - 1)
                {
                    _pagePreviews.Move(index, _pagePreviews.Count - 1);
                }
                break;
        }

        RefreshPreviewPageNumbers();
        _pageCount = _pagePreviews.Count;
        if (_activeMergeItem is not null)
        {
            _activeMergeItem.PageCount = _pageCount;
        }

        DocumentInfoTextBlock.Text = $"{_pageCount} trang";
        _currentPreviewPage = Math.Clamp(newCurrentPage, 1, Math.Max(_pageCount, 1));
        SyncCurrentPageFields();
        ClampSplitRangesToPageCount();
        UpdatePreviewPageHeader();
        ScrollToPreviewPage(_currentPreviewPage);
    }

    private void ReplacePreviewImage(int index, ImageSource image)
    {
        if (index < 0 || index >= _pagePreviews.Count)
        {
            return;
        }

        _pagePreviews[index] = CreatePagePreviewItem(
            _pagePreviews[index].PageNumber,
            image,
            _pagePreviews[index].IsCompatibilityView);
    }

    private void RefreshPreviewPageNumbers()
    {
        for (var index = 0; index < _pagePreviews.Count; index++)
        {
            if (_pagePreviews[index].PageNumber != index + 1)
            {
                _pagePreviews[index] = CreatePagePreviewItem(
                    index + 1,
                    _pagePreviews[index].Image,
                    _pagePreviews[index].IsCompatibilityView);
            }
        }
    }

    private static ImageSource RotateImageSource(ImageSource source, double angle)
    {
        var transform = new RotateTransform(angle);
        var rotated = new TransformedBitmap((BitmapSource)source, transform);
        rotated.Freeze();
        return rotated;
    }

    private static PageDpiResult AnalyzePageDpi(PdfPage page, int pageNumber, int pageCount)
    {
        var pageWidthPoints = Math.Max(0.01, page.MediaBox.Width);
        var pageHeightPoints = Math.Max(0.01, page.MediaBox.Height);
        var pageWidthInches = pageWidthPoints / 72d;
        var pageHeightInches = pageHeightPoints / 72d;
        var images = new List<PdfImageInfo>();
        CollectPageImages(page.Elements.GetDictionary("/Resources"), images, new HashSet<string>());

        if (images.Count == 0)
        {
            var message =
                $"Trang {pageNumber}/{pageCount}\n" +
                $"Kích thước trang: {pageWidthInches:0.##}\" x {pageHeightInches:0.##}\"\n\n" +
                "Không phát hiện ảnh raster trong trang này.\n" +
                "Kết luận: Text/vector - không có DPI gốc cố định.";
            return new PageDpiResult("Trang này là text/vector, không có DPI gốc cố định.", message);
        }

        var mainImage = images
            .OrderByDescending(image => (long)image.PixelWidth * image.PixelHeight)
            .First();
        var estimates = images
            .SelectMany(image => new[]
            {
                image.PixelWidth / pageWidthInches,
                image.PixelHeight / pageHeightInches
            })
            .Where(double.IsFinite)
            .ToList();
        var mainDpi = Math.Min(mainImage.PixelWidth / pageWidthInches, mainImage.PixelHeight / pageHeightInches);
        var minDpi = estimates.Min();
        var maxDpi = estimates.Max();
        var conclusion = DpiConclusion(mainDpi);
        var messageText =
            $"Trang {pageNumber}/{pageCount}\n" +
            $"Kích thước trang: {pageWidthInches:0.##}\" x {pageHeightInches:0.##}\"\n" +
            $"Số ảnh raster phát hiện: {images.Count}\n\n" +
            $"Ảnh chính: {mainImage.PixelWidth} x {mainImage.PixelHeight}px\n" +
            $"DPI ước tính ảnh chính: {Math.Round(mainDpi)} DPI\n" +
            $"DPI thấp nhất/cao nhất ước tính: {Math.Round(minDpi)} / {Math.Round(maxDpi)} DPI\n" +
            $"Kết luận: {conclusion}\n\n" +
            "Ghi chú: kết quả này ước tính theo kích thước trang. Với PDF scan toàn trang thường rất sát; PDF nhiều ảnh nhỏ có thể cần kiểm tra bằng tool PDF.js chuyên sâu.";

        return new PageDpiResult($"Trang {pageNumber}: khoảng {Math.Round(mainDpi)} DPI - {conclusion}.", messageText);
    }

    private static PdfStandardResult AnalyzePdfStandard(string filePath)
    {
        using var document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);
        var headerVersion = ReadPdfHeaderVersion(filePath);
        var catalogVersion = document.Internals.Catalog.Elements.GetName("/Version");
        var effectiveVersion = string.IsNullOrWhiteSpace(catalogVersion)
            ? headerVersion
            : catalogVersion.TrimStart('/');
        var metadata = ReadDocumentMetadata(document);
        var pdfaPart = FindXmlTagValue(metadata, "pdfaid:part") ?? FindXmlAttributeValue(metadata, "pdfaid:part");
        var pdfaConformance = FindXmlTagValue(metadata, "pdfaid:conformance") ?? FindXmlAttributeValue(metadata, "pdfaid:conformance");
        var pdfaRev = FindXmlTagValue(metadata, "pdfaid:rev") ?? FindXmlAttributeValue(metadata, "pdfaid:rev");
        var hasPdfA = !string.IsNullOrWhiteSpace(pdfaPart);
        var pdfLabel = hasPdfA
            ? BuildPdfALabel(pdfaPart!, pdfaConformance, pdfaRev)
            : "Không thấy khai báo PDF/A trong XMP metadata";

        var message =
            $"File: {Path.GetFileName(filePath)}\n" +
            $"Số trang: {document.PageCount}\n" +
            $"PDF version: {effectiveVersion}\n" +
            $"PDF/A: {pdfLabel}\n\n" +
            $"Metadata XMP: {(string.IsNullOrWhiteSpace(metadata) ? "Không tìm thấy" : "Có")}\n\n" +
            "Lưu ý: kiểm tra này phát hiện khai báo chuẩn trong metadata và thông tin PDF cơ bản. Nó không thay thế kiểm định conformance đầy đủ bằng validator chuyên dụng.";

        var status = hasPdfA
            ? $"Phát hiện {BuildPdfALabel(pdfaPart!, pdfaConformance, pdfaRev)}."
            : "Không thấy khai báo PDF/A trong metadata.";
        return new PdfStandardResult(status, message);
    }

    private static string BuildPdfALabel(string part, string? conformance, string? rev)
    {
        var label = $"PDF/A-{part}";
        if (!string.IsNullOrWhiteSpace(conformance))
        {
            label += conformance.Trim().ToLowerInvariant();
        }

        if (!string.IsNullOrWhiteSpace(rev))
        {
            label += $" rev {rev.Trim()}";
        }

        return label;
    }

    private static string ReadPdfHeaderVersion(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        Span<byte> header = stackalloc byte[32];
        var read = stream.Read(header);
        var text = System.Text.Encoding.ASCII.GetString(header[..read]);
        var match = Regex.Match(text, @"%PDF-(\d\.\d)");
        return match.Success ? match.Groups[1].Value : "Không xác định";
    }

    private static string? ReadDocumentMetadata(PdfDocument document)
    {
        if (document.Internals.Catalog.Elements["/Metadata"] is not PdfReference reference ||
            reference.Value is not PdfDictionary metadataDictionary ||
            metadataDictionary.Stream?.Value is null)
        {
            return null;
        }

        return System.Text.Encoding.UTF8.GetString(metadataDictionary.Stream.Value);
    }

    private static string? FindXmlTagValue(string? xml, string tagName)
    {
        if (string.IsNullOrWhiteSpace(xml))
        {
            return null;
        }

        var escaped = Regex.Escape(tagName);
        var match = Regex.Match(xml, $@"<{escaped}[^>]*>(.*?)</{escaped}>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        return match.Success ? Regex.Replace(match.Groups[1].Value, @"\s+", " ").Trim() : null;
    }

    private static string? FindXmlAttributeValue(string? xml, string attributeName)
    {
        if (string.IsNullOrWhiteSpace(xml))
        {
            return null;
        }

        var escaped = Regex.Escape(attributeName);
        var match = Regex.Match(xml, $@"{escaped}\s*=\s*[""']([^""']+)[""']", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value.Trim() : null;
    }

    private static void CollectPageImages(PdfDictionary? resources, List<PdfImageInfo> images, HashSet<string> visited)
    {
        if (resources is null)
        {
            return;
        }

        var xObjects = resources.Elements.GetDictionary("/XObject");
        if (xObjects is null)
        {
            return;
        }

        foreach (var name in xObjects.Elements.KeyNames)
        {
            if (xObjects.Elements[name] is not PdfReference reference || reference.Value is not PdfDictionary xObject)
            {
                continue;
            }

            var objectId = reference.ObjectID.ToString();
            if (!visited.Add(objectId))
            {
                continue;
            }

            var subtype = xObject.Elements.GetName("/Subtype");
            if (string.Equals(subtype, "/Image", StringComparison.Ordinal))
            {
                var width = xObject.Elements.GetInteger("/Width");
                var height = xObject.Elements.GetInteger("/Height");
                if (width > 0 && height > 0)
                {
                    images.Add(new PdfImageInfo(width, height));
                }
            }
            else if (string.Equals(subtype, "/Form", StringComparison.Ordinal))
            {
                CollectPageImages(xObject.Elements.GetDictionary("/Resources"), images, visited);
            }
        }
    }

    private static string DpiConclusion(double dpi)
    {
        if (!double.IsFinite(dpi))
        {
            return "Không xác định";
        }

        if (dpi < 150)
        {
            return "Rất thấp";
        }

        if (dpi < 200)
        {
            return "Thấp";
        }

        if (dpi < 300)
        {
            return "Khá";
        }

        if (dpi < 400)
        {
            return "Tốt cho OCR";
        }

        return "Cao";
    }

    private static int NormalizePdfRotation(int rotation)
    {
        var normalized = rotation % 360;
        return normalized < 0 ? normalized + 360 : normalized;
    }

    private string GetActivePdfPath()
    {
        if (!string.IsNullOrWhiteSpace(_workingPdfPath) && File.Exists(_workingPdfPath))
        {
            return _workingPdfPath;
        }

        return _inputFilePath ?? throw new InvalidOperationException("Vui lòng chọn tệp PDF đầu vào.");
    }

    private static string CreateWorkingPdfPath()
    {
        return Path.Combine(Path.GetTempPath(), $"TachFilePdf_Edit_{Guid.NewGuid():N}.pdf");
    }

    private void SetWorkingPdf(string filePath)
    {
        DeleteWorkingPdfIfNeeded();
        _workingPdfPath = filePath;
        if (_activeMergeItem is not null)
        {
            _activeMergeItem.EditedFilePath = filePath;
        }
    }

    private void ResetWorkingPdf()
    {
        DeleteWorkingPdfIfNeeded();
        if (_activeMergeItem is not null)
        {
            _activeMergeItem.EditedFilePath = null;
        }

        _workingPdfPath = null;
    }

    private void DeleteWorkingPdfIfNeeded()
    {
        if (string.IsNullOrWhiteSpace(_workingPdfPath) || !File.Exists(_workingPdfPath))
        {
            return;
        }

        try
        {
            File.Delete(_workingPdfPath);
        }
        catch
        {
            // Temporary edit files are best-effort cleanup.
        }
    }

    private void ClampSplitRangesToPageCount()
    {
        if (_pageCount <= 0)
        {
            return;
        }

        foreach (var item in _ranges)
        {
            if (int.TryParse(item.FromPage, out var from))
            {
                item.FromPage = Math.Clamp(from, 1, _pageCount).ToString(CultureInfo.InvariantCulture);
            }

            if (int.TryParse(item.ToPage, out var to))
            {
                item.ToPage = Math.Clamp(to, 1, _pageCount).ToString(CultureInfo.InvariantCulture);
            }
        }
    }

    private static string BuildPageEditStatus(PageEditAction action, int pageNumber)
    {
        return action switch
        {
            PageEditAction.RotateLeft => $"Đã xoay trái trang {pageNumber}.",
            PageEditAction.RotateRight => $"Đã xoay phải trang {pageNumber}.",
            PageEditAction.Delete => $"Đã xóa trang {pageNumber}.",
            PageEditAction.Duplicate => $"Đã nhân bản trang {pageNumber}.",
            PageEditAction.MovePrevious => $"Đã chuyển trang {pageNumber} lên trước.",
            PageEditAction.MoveNext => $"Đã chuyển trang {pageNumber} xuống sau.",
            PageEditAction.MoveFirst => $"Đã đưa trang {pageNumber} lên đầu.",
            PageEditAction.MoveLast => $"Đã đưa trang {pageNumber} xuống cuối.",
            _ => "Đã chỉnh sửa PDF."
        };
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

    private static void SaveCoveredPdf(PdfDocument inputDocument, string outputFile, IReadOnlyList<CoverPatch> coverPatches)
    {
        using var outputDocument = new PdfDocument();
        for (var pageNumber = 1; pageNumber <= inputDocument.PageCount; pageNumber++)
        {
            var outputPage = outputDocument.AddPage(inputDocument.Pages[pageNumber - 1]);
            ApplyCoverPatches(outputPage, pageNumber, coverPatches);
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

    private static void ApplyCoverPatches(PdfPage page, int pageNumber, IReadOnlyList<CoverPatch> coverPatches)
    {
        var relevantPatches = coverPatches
            .Where(patch => patch.ApplyAllPages || patch.PageNumber == pageNumber)
            .ToList();
        if (relevantPatches.Count == 0)
        {
            return;
        }

        using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
        var pageWidth = page.Width.Point;
        var pageHeight = page.Height.Point;

        foreach (var patch in relevantPatches)
        {
            var opacity = Math.Clamp(patch.Opacity, 0, 1);
            var color = opacity >= 0.995
                ? XColor.FromArgb(patch.Color.R, patch.Color.G, patch.Color.B)
                : XColor.FromArgb((int)Math.Round(opacity * 255), patch.Color.R, patch.Color.G, patch.Color.B);
            var brush = new XSolidBrush(color);
            var x = patch.Area.Left * pageWidth;
            var y = patch.Area.Top * pageHeight;
            var width = Math.Max(1, (patch.Area.Right - patch.Area.Left) * pageWidth);
            var height = Math.Max(1, (patch.Area.Bottom - patch.Area.Top) * pageHeight);
            gfx.DrawRectangle(brush, new XRect(x, y, width, height));
        }
    }

    private void SyncCurrentPageFields()
    {
        var pageText = _currentPreviewPage.ToString(CultureInfo.InvariantCulture);
        PreviewPageTextBox.Text = pageText;
        CropPageTextBox.Text = pageText;
        UpdatePreviewPageHeader();
        RefreshCoverOverlayVisuals();
        UpdateCoverAreasInfo();
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

        RefreshCoverOverlayVisuals();
    }

    private void ApplyZoomForCurrentPreviewLayout()
    {
        if (_previewLayoutMode == PreviewLayoutMode.TextSelection)
        {
            _isFitWidthZoom = false;
            SetPreviewZoom(1);
            return;
        }

        if (_previewLayoutMode == PreviewLayoutMode.Vertical)
        {
            FitPreviewToWidth(allowZoomIn: false);
            return;
        }

        _isFitWidthZoom = false;
        SetPreviewZoom(_previewLayoutMode == PreviewLayoutMode.Compatibility ? 0.35 : MultiPagePreviewZoom);
    }

    private void ApplyPreviewLayout()
    {
        if (PdfPagesItemsControl is null)
        {
            return;
        }

        var isTextSelectionLayout = _previewLayoutMode == PreviewLayoutMode.TextSelection;
        if (PdfPagesScrollViewer is not null)
        {
            PdfPagesScrollViewer.Visibility = isTextSelectionLayout ? Visibility.Collapsed : Visibility.Visible;
        }

        if (TextSelectionPdfWebView is not null)
        {
            TextSelectionPdfWebView.Visibility = isTextSelectionLayout ? Visibility.Visible : Visibility.Collapsed;
        }

        var usesWrappingLayout = _previewLayoutMode is not PreviewLayoutMode.Vertical and not PreviewLayoutMode.TextSelection;
        var panelFactory = usesWrappingLayout
            ? new FrameworkElementFactory(typeof(WrapPanel))
            : new FrameworkElementFactory(typeof(StackPanel));
        if (usesWrappingLayout)
        {
            panelFactory.SetValue(WrapPanel.OrientationProperty, System.Windows.Controls.Orientation.Horizontal);
        }
        else
        {
            panelFactory.SetValue(StackPanel.OrientationProperty, System.Windows.Controls.Orientation.Vertical);
        }
        PdfPagesItemsControl.ItemsPanel = new ItemsPanelTemplate(panelFactory);

        if (PdfPagesScrollViewer is not null)
        {
            PdfPagesScrollViewer.HorizontalScrollBarVisibility = usesWrappingLayout
                ? ScrollBarVisibility.Disabled
                : ScrollBarVisibility.Auto;
        }

        if (PreviewLayoutButton is not null)
        {
            PreviewLayoutButton.ToolTip = _previewLayoutMode switch
            {
                PreviewLayoutMode.Grid => "Đang xem dạng lưới nhiều trang. Bấm để chuyển sang chế độ chọn/copy text.",
                PreviewLayoutMode.TextSelection => "Đang xem bằng WebView2 để chọn/copy text. Bấm để quay lại xem dọc.",
                PreviewLayoutMode.Compatibility => "Đang xem fallback tương thích PDF/A. Bấm để thử lại preview thật.",
                _ => "Đang xem dọc từng trang. Bấm để xem dạng lưới nhiều trang."
            };
        }

        UpdatePreviewHoverPreviewState();
    }

    private PdfPagePreviewItem CreatePagePreviewItem(int pageNumber, ImageSource image, bool isCompatibilityView = false)
    {
        return new PdfPagePreviewItem(
            pageNumber,
            image,
            isCompatibilityView,
            _previewLayoutMode == PreviewLayoutMode.Grid);
    }

    private void UpdatePreviewHoverPreviewState()
    {
        var showHoverPreview = _previewLayoutMode == PreviewLayoutMode.Grid;
        for (var index = 0; index < _pagePreviews.Count; index++)
        {
            if (_pagePreviews[index].ShowHoverPreview != showHoverPreview)
            {
                _pagePreviews[index] = _pagePreviews[index] with { ShowHoverPreview = showHoverPreview };
            }
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
        if (CropOverlayCanvas is null || EditOverlayCanvas is null || MainActionTabControl is null)
        {
            return;
        }

        CropOverlayCanvas.Visibility = MainActionTabControl.SelectedItem == CropTabItem &&
                                       _previewLayoutMode != PreviewLayoutMode.TextSelection
            ? Visibility.Visible
            : Visibility.Collapsed;

        EditOverlayCanvas.Visibility = MainActionTabControl.SelectedItem == EditTabItem &&
                                       _previewLayoutMode != PreviewLayoutMode.TextSelection
            ? Visibility.Visible
            : Visibility.Collapsed;

        RefreshCoverOverlayVisuals();
    }

    private void ClearCropSelectionVisual()
    {
        _cropArea = null;
        CropSelectionRectangle.Visibility = Visibility.Collapsed;
        CropSelectionRectangle.Width = 0;
        CropSelectionRectangle.Height = 0;
        CropSelectionInfoTextBlock.Text = "Chưa chọn vùng";
    }

    private void ClearCoverAreasVisual()
    {
        _coverPatches.Clear();
        EditSelectionRectangle.Visibility = Visibility.Collapsed;
        EditSelectionRectangle.Width = 0;
        EditSelectionRectangle.Height = 0;
        RefreshCoverOverlayVisuals();
        UpdateCoverAreasInfo();
    }

    private void RefreshCoverOverlayVisuals()
    {
        if (EditOverlayCanvas is null || EditSelectionRectangle is null)
        {
            return;
        }

        for (var i = EditOverlayCanvas.Children.Count - 1; i >= 0; i--)
        {
            if (EditOverlayCanvas.Children[i] is FrameworkElement { Tag: "CoverPatchVisual" })
            {
                EditOverlayCanvas.Children.RemoveAt(i);
            }
        }

        if (MainActionTabControl?.SelectedItem != EditTabItem)
        {
            return;
        }

        var pageBounds = GetCurrentPageImageBoundsOnOverlay(EditOverlayCanvas);
        if (pageBounds.IsEmpty)
        {
            return;
        }

        foreach (var patch in _coverPatches.Where(patch => patch.ApplyAllPages || patch.PageNumber == _currentPreviewPage))
        {
            var alpha = (byte)Math.Round(Math.Clamp(patch.Opacity, 0, 1.0) * 255);
            var previewColor = WpfColor.FromArgb(alpha, patch.Color.R, patch.Color.G, patch.Color.B);
            var rectangle = new System.Windows.Shapes.Rectangle
            {
                Tag = "CoverPatchVisual",
                Fill = new SolidColorBrush(previewColor),
                Stroke = new SolidColorBrush(patch.Color),
                StrokeThickness = 1.5,
                IsHitTestVisible = false,
                Width = Math.Max(1, patch.Area.WidthPercent / 100.0 * pageBounds.Width),
                Height = Math.Max(1, patch.Area.HeightPercent / 100.0 * pageBounds.Height)
            };

            Canvas.SetLeft(rectangle, pageBounds.Left + patch.Area.Left * pageBounds.Width);
            Canvas.SetTop(rectangle, pageBounds.Top + patch.Area.Top * pageBounds.Height);
            EditOverlayCanvas.Children.Insert(Math.Max(0, EditOverlayCanvas.Children.Count - 1), rectangle);
        }
    }

    private void UpdateCoverAreasInfo()
    {
        if (CoverAreasInfoTextBlock is null)
        {
            return;
        }

        if (_coverPatches.Count == 0)
        {
            CoverAreasInfoTextBlock.Text = "Chưa có vùng bù màu";
            return;
        }

        var allPagesCount = _coverPatches.Count(patch => patch.ApplyAllPages);
        var currentPageCount = _coverPatches.Count(patch => !patch.ApplyAllPages && patch.PageNumber == _currentPreviewPage);
        CoverAreasInfoTextBlock.Text = $"Tổng {_coverPatches.Count} vùng. Trang hiện tại: {currentPageCount}. Tất cả trang: {allPagesCount}.";
    }

    private Rect GetCurrentPageImageBoundsOnOverlay(Canvas overlayCanvas)
    {
        if (_currentPreviewPage < 1 ||
            PdfPagesItemsControl.ItemContainerGenerator.ContainerFromIndex(_currentPreviewPage - 1) is not FrameworkElement container)
        {
            return Rect.Empty;
        }

        var image = FindVisualChild<System.Windows.Controls.Image>(container);
        if (image is null || image.ActualWidth <= 0 || image.ActualHeight <= 0)
        {
            return Rect.Empty;
        }

        var topLeft = overlayCanvas.PointFromScreen(image.PointToScreen(new WpfPoint(0, 0)));
        var bottomRight = overlayCanvas.PointFromScreen(image.PointToScreen(new WpfPoint(image.ActualWidth, image.ActualHeight)));
        return new Rect(topLeft, bottomRight);
    }

    private static T? FindVisualChild<T>(DependencyObject parent)
        where T : DependencyObject
    {
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T typedChild)
            {
                return typedChild;
            }

            var descendant = FindVisualChild<T>(child);
            if (descendant is not null)
            {
                return descendant;
            }
        }

        return null;
    }

    private static WpfPoint ClampPointToRect(WpfPoint point, Rect rect)
    {
        return new WpfPoint(
            Math.Clamp(point.X, rect.Left, rect.Right),
            Math.Clamp(point.Y, rect.Top, rect.Bottom));
    }

    private void PickCoverColorAtMouse(MouseButtonEventArgs e)
    {
        var overlayPoint = e.GetPosition(EditOverlayCanvas);
        if (TryPickRenderedPreviewColor(overlayPoint, out var pickedColor))
        {
            SetCoverColor(pickedColor);
            _isPickingCoverColor = false;
            StatusTextBlock.Text = $"Đã lấy màu nền: #{pickedColor.R:X2}{pickedColor.G:X2}{pickedColor.B:X2}.";
            return;
        }

        var screenPoint = EditOverlayCanvas.PointToScreen(overlayPoint);
        using var bitmap = new System.Drawing.Bitmap(1, 1);
        using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.CopyFromScreen((int)Math.Round(screenPoint.X), (int)Math.Round(screenPoint.Y), 0, 0, new System.Drawing.Size(1, 1));
        }

        var color = bitmap.GetPixel(0, 0);
        SetCoverColor(WpfColor.FromRgb(color.R, color.G, color.B));
        _isPickingCoverColor = false;
        StatusTextBlock.Text = $"Đã lấy màu nền: #{color.R:X2}{color.G:X2}{color.B:X2}.";
    }

    private bool TryPickRenderedPreviewColor(WpfPoint overlayPoint, out WpfColor color)
    {
        color = default;

        var pageBounds = GetCurrentPageImageBoundsOnOverlay(EditOverlayCanvas);
        if (pageBounds.IsEmpty || !pageBounds.Contains(overlayPoint))
        {
            return false;
        }

        if (_currentPreviewPage < 1 ||
            _currentPreviewPage > _pagePreviews.Count ||
            _pagePreviews[_currentPreviewPage - 1].Image is not BitmapSource bitmapSource)
        {
            return false;
        }

        var xRatio = Math.Clamp((overlayPoint.X - pageBounds.Left) / pageBounds.Width, 0, 1);
        var yRatio = Math.Clamp((overlayPoint.Y - pageBounds.Top) / pageBounds.Height, 0, 1);
        var pixelX = Math.Clamp((int)Math.Round(xRatio * (bitmapSource.PixelWidth - 1)), 0, bitmapSource.PixelWidth - 1);
        var pixelY = Math.Clamp((int)Math.Round(yRatio * (bitmapSource.PixelHeight - 1)), 0, bitmapSource.PixelHeight - 1);

        BitmapSource source = bitmapSource.Format == PixelFormats.Bgra32 || bitmapSource.Format == PixelFormats.Pbgra32
            ? bitmapSource
            : new FormatConvertedBitmap(bitmapSource, PixelFormats.Bgra32, null, 0);

        var pixels = new byte[4];
        source.CopyPixels(new Int32Rect(pixelX, pixelY, 1, 1), pixels, 4, 0);
        color = WpfColor.FromRgb(pixels[2], pixels[1], pixels[0]);
        return true;
    }

    private void SetCoverColor(WpfColor color)
    {
        _coverColor = color;
        var brush = new SolidColorBrush(color);
        CoverColorPreviewBorder.Background = brush;
        CoverColorTextBlock.Text = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
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

public sealed record PdfPagePreviewItem(
    int PageNumber,
    ImageSource Image,
    bool IsCompatibilityView = false,
    bool ShowHoverPreview = false)
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

public readonly record struct PageCopyPlan(int SourceIndex, int RotateDelta = 0);

public readonly record struct PdfImageInfo(int PixelWidth, int PixelHeight);

public readonly record struct PageDpiResult(string StatusText, string Message);

public readonly record struct PdfStandardResult(string StatusText, string Message);

public enum PreviewLayoutMode
{
    Vertical,
    Grid,
    TextSelection,
    Compatibility
}

public enum PageEditAction
{
    RotateLeft,
    RotateRight,
    Delete,
    Duplicate,
    MovePrevious,
    MoveNext,
    MoveFirst,
    MoveLast
}

public sealed class MergePdfItem : INotifyPropertyChanged
{
    private int _index;
    private int _pageCount;
    private string? _editedFilePath;

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

    public required int PageCount
    {
        get => _pageCount;
        set
        {
            if (_pageCount == value)
            {
                return;
            }

            _pageCount = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(Detail));
        }
    }

    public required long FileSizeBytes { get; init; }

    public string? EditedFilePath
    {
        get => _editedFilePath;
        set
        {
            if (_editedFilePath == value)
            {
                return;
            }

            _editedFilePath = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(Detail));
            OnPropertyChanged(nameof(ActiveFilePath));
        }
    }

    public string FileName => Path.GetFileName(FilePath);

    public string ActiveFilePath => !string.IsNullOrWhiteSpace(EditedFilePath) && File.Exists(EditedFilePath)
        ? EditedFilePath
        : FilePath;

    public string Detail => EditedFilePath is null
        ? $"{PageCount} trang - {FormatFileSize(FileSizeBytes)} - {FilePath}"
        : $"{PageCount} trang - đã chỉnh sửa - {FilePath}";

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

public sealed record CoverPatch(
    int PageNumber,
    bool ApplyAllPages,
    CropArea Area,
    WpfColor Color,
    double Opacity);
