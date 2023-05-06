using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.Intrinsics.Arm;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Text_Grab.Controls;
using Text_Grab.Models;
using Text_Grab.Properties;
using Text_Grab.Views;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;
using BitmapDecoder = Windows.Graphics.Imaging.BitmapDecoder;
using Point = System.Windows.Point;

namespace Text_Grab.Utilities;

public static class OcrExtensions
{

    public static void GetTextFromOcrLine(this OcrLine ocrLine, bool isSpaceJoiningOCRLang, StringBuilder text)
    {
        // (when OCR language is zh or ja)
        // matches words in a space-joining language, which contains:
        // - one letter that is not in "other letters" (CJK characters are "other letters")
        // - one number digit
        // - any words longer than one character
        // Chinese and Japanese characters are single-character words
        // when a word is one punctuation/symbol, join it without spaces

        if (isSpaceJoiningOCRLang)
        {
            text.AppendLine(ocrLine.Text);

            if (Settings.Default.CorrectErrors)
                text.TryFixEveryWordLetterNumberErrors();
        }
        else
        {
            bool isFirstWord = true;
            bool isPrevWordSpaceJoining = false;

            Regex regexSpaceJoiningWord = new(@"(^[\p{L}-[\p{Lo}]]|\p{Nd}$)|.{2,}");

            foreach (OcrWord ocrWord in ocrLine.Words)
            {
                string wordString = ocrWord.Text;

                bool isThisWordSpaceJoining = regexSpaceJoiningWord.IsMatch(wordString);

                if (Settings.Default.CorrectErrors)
                    wordString = wordString.TryFixNumberLetterErrors();

                if (isFirstWord || (!isThisWordSpaceJoining && !isPrevWordSpaceJoining))
                    _ = text.Append(wordString);
                else
                    _ = text.Append(' ').Append(wordString);

                isFirstWord = false;
                isPrevWordSpaceJoining = isThisWordSpaceJoining;
            }
        }

        if (Settings.Default.CorrectToLatin)
            text.ReplaceGreekOrCyrillicWithLatin();
    }

    public static void RemoveTrailingNewlines(this StringBuilder text)
    {
        while (text.Length > 0 && (text[^1] == '\n' || text[^1] == '\r'))
            text.Length--;
    }

    public static async Task<string> GetTextFromAbsoluteRectAsync(Rect rect, Language language)
    {
        Rectangle selectedRegion = ShapeExtensions.RectangleFromRect(rect);
        Bitmap bmp = ImageMethods.GetRegionOfScreenAsBitmap(selectedRegion);

        return GetStringFromOcrOutputs(await GetTextFromImageAsync(bmp, language));
    }

    /// <summary>
    /// 用于在指定窗口上截取一个矩形区域，将其作为一张图片进行 OCR 文本识别<br/>
    /// 调用 passedWindow.GetAbsolutePosition() 方法获取传递进来的窗口 passedWindow 在屏幕上的绝对位置。<br/>
    ///根据 selectedRegion 确定被选择的区域所在的信息，并计算相应的坐标，得到经过偏移矫正后的 correctedRegion。<br/>
    ///使用 ImageMethods.GetRegionOfScreenAsBitmap 方法截取经过偏移矫正后的正确位置的位图 bmp。<br/>
    ///调用 GetTextFromImageAsync 方法，对 bmp 进行 OCR 文本识别，并获取文本识别引擎返回的异步任务结果。<br/>
    ///调用 GetStringFromOcrOutputs 方法，将 OcrOutput List 转换为 OCR 识别的字符串结果，并返回该字符串。<br/>
    /// </summary>
    /// <param name="passedWindow"></param>
    /// <param name="selectedRegion"></param>
    /// <param name="language"></param>
    /// <returns></returns>
    public static async Task<string> GetRegionsTextAsync(Window passedWindow, Rectangle selectedRegion, Language language)
    {
        Point absPosPoint = passedWindow.GetAbsolutePosition();

        int thisCorrectedLeft = (int)absPosPoint.X + selectedRegion.Left;
        int thisCorrectedTop = (int)absPosPoint.Y + selectedRegion.Top;

        Rectangle correctedRegion = new(thisCorrectedLeft, thisCorrectedTop, selectedRegion.Width, selectedRegion.Height);
        Bitmap bmp = ImageMethods.GetRegionOfScreenAsBitmap(correctedRegion);

        return GetStringFromOcrOutputs( await GetTextFromImageAsync(bmp, language));
    }

    /// <summary>
    /// 用于在指定窗口上截取一个矩形区域，将其作为一张图片进行 OCR 文本识别，并将结果按表格形式输出<br/>
    /// 调用 passedWindow.GetAbsolutePosition() 方法获取传递进来的窗口 passedWindow 在屏幕上的绝对位置。<br/>
    ///根据 selectedRegion 确定被选择的区域所在的信息，并计算相应的坐标，得到经过偏移矫正后的 correctedRegion。<br/>
    ///使用 ImageMethods.GetRegionOfScreenAsBitmap 方法截取经过偏移矫正后的正确位置的位图 bmp。<br/>
    ///使用 GetIdealScaleFactorForOcrAsync 方法获取理想的 OCR 识别缩放比例 scale。<br/>
    ///使用 ImageMethods.ScaleBitmapUniform 方法对位图 bmp 进行等比例缩放，得到高清晰度的 scaledBitmap。<br/>
    ///调用 GetOcrResultFromImageAsync 方法，对 scaledBitmap 进行 OCR 文本识别，并获取文本识别引擎返回的 OcrResult 结果。<br/>
    ///使用 VisualTreeHelper.GetDpi 获取 dpiScale（屏幕 DPI 缩放比例）。<br/>
    ///调用 ResultTable.ParseOcrResultIntoWordBorders 方法，将 OcrResult 中所有单词的边框信息解析为 WordBorder 对象，并将 dpiScale 作为参数传入此方法中。<br/>
    ///调用 ResultTable.GetWordsAsTable 方法，将 WordBorder 的 List 转换为表格形式的字符串，并返回该字符串。<br/>
    /// </summary>
    /// <param name="passedWindow"></param>
    /// <param name="selectedRegion"></param>
    /// <param name="language"></param>
    /// <returns></returns>
    public static async Task<string> GetRegionsTextAsTableAsync(Window passedWindow, Rectangle selectedRegion, Language language)
    {
        Point absPosPoint = passedWindow.GetAbsolutePosition();

        int thisCorrectedLeft = (int)absPosPoint.X + selectedRegion.Left;
        int thisCorrectedTop = (int)absPosPoint.Y + selectedRegion.Top;

        Rectangle correctedRegion = new(thisCorrectedLeft, thisCorrectedTop, selectedRegion.Width, selectedRegion.Height);
        Bitmap bmp = ImageMethods.GetRegionOfScreenAsBitmap(correctedRegion);
        double scale = await GetIdealScaleFactorForOcrAsync(bmp, language);
        using Bitmap scaledBitmap = ImageMethods.ScaleBitmapUniform(bmp, scale);
        DpiScale dpiScale = VisualTreeHelper.GetDpi(passedWindow);
        OcrResult ocrResult = await GetOcrResultFromImageAsync(scaledBitmap, language);
        List<WordBorder> wordBorders = ResultTable.ParseOcrResultIntoWordBorders(ocrResult, dpiScale);
        return ResultTable.GetWordsAsTable(wordBorders, dpiScale, LanguageUtilities.IsLanguageSpaceJoining(language));
    }

    public static async Task<(OcrResult, double)> GetOcrResultFromRegionAsync(Rectangle region, Language language)
    {
        Bitmap bmp = ImageMethods.GetRegionOfScreenAsBitmap(region);

        double scale = await GetIdealScaleFactorForOcrAsync(bmp, language);
        using Bitmap scaledBitmap = ImageMethods.ScaleBitmapUniform(bmp, scale);

        OcrResult ocrResult = await GetOcrResultFromImageAsync(scaledBitmap, language);

        return (ocrResult, scale);
    }

    public async static Task<OcrResult> GetOcrFromStreamAsync(MemoryStream memoryStream, Language language)
    {
        using WrappingStream wrapper = new(memoryStream);
        wrapper.Position = 0;
        BitmapDecoder bmpDecoder = await BitmapDecoder.CreateAsync(wrapper.AsRandomAccessStream());
        using SoftwareBitmap softwareBmp = await bmpDecoder.GetSoftwareBitmapAsync();

        await wrapper.FlushAsync();

        return await GetOcrResultFromImageAsync(softwareBmp, language);
    }

    public async static Task<OcrResult> GetOcrFromStreamAsync(IRandomAccessStream stream, Language language)
    {
        BitmapDecoder bmpDecoder = await BitmapDecoder.CreateAsync(stream);
        using SoftwareBitmap softwareBmp = await bmpDecoder.GetSoftwareBitmapAsync();

        return await GetOcrResultFromImageAsync(softwareBmp, language);
    }

    public async static Task<OcrResult> GetOcrResultFromImageAsync(BitmapImage scaledBitmap, Language language)
    {
        Bitmap bitmap = ImageMethods.BitmapImageToBitmap(scaledBitmap);
        return await GetOcrResultFromImageAsync(bitmap, language);
    }

    public async static Task<OcrResult> GetOcrResultFromImageAsync(SoftwareBitmap scaledBitmap, Language language)
    {
        OcrEngine ocrEngine = OcrEngine.TryCreateFromLanguage(language);
        return await ocrEngine.RecognizeAsync(scaledBitmap);
    }

    public async static Task<OcrResult> GetOcrResultFromImageAsync(Bitmap scaledBitmap, Language language)
    {
        await using MemoryStream memory = new();
        using WrappingStream wrapper = new(memory);

        scaledBitmap.Save(wrapper, ImageFormat.Bmp);
        wrapper.Position = 0;
        BitmapDecoder bmpDecoder = await BitmapDecoder.CreateAsync(wrapper.AsRandomAccessStream());
        using SoftwareBitmap softwareBmp = await bmpDecoder.GetSoftwareBitmapAsync();
        await wrapper.FlushAsync();


        return await GetOcrResultFromImageAsync(softwareBmp, language);
    }

    public async static Task<List<OcrOutput>> GetTextFromRandomAccessStream(IRandomAccessStream randomAccessStream, Language language)
    {
        OcrResult ocrResult = await GetOcrFromStreamAsync(randomAccessStream, language);

        List<OcrOutput> outputs = new();

        OcrOutput paragraphsOutput = GetTextFromOcrResult(language, null, ocrResult);

        outputs.Add(paragraphsOutput);

        if (Settings.Default.TryToReadBarcodes)
        {
            Bitmap bitmap = ImageMethods.GetBitmapFromIRandomAccessStream(randomAccessStream);
            OcrOutput barcodeResult = BarcodeUtilities.TryToReadBarcodes(bitmap);
            outputs.Add(barcodeResult);
        }

        return outputs;
    }

    public static Task<List<OcrOutput>> GetTextFromImageAsync(SoftwareBitmap softwareBitmap, Language language)
    {
        throw new NotImplementedException();

        // TODO:    scale software bitmaps
        //          Store software bitmaps on OcrOutput
        //          Read QR Codes from software bitmaps
    }

    public async static Task<List<OcrOutput>> GetTextFromImageAsync(BitmapImage bitmapImage, Language language)
    {
        Bitmap bitmap = ImageMethods.BitmapImageToBitmap(bitmapImage);
        return await GetTextFromImageAsync(bitmap, language);
    }

    public static Task<List<OcrOutput>> GetTextFromStreamAsync(MemoryStream stream, Language language)
    {
        throw new NotImplementedException();
    }

    public static Task<List<OcrOutput>> GetTextFromStreamAsync(IRandomAccessStream stream, Language language)
    {
        throw new NotImplementedException();
    }

    public async static Task<List<OcrOutput>> GetTextFromImageAsync(Bitmap bitmap, Language language)
    {
        List<OcrOutput> outputs = new();

        double scale = await GetIdealScaleFactorForOcrAsync(bitmap, language);
        Bitmap scaledBitmap = ImageMethods.ScaleBitmapUniform(bitmap, scale);

        if (Settings.Default.UseTesseract)
        {
            OcrOutput tesseractOutput = await TesseractHelper.GetOcrOutputFromBitmap(scaledBitmap, language);
            outputs.Add(tesseractOutput);
        }
        else
        {
            OcrResult ocrResult = await OcrExtensions.GetOcrResultFromImageAsync(scaledBitmap, language);
            OcrOutput paragraphsOutput = GetTextFromOcrResult(language, scaledBitmap, ocrResult);
            outputs.Add(paragraphsOutput);
        }

        if (Settings.Default.TryToReadBarcodes)
        {
            OcrOutput barcodeResult = BarcodeUtilities.TryToReadBarcodes(scaledBitmap);
            outputs.Add(barcodeResult);
        }

        return outputs;
    }

    private static OcrOutput GetTextFromOcrResult(Language language, Bitmap? scaledBitmap, OcrResult ocrResult)
    {
        StringBuilder text = new();

        bool isSpaceJoiningOCRLang = LanguageUtilities.IsLanguageSpaceJoining(language);

        foreach (OcrLine ocrLine in ocrResult.Lines)
            ocrLine.GetTextFromOcrLine(isSpaceJoiningOCRLang, text);

        if (LanguageUtilities.IsLanguageRightToLeft(language))
            text.ReverseWordsForRightToLeft();

        OcrOutput paragraphsOutput = new()
        {
            Kind = OcrOutputKind.Paragraph,
            RawOutput = text.ToString(),
            Language = language,
            SourceBitmap = scaledBitmap,
        };
        return paragraphsOutput;
    }

    public static string GetStringFromOcrOutputs(List<OcrOutput> outputs)
    {
        StringBuilder text = new();

        foreach (OcrOutput output in outputs)
        {
            output.CleanOutput();

            if (!string.IsNullOrWhiteSpace(output.CleanedOutput))
                text.Append(output.CleanedOutput);
            else if (!string.IsNullOrWhiteSpace(output.RawOutput))
                text.Append(output.RawOutput);
        }

        return text.ToString();
    }

    public static async Task<string> OcrAbsoluteFilePathAsync(string absolutePath)
    {
        Uri fileURI = new(absolutePath, UriKind.Absolute);
        BitmapImage droppedImage = new(fileURI);
        droppedImage.Freeze();
        Bitmap bmp = ImageMethods.BitmapImageToBitmap(droppedImage);
        Language language = LanguageUtilities.GetOCRLanguage();
        return GetStringFromOcrOutputs(await GetTextFromImageAsync(bmp, language));
    }
    /// <summary>
    /// GetClickedWordAsync 是一个用于将在指定窗口上指定位置的单个词语截屏并进行 OCR 文本识别的工具方法。它通常在用户截取选区时，当选取区域比较小（不足以拥有多行文字）时，调用此方法会使得整个流程变得更加简化和快速
    /// </summary>
    /// <param name="passedWindow">调用 ImageMethods.GetWindowsBoundsBitmap 方法，获取传递进来的窗口 passedWindow 对应的桌面区域的位图</param>
    /// <param name="clickedPoint">将鼠标点击的坐标 clickedPoint 作为参数，和刚才获取的位图一起传给 GetTextFromClickedWordAsync 方法，调用 OCR 引擎进行文本识别</param>
    /// <param name="OcrLang">返回 OCR 引擎返回的字符串，并去除首尾空格</param>
    /// <returns></returns>
    public static async Task<string> GetClickedWordAsync(Window passedWindow, Point clickedPoint, Language OcrLang)
    {
        using Bitmap bmp = ImageMethods.GetWindowsBoundsBitmap(passedWindow);
        string ocrText = await GetTextFromClickedWordAsync(clickedPoint, bmp, OcrLang);
        return ocrText.Trim();
    }

    private static async Task<string> GetTextFromClickedWordAsync(Point singlePoint, Bitmap bitmap, Language language)
    {
        return GetTextFromClickedWord(singlePoint, await OcrExtensions.GetOcrResultFromImageAsync(bitmap, language));
    }

    private static string GetTextFromClickedWord(Point singlePoint, OcrResult ocrResult)
    {
        Windows.Foundation.Point fPoint = new(singlePoint.X, singlePoint.Y);

        foreach (OcrLine ocrLine in ocrResult.Lines)
            foreach (OcrWord ocrWord in ocrLine.Words)
                if (ocrWord.BoundingRect.Contains(fPoint))
                    return ocrWord.Text;

        return string.Empty;
    }

    public async static Task<double> GetIdealScaleFactorForOcrAsync(SoftwareBitmap bitmap, Language selectedLanguage)
    {
        OcrResult ocrResult = await OcrExtensions.GetOcrResultFromImageAsync(bitmap, selectedLanguage);

        return GetIdealScaleFactorForOcrResult(ocrResult, bitmap.PixelHeight, bitmap.PixelWidth);
    }

    public async static Task<double> GetIdealScaleFactorForOcrAsync(Bitmap bitmap, Language selectedLanguage)
    {
        OcrResult ocrResult = await OcrExtensions.GetOcrResultFromImageAsync(bitmap, selectedLanguage);

        return GetIdealScaleFactorForOcrResult(ocrResult, bitmap.Height, bitmap.Width);
    }

    private static double GetIdealScaleFactorForOcrResult(OcrResult ocrResult, int height, int width)
    {
        List<double> heightsList = new();
        double scaleFactor = 1.5;

        foreach (OcrLine ocrLine in ocrResult.Lines)
            foreach (OcrWord ocrWord in ocrLine.Words)
                heightsList.Add(ocrWord.BoundingRect.Height);

        double lineHeight = 10;

        if (heightsList.Count > 0)
            lineHeight = heightsList.Average();

        // Ideal Line Height is 40px
        const double idealLineHeight = 40.0;

        scaleFactor = idealLineHeight / lineHeight;

        if (width * scaleFactor > OcrEngine.MaxImageDimension || height * scaleFactor > OcrEngine.MaxImageDimension)
        {
            int largerDim = Math.Max(width, height);
            // find the largest possible scale factor, because the ideal scale factor is too high

            scaleFactor = OcrEngine.MaxImageDimension / largerDim;
        }

        return scaleFactor;
    }

    public static Rect GetBoundingRect(this OcrLine ocrLine)
    {
        double top = ocrLine.Words.Select(x => x.BoundingRect.Top).Min();
        double bottom = ocrLine.Words.Select(x => x.BoundingRect.Bottom).Max();
        double left = ocrLine.Words.Select(x => x.BoundingRect.Left).Min();
        double right = ocrLine.Words.Select(x => x.BoundingRect.Right).Max();

        return new()
        {
            X = left,
            Y = top,
            Width = Math.Abs(right - left),
            Height = Math.Abs(bottom - top)
        };
    }
}
