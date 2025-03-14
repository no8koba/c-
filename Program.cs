using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfPagesCount;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public class PDFProcessor
{
    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        PDFProcessor processor = new PDFProcessor();
        processor.SelectFolderAndProcessPDFs();
    }

    /// <summary>
    /// フォルダを選択し、PDFファイルを処理するメソッド
    /// </summary>
    private void SelectFolderAndProcessPDFs()
    {
        string folderPath;
        string saveFilePath = string.Empty;
        long fileCount = 0;
        DateTime startTime;
        DateTime endTime;
        string defaultFileName = DateTime.Now.ToString("yyyyMMddHHmmss") + "_pdf_files.csv";
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        try
        {
            // フォルダ選択ダイアログを作成
            var dlg = new FolderSelectDialog
            {
                Title = "pdfファイルが存在するフォルダを選択してください",
                Path = ""
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                folderPath = dlg.Path;
            }
            else
            {
                Console.WriteLine("フォルダが選択されませんでした。");
                return;
            }

            // ファイル保存ダイアログを作成
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.InitialDirectory = desktopPath;
                saveFileDialog.FileName = defaultFileName;
                saveFileDialog.Filter = "CSVファイル (*.csv)|*.csv";
                saveFileDialog.Title = "CSVファイルの保存場所を選択してください";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    saveFilePath = saveFileDialog.FileName;
                }
                else
                {
                    Console.WriteLine("ファイル保存場所が選択されませんでした。");
                    return;
                }
            }

            // CSVファイルの作成
            using (var writer = new StreamWriter(saveFilePath, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("ディレクトリパス,PDFファイル名,ページ数,ページ番号,用紙サイズ,カラー情報");

                // 処理開始時間を取得
                startTime = DateTime.Now;
                Console.WriteLine("処理開始時間: " + startTime.ToString("yyyy/MM/dd HH:mm:ss"));

                // PDFファイルを再帰的に検索する関数を呼び出す
                SearchPDFs(new DirectoryInfo(folderPath), writer, ref fileCount);

                // 処理終了時間を取得
                endTime = DateTime.Now;
                Console.WriteLine("処理終了時間: " + endTime.ToString("yyyy/MM/dd HH:mm:ss"));
            }

            Console.WriteLine($"PDFファイルの詳細が {saveFilePath} に保存されました。");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラーが発生しました: {ex.Message}");
        }
        finally
        {
            Console.WriteLine("終了するには何かキーを押してください...");
            Console.ReadKey();
        }
    }

    /// <summary>
    /// 指定されたフォルダ内のPDFファイルを再帰的に検索し、情報を取得するメソッド
    /// </summary>
    /// <param name="folder">検索するフォルダ</param>
    /// <param name="writer">CSVファイルに書き込むためのStreamWriter</param>
    /// <param name="fileCount">処理したファイルの数</param>
    private void SearchPDFs(DirectoryInfo folder, StreamWriter writer, ref long fileCount)
    {
        foreach (var file in folder.GetFiles("*.pdf"))
        {
            fileCount++;

            long pageCnt = -1;
            List<string> pageSizes = new List<string>();
            List<string> pageColor = new List<string>();

            // ページ数を取得
            pageCnt = GetPdfPageCount(file.FullName);
            if (pageCnt == -1)
            {
                // ページ数が取得できなかった場合はエラーメッセージを出力
                writer.WriteLine($"{file.DirectoryName},{file.Name},-1");
            }
            else
            {
                // ページ数が取得できた場合は用紙サイズとカラー情報を取得  
                try
                {
                    pageSizes = GetPdfPageSizes(file.FullName);
                    pageColor = GetPdfPageColor(file.FullName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"エラー: {file.FullName} の情報を取得できませんでした。詳細: {ex.Message}");
                }
                //  CSVファイルに書き込む
                for (int i = 0; i < pageSizes.Count; i++)
                {
                    writer.WriteLine($"{file.DirectoryName},\"{file.Name}\",1,{i + 1},{pageSizes[i]},{pageColor[i]}");
                    Console.WriteLine($"{file.DirectoryName},{file.Name},1,{i + 1},{pageSizes[i]},{pageColor[i]}");
                }

                // リストをクリアしてメモリを解放
                pageSizes.Clear();
                pageColor.Clear();

                // ガベージコレクションを強制的に実行
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        foreach (var subFolder in folder.GetDirectories())
        {
            SearchPDFs(subFolder, writer, ref fileCount);
        }
    }

    /// <summary>
    /// PDFファイルのページ数を取得するメソッド
    /// </summary>
    /// <param name="filePath">PDFファイルのパス</param>
    /// <returns>ページ数</returns>
    private long GetPdfPageCount(string filePath)
    {
        try
        {
            using (PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(filePath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
            {
                return document.PageCount;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラー: {filePath} のページ数を取得できませんでした。詳細: {ex.Message}");
            return -1; // エラーの場合は -1 を返す
        }
    }

    #region "用紙サイズ取得処理"

    /// <summary>
    /// PDFファイルの各ページの用紙サイズを取得するメソッド
    /// </summary>
    /// <param name="filePath">PDFファイルのパス</param>
    /// <returns>用紙サイズのリスト</returns>
    private static List<string> GetPdfPageSizes(string filePath)
    {
        var pageSizes = new List<string>();
        try
        {
            using (PdfDocument document = PdfReader.Open(filePath, PdfDocumentOpenMode.Import))
            {
                for (int i = 0; i < document.PageCount; i++)
                {
                    PdfPage page = document.Pages[i];

                    // 余白をトリミング
                    //TrimMargins(page);

                    //double width = page.TrimBox.Width;
                    //double height = page.TrimBox.Height;
                    double width = page.Width;
                    double height = page.Height;

                    // 用紙サイズを判定
                    string sizeDescription = GetPaperSizeDescription(width, height);
                    pageSizes.Add(sizeDescription);
                }
            }
        }
        catch (Exception ex)
        {
            throw new ApplicationException("用紙サイズの取得に失敗しました。", ex);
        }
        return pageSizes;
    }

    /// <summary>
    /// PDFページの余白をトリミングするメソッド
    /// </summary>
    /// <param name="page">PDFページ</param>
    private static void TrimMargins(PdfPage page)
    {
        // TrimBoxが定義されていない場合はMediaBoxを使用
        if (page.TrimBox.Width == 0 || page.TrimBox.Height == 0)
        {
            page.TrimBox = page.MediaBox;
        }

        // トリミングされたサイズをページのサイズに設定
        page.Width = XUnit.FromPoint(page.TrimBox.Width);
        page.Height = XUnit.FromPoint(page.TrimBox.Height);
    }

    /// <summary>
    /// 幅と高さに基づいて用紙サイズを判定するメソッド
    /// </summary>
    /// <param name="width">幅</param>
    /// <param name="height">高さ</param>
    /// <returns>用紙サイズの説明</returns>
    private static string GetPaperSizeDescription(double width, double height)
    {
        // 用紙サイズの定義
        (string Name, double Width, double Height, double AspectRatio)[] paperSizes = new (string, double, double, double)[11];
        paperSizes[0] = ("A0", 2384.65, 3370.79, 2384.65 / 3370.79);
        paperSizes[1] = ("A1", 1683.78, 2384.65, 1683.78 / 2384.65);
        paperSizes[2] = ("A2", 1190.55, 1683.78, 1190.55 / 1683.78);
        paperSizes[3] = ("A3", 841.89, 1190.55, 841.89 / 1190.55);
        paperSizes[4] = ("A4", 595.28, 841.89, 595.28 / 841.89);
        paperSizes[5] = ("A5", 419.53, 595.28, 419.53 / 595.28);
        paperSizes[6] = ("A6", 297.64, 419.53, 297.64 / 419.53);
        paperSizes[7] = ("A7", 209.76, 297.64, 209.76 / 297.64);
        paperSizes[8] = ("A8", 147.40, 209.76, 147.40 / 209.76);
        paperSizes[9] = ("A9", 104.88, 147.40, 104.88 / 147.40);
        paperSizes[10] = ("A10", 73.70, 104.88, 73.70 / 104.88);

        // 基準サイズの定義（A4を基準とする）
        double baseWidth = 595.28;
        double baseHeight = 841.89;

        // 動的な許容誤差の計算
        double tolerance = Math.Max(baseWidth, baseHeight) * 0.02; // 基準サイズの2%
        double ratioTolerance = 0.01; // 比率に対する許容誤差

        // 最も近いサイズを見つける
        double minDifference = double.MaxValue;
        string closestSize = "不明なサイズ";

        double aspectRatio = width / height;

        foreach ((string name, double sizeWidth, double sizeHeight, double sizeAspectRatio) in paperSizes)
        {
            double widthDifference = Math.Abs(width - sizeWidth);
            double heightDifference = Math.Abs(height - sizeHeight);

            double difference = widthDifference + heightDifference;
            double ratioDifference = Math.Abs(aspectRatio - sizeAspectRatio);

            if (difference < tolerance && ratioDifference < ratioTolerance && difference < minDifference)
            {
                minDifference = difference;
                closestSize = (width > height) ? $"{name}横" : $"{name}縦";
            }
        }

        return closestSize;
    }

    #endregion

    #region "メソッド：カラーモノクロ判定処理"
    /// <summary>
    /// PDFファイルの各ページのカラー情報を取得するメソッド
    /// </summary>
    /// <param name="filePath">PDFファイルのパス</param>
    /// <returns>各ページのカラー情報のリスト（"カラー" または "モノクロ"）</returns>
    private static List<string> GetPdfPageColor(string filePath)
    {
        List<string> pageColors = new List<string>();

        try
        {
            using (var document = PdfiumViewer.PdfDocument.Load(filePath))
            {
                for (int pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
                {
                    using (var image = RenderPageToImage(document, pageIndex))
                    {
                        if (IsColorImage(image))
                        {
                            pageColors.Add("カラー");
                        }
                        else
                        {
                            pageColors.Add("モノクロ");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new ApplicationException("カラー情報の取得に失敗しました。", ex);
        }

        return pageColors;
    }

    /// <summary>
    /// PDFのページを画像にレンダリングするメソッド
    /// </summary>
    /// <param name="document">PDFドキュメント</param>
    /// <param name="pageIndex">ページ番号</param>
    /// <returns>レンダリングされた画像</returns>
    private static Bitmap RenderPageToImage(PdfiumViewer.PdfDocument document, int pageIndex)
    {
        var page = document.Render(pageIndex, 300, 300, true);
        return new Bitmap(page);
    }

    /// <summary>
    /// 画像がカラーかモノクロかを判別するメソッド
    /// </summary>
    /// <param name="image">判別する画像</param>
    /// <returns>カラーの場合はtrue、モノクロの場合はfalse</returns>
    private static bool IsColorImage(Bitmap image)
    {
        for (int y = 0; y < image.Height; y++)
        {
            for (int x = 0; x < image.Width; x++)
            {
                Color pixelColor = image.GetPixel(x, y);
                if (pixelColor.R != pixelColor.G || pixelColor.G != pixelColor.B)
                {
                    return true;
                }
            }
        }
        return false;
    }
    #endregion

}
