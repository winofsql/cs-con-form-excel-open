using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace cs_con_form_excel_open
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            OpenFileDialog obj = new OpenFileDialog();
            obj.Filter = "Excel(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*";
            obj.InitialDirectory = @"C:\";

            if (obj.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            // Excel アプリケーション
            dynamic excelApp =
                Activator
                    .CreateInstance(Type
                        .GetTypeFromProgID("Excel.Application"));

            // Excel のパス
            string path = obj.FileName;

            // Excel ブック( 既存 )
            dynamic workBook = excelApp.Workbooks.Open(path);

            // Excel を表示( 完成したらコメント化 )
            excelApp.Visible = true;

            // 警告を出さない
            excelApp.DisplayAlerts = false;

            // シート数
            int Count = workBook.Sheets.Count;

            // 最後のシートの後にシートを追加
            workBook.Sheets.Add(After: workBook.Sheets(Count));
            workBook.Sheets(Count + 1).Name = $"追加のシート {Count}";

            // 追加シートをアクティブにする
            workBook.Sheets(Count + 1).Activate();

            // セルに値をセット
            workBook.Sheets(Count + 1).Cells(1, 1).Value = "社員名";
            workBook.Sheets(Count + 1).Cells(2, 1).Value = "山田　太郎甚左衛門";
            workBook.Sheets(Count + 1).Cells(3, 1).Value = "鈴木　一郎";
            workBook.Sheets(Count + 1).Cells(4, 1).Value = "佐藤　洋子";

            // 列幅自動調整
            workBook.Sheets(Count + 1).Columns("A:A").EntireColumn.AutoFit();

            // さらに追加
            workBook.Sheets.Add(After: workBook.Sheets(Count + 1));

            // Shreet 参照で処理
            dynamic Sheet = workBook.Sheets(Count + 2);

            // ****************************
            // セルに値を直接セット
            // ****************************
            for (int i = 1; i <= 10; i++)
            {
                Sheet.Cells(i, 1).Value = "処理 : " + i;
            }

            // ****************************
            // 1つのセルから
            // AutoFill で値をセット
            // ****************************
            Sheet.Cells(1, 2).Value = "子";

            // 基となるセル範囲
            dynamic SourceRange =
                Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(1, 2));

            // オートフィルの範囲(基となるセル範囲を含む )
            dynamic FillRange =
                Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(10, 2));
            SourceRange.AutoFill (FillRange);

            // 保存
            workBook.Save();

            // 閉じる
            workBook.Close();

            // 終了
            excelApp.Quit();

            // 解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject (excelApp);

            // C# ではほぼ完全解放無理なので強制終了させる
            foreach (var p in Process.GetProcessesByName("EXCEL"))
            {
                if (p.MainWindowTitle == "")
                {
                    p.Kill();
                }
            }

            // ファイルの種類によってアプリケーションを起動する
            ProcessStartInfo processStartInfo = new ProcessStartInfo("RunDLL32.EXE", $"url.dll,FileProtocolHandler {path}" );
            Process.Start(processStartInfo);
        }
    }
}
