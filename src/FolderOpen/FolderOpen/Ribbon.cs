using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Diagnostics;

// TODO:  リボン (XML) アイテムを有効にするには、次の手順に従います。

// 1: 次のコード ブロックを ThisAddin、ThisWorkbook、ThisDocument のいずれかのクラスにコピーします。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. ボタンのクリックなど、ユーザーの操作を処理するためのコールバック メソッドを、このクラスの
//    "リボンのコールバック" 領域に作成します。メモ: このリボンがリボン デザイナーからエクスポートされたものである場合は、
//    イベント ハンドラー内のコードをコールバック メソッドに移動し、リボン拡張機能 (RibbonX) のプログラミング モデルで
//    動作するように、コードを変更します。

// 3. リボン XML ファイルのコントロール タグに、コードで適切なコールバック メソッドを識別するための属性を割り当てます。  

// 詳細については、Visual Studio Tools for Office ヘルプにあるリボン XML のドキュメントを参照してください。


namespace FolderOpen
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("FolderOpen.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック
        //ここでコールバック メソッドを作成します。コールバック メソッドの追加について詳しくは https://go.microsoft.com/fwlink/?LinkID=271226 をご覧ください

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// 選択したテキストから動作
        /// </summary>
        /// <param name="control"></param>
        public void SelectedPath(Office.IRibbonControl control)
        {
            var selStr = GetSelectedText();
            if (selStr != "")
            {
                selStr = PathMake(selStr);
                OpenFolderFromSelected(selStr);
            }
        }

        /// <summary>
        /// 選択したハイパーリンクから動作
        /// </summary>
        /// <param name="control"></param>
        public void SelectedHyperLinkPath(Office.IRibbonControl control)
        {
            var selStr = GetSelectedHyperLinkText();
            if (selStr != "")
            {
                selStr = PathMake(selStr);
                OpenFolderFromSelected(selStr);
            }
        }

        /// <summary>
        /// 選択したテキストから動作（ファイルを開く）
        /// </summary>
        /// <param name="control"></param>
        public void SelectedFileOpen(Office.IRibbonControl control)
        {
            var selStr = GetSelectedText();
            if (selStr != "")
            {
                selStr = PathMake(selStr);
                OpenFileFromSelected(selStr);
            }
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        #region 公開サービス

        /// <summary>
        /// テキストを取得します
        /// </summary>
        /// <returns></returns>
        private string GetSelectedText()
        {
            Outlook.Inspector inspector;
            string selStr;
            var windowType = Globals.ThisAddIn.Application.ActiveWindow();

            switch (windowType)
            {
                // メール画面 or 閲覧ウィンドウ を判定しインスペクタを取得します
                case Outlook.Inspector _:
                    inspector = Globals.ThisAddIn.Application.ActiveInspector();
                    break;
                case Outlook.Explorer _:
                    inspector = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1].GetInspector();
                    break;
                default:
                    selStr = "";
                    return selStr;
            }

            // テキストを取得します
            Word.Selection selection = inspector.WordEditor.Parent.Selection;
            selStr = selection.Text;

            return selStr;
        }

        /// <summary>
        /// ハイパーリンクを取得します
        /// </summary>
        /// <returns></returns>
        private string GetSelectedHyperLinkText()
        {
            Outlook.Inspector inspector;
            string selStr;
            var windowType = Globals.ThisAddIn.Application.ActiveWindow();

            switch (windowType)
            {
                // メール画面 or 閲覧ウィンドウ を判定しインスペクタを取得します
                case Outlook.Inspector _:
                    inspector = Globals.ThisAddIn.Application.ActiveInspector();
                    break;
                case Outlook.Explorer _:
                    inspector = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1].GetInspector();
                    break;
                default:
                    selStr = "";
                    return selStr;
            }

            // ハイパーリンクを取得します
            selStr = inspector.WordEditor.Parent.Selection.Hyperlinks.item[1].Address;

            return selStr;
        }

        /// <summary>
        /// 選択したパスよりフォルダを開きます
        /// </summary>
        /// <param name="filePath"></param>
        private void OpenFolderFromSelected(string filePath)
        {

            if (!File.Exists(filePath))
            {
                if (!Directory.Exists(filePath))
                {
                    return;
                }
            }

            var fileInfo = new FileInfo(filePath);
            var folderPath = fileInfo.DirectoryName;

            if (Directory.Exists(filePath) || Directory.Exists(folderPath))
            {
                if ((File.GetAttributes(filePath) & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    Process.Start(filePath);
                }
                else
                {
                    Process.Start(folderPath);
                }
            }
            else if (Directory.Exists(SpaceDelete(filePath)) || Directory.Exists(SpaceDelete(folderPath)))
            {
                if ((File.GetAttributes(SpaceDelete(filePath)) & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    Process.Start(SpaceDelete(filePath));
                }
                else
                {
                    Process.Start(SpaceDelete(folderPath));
                }
            }
            else
            {
                MessageBox.Show("フォルダパスが正しく選択されていない、もしくはフォルダが存在しません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 選択したパスよりファイルを開きます
        /// </summary>
        /// <param name="selStr"></param>
        private void OpenFileFromSelected(string selStr)
        {
            if (File.Exists(selStr))
            {
                Process.Start(selStr);
            }
            else
            {
                if (File.Exists(SpaceDelete(selStr)))
                {
                    Process.Start(SpaceDelete(selStr));
                }
                else
                {
                    MessageBox.Show("ファイルパスが正しく選択されていない、もしくはファイルが存在しません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        /// <summary>
        /// パスをエクスプローラーで開ける形に加工します
        /// </summary>
        /// <param name="selStr"></param>
        /// <returns></returns>
        private string PathMake(string selStr)
        {

            // フォルダ名に使用できない文字群の定義（/\:は区切りに使用するため除外）
            selStr = System.Text.RegularExpressions.Regex.Replace(selStr, @"[*?|<>\r\n\0\t\b\f\v]", string.Empty);
            selStr = System.Text.RegularExpressions.Regex.Replace(selStr, "\"", string.Empty);

            // 先頭にスペースおよび＜＞が存在すればパスに関係ないとして削除します
            int iCnt;
            for (iCnt = 0; iCnt < selStr.Length - 1; iCnt++)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(selStr.Substring(0, 1), @"[\s＜＞]"))
                {
                    selStr = selStr.Remove(0, 1);
                }
                else
                {
                    break;
                }
            }

            // 末尾にスペースおよび＜＞が存在すればパスに関係ないとして削除
            for (iCnt = selStr.Length - 1; iCnt > 0; iCnt--)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(selStr.Substring(iCnt, 1), @"[\s＜＞]"))
                {
                    selStr = selStr.Remove(iCnt, 1);
                }
                else
                {
                    break;
                }
            }

            // \\～　Y:～　で開始していれば変更せず返します
            string retStr;
            if (System.Text.RegularExpressions.Regex.IsMatch(selStr, @"^\\\\|^[A-Z]:"))
            {
                retStr = selStr;
            }
            else
            {
                // fileから始まっていればfileを削除します
                if (System.Text.RegularExpressions.Regex.IsMatch(selStr, "^(file:)"))
                {
                    selStr = selStr.Remove(0, 5);
                }

                // fileの後ろが//\\であれば//を削除します
                if (System.Text.RegularExpressions.Regex.IsMatch(selStr, @"^(//\\\\)"))
                {
                    selStr = System.Text.RegularExpressions.Regex.Replace(selStr, "^(//)", string.Empty);
                }
                // fileの後ろが\\//であれば//を削除します
                else if (System.Text.RegularExpressions.Regex.IsMatch(selStr, @"^(\\\\//)"))
                {
                    selStr = System.Text.RegularExpressions.Regex.Replace(selStr, @"^(\\\\//)", "\\\\");
                }
                // fileの後ろが\\Y:等であれば\\を削除します
                else if (System.Text.RegularExpressions.Regex.IsMatch(selStr, @"^(\\\\[A-Z]:)"))
                {
                    selStr = System.Text.RegularExpressions.Regex.Replace(selStr, @"^\\\\", string.Empty);
                }
                // fileの後ろが//Y:等であれば//を削除します
                else if (System.Text.RegularExpressions.Regex.IsMatch(selStr, @"^(//[A-Z]:)"))
                {
                    selStr = System.Text.RegularExpressions.Regex.Replace(selStr, @"^//", string.Empty);
                }
                // fileの後ろが//であれば削除します
                else if (System.Text.RegularExpressions.Regex.IsMatch(selStr, "^(//)"))
                {
                    //selStr = System.Text.RegularExpressions.Regex.Replace(selStr, "^(//)", "\\\\");
                    selStr = System.Text.RegularExpressions.Regex.Replace(selStr, "^(//)", string.Empty);
                }

                // 上記の加工後、Y:等で始まっていればそのまま返します
                if (System.Text.RegularExpressions.Regex.IsMatch(selStr, "^[A-Z]:"))
                {
                    retStr = selStr;
                }
                // 上記の加工後、先頭に\\がなく、dcinc～のようであれば\\を先頭に追加します
                else if (System.Text.RegularExpressions.Regex.IsMatch(selStr, @"^[^(\\\\)]"))
                {
                    selStr = @"\\" + selStr;
                    retStr = selStr;
                }
                // 先頭に\\がついていればそのまま返します
                else
                {
                    retStr = selStr;
                }
            }

            // /を\に置換します
            if (System.Text.RegularExpressions.Regex.IsMatch(selStr, "/"))
            {
                selStr = System.Text.RegularExpressions.Regex.Replace(selStr, "/", "\\");
                retStr = selStr;
            }


            return retStr;
        }

        /// <summary>
        /// 半角スペースと全角スペースを除外します
        /// </summary>
        /// <param name="selStr"></param>
        /// <returns></returns>
        private string SpaceDelete(string selStr)
        {
            var retStr = System.Text.RegularExpressions.Regex.Replace(selStr, @"\s", string.Empty);
            return retStr;
        }

        #endregion
    }
}
