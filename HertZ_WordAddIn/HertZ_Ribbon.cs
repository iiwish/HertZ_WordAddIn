using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace HertZ_WordAddIn
{
    public partial class HertZ_Ribbon
    {
        private Word.Application WordApp;
        private Word.Document WordDoc;

        private void HertZ_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// 设置文本样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextFormat_Click(object sender, RibbonControlEventArgs e)
        {
            WordApp = Globals.ThisAddIn.Application;
            WordDoc = (Word.Document)WordApp.ActiveDocument;


        }

        /// <summary>
        /// 设置默认格式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SettingFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Form SettingFormatForm = new SettingFormatForm();
            SettingFormatForm.StartPosition = FormStartPosition.CenterScreen;
            SettingFormatForm.Show();
        }

        
    }
}
