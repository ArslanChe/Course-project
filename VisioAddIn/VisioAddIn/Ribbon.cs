using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace VisioAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void openFileButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string input = File.ReadAllText(openFileDialog.FileName);
                try
                {
                    Globals.ThisAddIn.ShowGraph(input);
                }
                catch (Exception exc)
                {
                    Globals.ThisAddIn.RemovePageIfError();
                    Globals.ThisAddIn.ErrorMessage("Во время импорта графа возникла следующая ошибка:\n" + exc.Message, "Не удалось отобразить граф");
                }
            }
        }

        private void exportGraphButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Globals.ThisAddIn.ExportGraph(saveFileDialog.FileName);
                }
                catch (Exception exc)
                {
                    Globals.ThisAddIn.ErrorMessage("Во время экспорта графа возникла следующая ошибка:\n\n" + exc.Message, "Не удалось экспортировать граф");
                }
            }
        }

        private void selectAllNodesButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Select(1);
        }

        private void selectConnectedNodeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Select(2);
        }

        private void selectNonConnectedNodesButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Select(3);
        }

        private void selectEdgesButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Select(4);
        }

        private void invertButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Invert();
        }

        private void layoutButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Layout();
        }
    }
}
