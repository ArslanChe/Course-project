using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using GraphLibrary;


namespace VisioAddIn
{
    public partial class ThisAddIn
    {
        private Dictionary<Visio.Page, VisioGraph> graphs = new Dictionary<Visio.Page, VisioGraph>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }
        
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }
        
        /// <summary>
        /// Метод отображения графа в Visio
        /// </summary>
        /// <param name="input"></param>
        public void ShowGraph(string input)
        {
            Application.ActiveDocument.Pages.BeforePageDelete += Globals.ThisAddIn.DeleteGraph;

            Visio.Documents visioDocs = Application.Documents;
            Visio.Page visioPage = Application.ActiveDocument.Pages.Add();

            graphs.Add(visioPage, new VisioGraph(input));
            graphs[visioPage].PresentGraphInVisio(visioDocs, visioPage);

            Application.ActivePage.BeforeShapeDelete += DeleteShape;
            Application.ActivePage.ConnectionsDeleted += DeleteEdge;
            Application.ActivePage.ConnectionsAdded += AddEdge;
            Application.ActivePage.TextChanged += ChangeText;
            Application.ActivePage.ShapeAdded += AddShape;
        }

        /// <summary>
        /// Метод добавления фигуры
        /// </summary>
        /// <param name="Shape"></param>
        private void AddShape(Visio.Shape Shape)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                try
                {
                    graphs[Application.ActivePage].AddNode(Shape);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время добавления вершины возникла следующая ошибка:\n" + exc.Message, "Не удалось добавить вершину");
                }
            }
        }

        /// <summary>
        /// Метод, вызывающий при попытке изменить ребро
        /// </summary>
        public void Invert() => graphs[Application.ActivePage].Invert(Application.ActiveWindow);

        /// <summary>
        /// Лэйаутинг
        /// </summary>
        public void Layout()
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                try
                {
                    Application.ActivePage.Layout();
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время планировки возникла следующая ошибка:\n" + exc.Message, "Не удалось исправить планировку");
                }
            }
            else
            {
                ErrorMessage("На данной странице отсутствует граф!", "Не удалось исправить планировку");
            }
        }
        /// <summary>
        /// Метод выделения вершин
        /// </summary>
        /// <param name="key"></param>
        public void Select(int key)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                try
                {
                    graphs[Application.ActivePage].Select(key, Application.ActiveWindow);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время выделения вершин возникла следующая ошибка:\n" + exc.Message, "Не удалось выделить вершины");
                }
            }
        }

        /// <summary>
        /// Метод изменения текста вершины
        /// </summary>
        /// <param name="Shape"></param>
        private void ChangeText(Visio.Shape Shape)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                try
                {
                    graphs[Application.ActivePage].ChangeLabel(Shape);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время изменения текста возникла следующая ошибка:\n" + exc.Message, "Не удалось изменить текст");
                }
            }
        }
        /// <summary>
        /// Метод изменения текста вершины
        /// </summary>
        /// <param name="Shape"></param>
        private void ChangeColor(Visio.Shape Shape)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                try
                {
                    graphs[Application.ActivePage].ChangeColor(Shape);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время изменения цвета возникла следующая ошибка:\n" + exc.Message, "Не удалось изменить цвет");
                }
            }
        }

        /// <summary>
        /// Метод разрыва соединения вершин
        /// </summary>
        /// <param name="Connects"></param>
        private void DeleteEdge(Visio.Connects Connects)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                try
                {
                    graphs[Application.ActivePage].DeleteEdge(Connects);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время удаления ребра возникла следующая ошибка:\n" + exc.Message, "Не удалось удалить ребро");
                }

            }
        }

        /// <summary>
        /// Метод удаления фигуры
        /// </summary>
        /// <param name="Shape"></param>
        private void DeleteShape(Visio.Shape Shape)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                try
                {
                    graphs[Application.ActivePage].DeleteShape(Shape);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время удаления объекта возникла следующая ошибка:\n" + exc.Message, "Не удалось удалить объект");
                }
            }
        }

        /// <summary>
        /// Метод удаления ребра
        /// </summary>
        /// <param name="Connects"></param>
        private void AddEdge(Visio.Connects Connects)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                try
                {
                    graphs[Application.ActivePage].AddEdge(Connects);
                }
                catch (Exception exc)
                {
                    ErrorMessage("Во время добавления ребра возникла следующая ошибка:\n" + exc.Message, "Не удалосьдобавить ребро");
                }
            }
        }

        /// <summary>
        /// Метод удаления страницы, в случае, если возникла ошибка
        /// </summary>
        public void RemovePageIfError()
        {
            DeleteGraph(Application.ActiveDocument.Pages[Application.ActiveDocument.Pages.Count]);
            Application.ActiveDocument.Pages[Application.ActiveDocument.Pages.Count].Delete(1);
        }

        /// <summary>
        /// Удаление графа из словаря, если была удалена страница
        /// </summary>
        /// <param name="Page"></param>
        private void DeleteGraph(Visio.Page Page)
        {
            if (graphs.ContainsKey(Page))
            {
                graphs.Remove(Page);
            }
        }

        /// <summary>
        /// Метод экспорта графа в файл
        /// </summary>
        /// <param name="filePath"></param>
        public void ExportGraph(string filePath)
        {
            if (graphs.ContainsKey(Application.ActivePage))
            {
                graphs[Application.ActivePage].ExportGraph(filePath);
            }
            else throw new ArgumentException("На данной странице не представлен граф");
        }

        /// <summary>
        /// Метод для отображения сообщения об ошибке
        /// </summary>
        /// <param name="message"></param>
        /// <param name="caption"></param>
        public void ErrorMessage(string message, string caption)
        {
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            // Отобразить окошко об ошибке
            result = MessageBox.Show(message, caption, buttons);
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
