using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Graphviz4Net.Dot;
using Graphviz4Net.Dot.AntlrParser;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.IO;
using Microsoft.Office.Interop.Visio;
using System.Drawing;
namespace GraphLibrary
{
    /// <summary>
    /// Класс, объединяющий библиотеку Graphviz4Net и объектную модель Visio
    /// </summary>
    public class VisioGraph
    {
        private GraphParser graphParser = new GraphParser();
        private DotGraph<string> graph;
        private Dictionary<DotVertex<string>, Visio.Shape> vertices = new Dictionary<DotVertex<string>, Visio.Shape>();
        private Dictionary<DotEdge<string>, Visio.Shape> edges = new Dictionary<DotEdge<string>, Visio.Shape>();
        private Dictionary<Visio.Shape, List<Visio.Shape>> newEdges = new Dictionary<Visio.Shape, List<Visio.Shape>>();

        /// <summary>
        /// Конструктор класса
        /// </summary>
        /// <param name="input">Код графа</param>
        public VisioGraph(string input)
        {
            graph = graphParser.ParseGraphData(input);
        }

        /// <summary>
        /// Функция отображения графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница в Visio</param>
        public void PresentGraphInVisio(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            PresentNodes(visioDocs, visioPage);
            PresentEdges(visioDocs, visioPage);

            visioPage.Layout();
        }

        /// <summary>
        /// Функция отображения вершин графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница в Visio</param>
        private void PresentNodes(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            // Мастер-объект базовых фигур Visio
            Dictionary<string, Visio.Master> visioMasters = GetMasterShapes(visioDocs);

            // Расстановка вершин графа
            for (int i = 0; i < graph.AllVertices.Count(); ++i)
            {
                // Вершина
                var node = graph.AllVertices.ElementAt(i);

                // Стили вершины
                string shape = node.Attributes.ContainsKey("shape") ? node.Attributes["shape"] : "ELLIPSE";
                string label = node.Attributes.ContainsKey("label") ? node.Attributes["label"] : node.Id;
                string color = node.Attributes.ContainsKey("color") ? node.Attributes["color"] : "black";
                string fontcolor = node.Attributes.ContainsKey("fontcolor") ? node.Attributes["fontcolor"] : "black";
                string style = node.Attributes.ContainsKey("style") ? node.Attributes["style"] == "filled" ? "filled" : LineStyle(node.Attributes["style"].ToLower()) : "1";

                // Добавление вершины на страницу Visio
                vertices.Add(node, visioPage.Drop(visioMasters[shape.ToUpper()], 1 + i / 2.0, 11 - i / 2.0));

                // Установка стилей для фигуры на странице Visio
                vertices[node].Text = label;
                vertices[node].get_CellsSRC((short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowCharacter, (short)Visio.VisCellIndices.visCharacterColor).FormulaU = VisioColor.ColorToRgb(fontcolor.ToLower());
                if (style == "filled")
                    vertices[node].get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visFillForegnd).FormulaU = VisioColor.ColorToRgb(color.ToLower());
                else
                    vertices[node].get_CellsU("LinePattern").FormulaU = style;
                vertices[node].get_CellsU("LineColor").FormulaU = VisioColor.ColorToRgb(color.ToLower());

                // Ресайзинг
                vertices[node].Resize(Visio.VisResizeDirection.visResizeDirNW, -0.8, Visio.VisUnitCodes.visInches);

                // Чтобы не допустить пересечений
                vertices[node].get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowMisc, (short)Visio.VisCellIndices.visLOFlags).FormulaU = "1";
                vertices[node].get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowShapeLayout, (short)Visio.VisCellIndices.visSLOPlowCode).FormulaU = "2";
            }
        }

        /// <summary>
        /// Функция отображения ребер графа в документе Visio
        /// </summary>
        /// <param name="visioDocs">Документы Visio</param>
        /// <param name="visioPage">Текущая страница в Visio</param>
        private void PresentEdges(Visio.Documents visioDocs, Visio.Page visioPage)
        {
            Visio.Document visioConnectors = visioDocs.OpenEx("Basic Flowchart Shapes (US units).vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);

            // Соединение вершин графа ребрами
            for (int i = 0; i < graph.VerticesEdges.Count(); ++i)
            {
                // Ребро
                var edge = graph.VerticesEdges.ElementAt(i);

                // Фигура соединидельной линии (коннектора)
                Visio.Shape connector = visioPage.Drop(visioConnectors.Masters.get_ItemU("Dynamic connector"), 0, 0);
                connector.get_Cells("ConLineRouteExt").FormulaU = "2";
                connector.get_Cells("EndArrow").Formula = "=5";

                // Стили ребра
                string label = edge.Attributes.ContainsKey("label") ? edge.Attributes["label"] : "";
                string color = edge.Attributes.ContainsKey("color") ? edge.Attributes["color"] : "black";
                string linestyle = edge.Attributes.ContainsKey("style") ? LineStyle(edge.Attributes["style"].ToLower()) : "1";

                // Установка стилей для фигуры на странице Visio
                connector.Text = label;
                connector.get_CellsU("LineColor").FormulaU = VisioColor.ColorToRgb(color.ToLower());
                connector.get_CellsU("LinePattern").FormulaU = linestyle;

                // Соединение вершин при помощи данного коннектора
                vertices[edge.Source].AutoConnect(vertices[edge.Destination], Visio.VisAutoConnectDir.visAutoConnectDirDown, connector);

                edges.Add(edge, vertices[edge.Source].FromConnects[vertices[edge.Source].FromConnects.Count].FromSheet);

                // Удаление коннектора-болванки
                connector.Delete();
            }
        }

        /// <summary>
        /// Метод инвертирования ребра
        /// </summary>
        /// <param name="window"></param>
        public void Invert(Visio.Window window)
        {
            if (window != null)
            {
                // Для всех выделенных фигур
                foreach (Visio.Shape shape in window.Selection)
                {
                    // Проверяем, что это ребро
                    if (edges.ContainsValue(shape))
                    {
                        foreach (var edge in edges)
                        {
                            if (edge.Value == shape)
                            {
                                // Инвертируем ребро
                                Visio.Cell beginXCell = shape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXForm1D, (short)Visio.VisCellIndices.vis1DBeginX);
                                beginXCell.GlueTo(vertices[edge.Key.Destination].get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX));
                                Visio.Cell endXCell = shape.get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXForm1D, (short)Visio.VisCellIndices.vis1DEndX);
                                endXCell.GlueTo(vertices[edge.Key.Source].get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX));

                                // Заменяем старое ребро новым инвертированным
                                DotEdge<string> invertedEdge = new DotEdge<string>(edge.Key.Destination, edge.Key.Source, edge.Key.Attributes);
                                break;
                            }
                        }
                    }
                    else
                    {
                        throw new ArgumentException("Допустимо только инвертирование ребер!");
                    }
                }
            }
        }

        /// <summary>
        /// Метод выделения
        /// </summary>
        /// <param name="key"></param>
        /// <param name="window"></param>
        public void Select(int key, Visio.Window window)
        {
            if (window != null)
            {
                window.DeselectAll();
                switch (key)
                {
                    // Выделить все вершины
                    case 1:
                        foreach (var node in vertices)
                        {
                            window.Select(node.Value, 2);
                        }
                        break;
                    // Выделить все соединенные вершины
                    case 2:
                        foreach (var node in vertices)
                        {
                            if (node.Value.ConnectedShapes(Visio.VisConnectedShapesFlags.visConnectedShapesAllNodes, "").Length != 0)
                            {
                                window.Select(node.Value, 2);
                            }
                        }
                        break;
                    // Выделить все несоединенные вершины
                    case 3:
                        foreach (var node in vertices)
                        {
                            if (node.Value.ConnectedShapes(Visio.VisConnectedShapesFlags.visConnectedShapesAllNodes, "").Length == 0)
                            {
                                window.Select(node.Value, 2);
                            }
                        }
                        break;
                    // Выделить все ребра
                    case 4:
                        foreach (var edge in edges)
                        {
                            window.Select(edge.Value, 2);
                        }
                        break;
                    default:
                        throw new ArgumentException("Ключ не соответствует допустимому диапозону");
                }
            }
        }

        /// <summary>
        /// Метод, вызывающийся при добавлении вершины
        /// </summary>
        /// <param name="shape"></param>
        public void AddNode(Visio.Shape shape)
        {
            // Проверяем, что это соединительная линия
            if (shape.Master.NameU == "Dynamic connector")
            {
            }
            else
            {
                // Если новая фигура допустимой формы
                if (GetDotShapes().ContainsKey(shape.Master.NameU.ToUpper()))
                {
                    if (shape.Text == "")
                    {
                        shape.Text = shape.GetHashCode().ToString();
                    }
                    DotVertex<string> newNode = new DotVertex<string>(shape.GetHashCode().ToString());
                    newNode.Attributes.Add("shape", GetDotShapes()[shape.Master.NameU.ToUpper()]);
                    newNode.Attributes.Add("label", shape.Text);
                    
                    graph.AddVertex(newNode);
                    vertices.Add(newNode, shape);
                }
                else
                {
                    throw new ArgumentException("Фигура не поддерживается языком DOT! Она не будет включена в список вершин графа!");
                }
            }
        }

        /// <summary>
        /// Метод, вызывающийся при изменении текста фигуры
        /// </summary>
        /// <param name="shape"></param>
        public void ChangeLabel(Visio.Shape shape)
        {
            // Если изменен текст вершины
            if (vertices.ContainsValue(shape))
            {
                foreach (var node in vertices)
                {
                    if (node.Value == shape)
                    {
                        if (node.Key.Attributes.ContainsKey("label"))
                        {
                            node.Key.Attributes["label"] = shape.Text;
                        }
                        else
                        {
                            node.Key.Attributes.Add("label", shape.Text);
                        }
                        break;
                    }
                }
            }
            else
            // Если изменен текст ребра
            if (edges.ContainsValue(shape))
            {
                foreach (var edge in edges)
                {
                    if (edge.Value == shape)
                    {
                        if (edge.Key.Attributes.ContainsKey("label"))
                        {
                            edge.Key.Attributes["label"] = shape.Text;
                        }
                        else
                        {
                            edge.Key.Attributes.Add("label", shape.Text);
                        }
                        break;
                    }
                }
            }
        }
        /// <summary>
        /// Метод, вызывающийся при изменении цвета фигуры
        /// </summary>
        /// <param name="shape"></param>
        public void ChangeColor(Visio.Shape shape)
        {
            // Если изменен цвет вершины
            /*
            if (vertices.ContainsValue(shape))
            {
                foreach (var node in vertices)
                {
                    if (node.Value == shape)
                    {

                        if (node.Key.Attributes.ContainsKey("color"))
                        {
                            node.Key.Attributes["color"] = "black";
                        }
                        else
                        {
                            node.Key.Attributes.Add("color", "black");
                        }
                        break;
                    }
                }
            }
            */
        }

        /// <summary>
        /// Метод, вызывающийся при добавлении нового ребра между вершинами графа
        /// </summary>
        /// <param name="connects">Коннектор</param>
        public void AddEdge(Visio.Connects connects)
        {
            // Если соединение добавлено с использованием существующей соединительной линии, т.е. добавляем вторую вершину
            if (newEdges.ContainsKey(connects.FromSheet))
            {
                // Добавляем вторую вершину в пару связанных вершин
                newEdges[connects.FromSheet].Add(connects.ToSheet);

                // Если обе соединенные фигуры - это вершины нашего графа
                if (vertices.ContainsValue(newEdges[connects.FromSheet][0]) && vertices.ContainsValue(newEdges[connects.FromSheet][1]))
                {
                    DotVertex<string> from = new DotVertex<string>("");
                    DotVertex<string> to = new DotVertex<string>("");

                    // Ищем, каким вершинам они соответствуют
                    foreach (var node in vertices)
                    {
                        if (node.Value == newEdges[connects.FromSheet][0])
                        {
                            from = node.Key;
                        }
                        if (node.Value == newEdges[connects.FromSheet][1])
                        {
                            to = node.Key;
                        }
                    }

                    // Добавляем ребро с данными вершинами
                    DotEdge<string> edge = new DotEdge<string>(from, to);
                    graph.AddEdge(edge);

                    // Добавляем новое ребро в список ребер
                    edges.Add(edge, connects.FromSheet);
                }

                // Удаляем пару
                newEdges.Remove(connects.FromSheet);
            }
            else
            {
                // Применяем стили к коннектору
                connects.FromSheet.get_Cells("ConLineRouteExt").FormulaU = "2";
                connects.FromSheet.get_Cells("EndArrow").Formula = "=5";

                // Добавляем новый коннектор в словарь + вершину, от которой он стартовал, ожидая соединения ее со второй вершиной
                List<Visio.Shape> listOfConnectingNodes = new List<Visio.Shape>();
                listOfConnectingNodes.Add(connects.ToSheet);
                newEdges.Add(connects.FromSheet, listOfConnectingNodes);
            }
        }

        /// <summary>
        /// Метод, вызывающийся при удалении ребра
        /// </summary>
        /// <param name="connects"></param>
        public void DeleteEdge(Visio.Connects connects)
        {
            Visio.Shape shape = connects.FromSheet;
            if (edges.ContainsValue(shape))
            {
                foreach (var edge in edges)
                {
                    if (edge.Value == shape)
                    {
                        graph.RemoveEdge(edge.Key);
                        edges.Remove(edge.Key);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Метод, вызывающийся при удалении фигуры
        /// </summary>
        /// <param name="shape"></param>
        public void DeleteShape(Visio.Shape shape)
        {
            // Если удалено ребро
            if (edges.ContainsValue(shape))
            {
                Visio.Shape v1 = shape.Connects[1].ToSheet;
                Visio.Shape v2 = shape.Connects[2].ToSheet;
                if (vertices.ContainsValue(v1) && vertices.ContainsValue(v2))
                {
                    foreach (var edge in edges)
                    {
                        if (edge.Value == shape)
                        {
                            graph.RemoveEdge(edge.Key);
                            edges.Remove(edge.Key);
                            break;
                        }
                    }
                }
            }
            else
            // Если удалена вершина
            if (vertices.ContainsValue(shape))
            {
                foreach (var node in vertices)
                {
                    if (node.Value == shape)
                    {
                        // Удаляем все ребра, смежные с данной вершиной
                        Dictionary<DotEdge<string>, Visio.Shape> edgesToDelete = new Dictionary<DotEdge<string>, Visio.Shape>();
                        foreach (var edge in edges)
                        {
                            if (edge.Key.Source == node.Key || edge.Key.Destination == node.Key)
                            {
                                graph.RemoveEdge(edge.Key);
                                edgesToDelete.Add(edge.Key, edge.Value);
                            }
                        }
                        foreach (var edge in edgesToDelete)
                        {
                            edges.Remove(edge.Key);
                            edge.Value.Delete();
                        }

                        // Удаляем вершину
                        graph.RemoveVertex(node.Key);
                        vertices.Remove(node.Key);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Функция, сопоставляющая ключу-строке из допустимых фигур DOT мастер-фигуру Visio
        /// </summary>
        /// <param name="visioDocs"></param>
        /// <returns>Словарь string-Visio.Master</returns>
        private Dictionary<string, string> GetDotShapes()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            result.Add("TRIANGLE", "triangle");
            result.Add("SQUARE", "square");
            result.Add("HEXAGON", "hexagon");
            result.Add("OCTAGON", "octagon");
            result.Add("RECTANGLE", "rectangle");
            result.Add("CIRCLE", "circle");
            result.Add("ELLIPSE", "ellipse");
            result.Add("DIAMOND", "diamond");
            return result;
        }

        /// <summary>
        /// Функция, сопоставляющая ключу-строке из допустимых фигур DOT мастер-фигуру Visio
        /// </summary>
        /// <param name="visioDocs"></param>
        /// <returns>Словарь string-Visio.Master</returns>
        private Dictionary<string, Visio.Master> GetMasterShapes(Visio.Documents visioDocs)
        {
            Visio.Document visioStencil1 = visioDocs.OpenEx("Basic Shapes.vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);
            Visio.Document visioStencil2 = visioDocs.OpenEx("Audit Diagram Shapes.vss", (short)Visio.VisOpenSaveArgs.visOpenDocked);

            Dictionary<string, Visio.Master> result = new Dictionary<string, Visio.Master>();
            result.Add("TRIANGLE", visioStencil1.Masters.get_ItemU(@"Triangle"));
            result.Add("SQUARE", visioStencil1.Masters.get_ItemU(@"Square"));
            result.Add("PENTAGON", visioStencil1.Masters.get_ItemU(@"Pentagon"));
            result.Add("HEXAGON", visioStencil1.Masters.get_ItemU(@"Hexagon"));
            result.Add("OCTAGON", visioStencil1.Masters.get_ItemU(@"Octagon"));
            result.Add("RECTANGLE", visioStencil1.Masters.get_ItemU(@"Rectangle"));
            result.Add("RECT", visioStencil1.Masters.get_ItemU(@"Rectangle"));
            result.Add("BOX", visioStencil1.Masters.get_ItemU(@"Rectangle"));
            result.Add("CIRCLE", visioStencil1.Masters.get_ItemU(@"Circle"));
            result.Add("ELLIPSE", visioStencil1.Masters.get_ItemU(@"Ellipse"));
            result.Add("OVAL", visioStencil1.Masters.get_ItemU(@"Ellipse"));
            result.Add("DIAMOND", visioStencil1.Masters.get_ItemU(@"Diamond"));
            result.Add("PARALLELOGRAM", visioStencil2.Masters.get_ItemU(@"I/O"));
            return result;
        }

        /// <summary>
        /// Функция, возвращающая номер в формуле стиля линии
        /// </summary>
        /// <param name="styleString">Стиль в строковом формате</param>
        /// <returns>Номер формулы</returns>
        private string LineStyle(string styleString)
        {
            switch (styleString)
            {
                case "dashed":
                    return "2";
                case "dotted":
                    return "3";
                default:
                    return "1";
            }
        }

        /// <summary>
        /// Функция экспорта графа в файл
        /// </summary>
        /// <param name="filePath">Путь к файлу</param>
        public void ExportGraph(string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath))
            {
                new GraphToDotConverter().Convert(sw, graph, new AttributesProvider());
            }
        }
    }

    public class AttributesProvider : IAttributesProvider
    {
        public IDictionary<string, string> GetVertexAttributes(object vertex)
        {
            return new Dictionary<string, string>();
        }
    }
}
