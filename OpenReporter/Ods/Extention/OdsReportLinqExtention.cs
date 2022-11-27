using System.Xml;

namespace Rugal.Net.OpenReporter.Ods.Extention
{
    public static class OdsReportLinqExtention
    {
        #region Node List Linq
        public static IEnumerable<XmlNode> Select(this XmlNodeList NodeList)
        {
            return NodeList.Cast<XmlNode>();
        }
        public static IEnumerable<XmlNode> Where(this XmlNodeList NodeList, Func<XmlNode, bool> WhereFunc)
        {
            return NodeList.Select().Where(Item => WhereFunc.Invoke(Item));
        }
        public static XmlNode FirstOrDefault(this XmlNodeList NodeList, Func<XmlNode, bool> WhereFunc)
        {
            return NodeList.Select().FirstOrDefault(Item => WhereFunc.Invoke(Item));
        }
        public static XmlNode First(this XmlNodeList NodeList, Func<XmlNode, bool> WhereFunc)
        {
            return NodeList.Select().FirstOrDefault(Item => WhereFunc.Invoke(Item));
        }
        #endregion

        public static IEnumerable<XmlNode> Select(this XmlNode Node)
        {
            return Node.ChildNodes.Select();
        }

        public static IEnumerable<XmlNode> ForEach(this IEnumerable<XmlNode> NodeList, Action<XmlNode> EachAction)
        {
            foreach (var Item in NodeList)
                EachAction.Invoke(Item);
            return NodeList;
        }
    }
}
