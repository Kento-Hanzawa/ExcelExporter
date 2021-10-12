using Microsoft.Office.Interop.Excel;

namespace ExcelInteropBridging.Internal
{
    public sealed class RangeInfo
    {
        public string RangeName { get; }
        public string RangeString { get; }

        internal RangeInfo(IComManaged<Worksheet> source)
        {
            RangeName = source.ComObject.Name;
            RangeString = source.GetRangeString();
        }

        internal RangeInfo(IComManaged<ListObject> source)
        {
            RangeName = source.ComObject.Name;
            RangeString = source.GetRangeString();
        }
    }
}
