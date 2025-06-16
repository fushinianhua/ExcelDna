using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Radiant.MyClass
{
    internal class ItemData
    {
        public string Name { get; set; }//名称
        public int Row { get; set; }//行号
        public int Count { get; set; }//数量

        public ItemData(string name, int row, int count)//构造函数
        {
            Name = name;
            Row = row;
            Count = count;
        }
    }
}