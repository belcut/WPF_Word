using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_Word
{
    class UiTbItem
    {
        
        public string ElementName;
        public string Text;

        public UiTbItem (string elementName, string text)
        {
            
            ElementName = elementName;
            Text = text;
        }
    }
}
