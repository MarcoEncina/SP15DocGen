using System;
using System.Collections.Generic;
using System.Text;

namespace Sezai.SPDocGen
{
    /// <summary>
    /// This struct is a two string tuple used for storing attribute-value-pairs returned from SharePoint object properties.
    /// </summary>
    public struct AttributeValuePair
    {     
        public string Attribute;        
        public string Value;
        public AttributeValuePair(string Attribute, string Value)
        {
            this.Attribute = Attribute;
            this.Value = Value;
        }
        public AttributeValuePair(string Attribute)
        {
            this.Attribute = Attribute;
            this.Value = "";
        }
    }
}
