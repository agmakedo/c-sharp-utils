using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.IO;
using System.Web.UI;

namespace Utility.HTMLWrite
{
    class HTMLWrite
    {
        private static StringWriter HtmlWriter;
        private static HtmlTextWriter Writer;
        
        public HTMLWrite()
        {
          HtmlWriter = new StringWriter();                                                    
          Writer = new HtmlTextWriter(HtmlWriter);
          
          Writer.RenderBeginTag("html");
          Writer.RenderBeginTag("body");
        }
        
        public static void CreateTable(string[] Headers, string[][] Body,string[] style = null)
        {           
            if (style != null)
            {
                Writer.RenderBeginTag("style");
                foreach(string format in style)
                {
                    Writer.Write(format);
                }
                Writer.RenderEndTag();
            }

            Writer.RenderBeginTag(HtmlTextWriterTag.Table);
            Writer.RenderBeginTag("thead");
            Writer.RenderBeginTag("tr");

            foreach(string Title in Headers)
            {
                Writer.RenderBeginTag("th");
                Writer.RenderBeginTag("b");
                Writer.Write(Title);
                Writer.RenderEndTag();
                Writer.RenderEndTag();
            }

            Writer.RenderEndTag();
            Writer.RenderEndTag();

            foreach (string[] row in Body)
            {
                foreach(string Data in row)
                {
                    Writer.RenderBeginTag("tr");
                    Writer.RenderBeginTag("td");
                    Writer.Write(Data);
                    Writer.RenderEndTag();
                }
                Writer.RenderEndTag();
            }
        }
        
        public static void EndHTML(int TagCount = 2)
        {
          for(int i = 0; i< TagCount; i++)
            Writer.RenderEndTag();
        }
             
    }
