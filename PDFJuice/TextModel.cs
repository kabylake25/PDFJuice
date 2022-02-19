using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFJuice
{
    public class TextModel
    {
        public TextModel() { }
        public TextModel(TextModel txtModel)
        {
            this.Kap1 = txtModel.Kap1;
            this.Text1 = txtModel.Text1;
            this.Kap2 = txtModel.Kap2;
            this.Text2 = txtModel.Text2;
            this.Kap3 = txtModel.Kap3;
            this.Text3 = txtModel.Text3;
            this.Kap4 = txtModel.Kap4;
            this.Text4 = txtModel.Text4;
            this.Kap5 = txtModel.Kap5;
            this.Text5 = txtModel.Text5;
            this.Text = txtModel.Text;
            this.Font = txtModel.Font;
            this.Menge = txtModel.Menge;
            this.Me = txtModel.Me;
            this.Underlined = txtModel.Underlined;
        }
        public string Kap1 { get; set; }
        public string Text1 { get; set; }
        public string Kap2 { get; set; }
        public string Text2 { get; set; }
        public string Kap3 { get; set; }
        public string Text3 { get; set; }
        public string Kap4 { get; set; }
        public string Text4 { get; set; }
        public string Kap5 { get; set; }
        public string Text5 { get; set; }
        public string Text { get; set; }
        public string Font { get; set; }
        public string Menge { get; set; }
        public string Me { get; set; }
        public string Underlined { get; set; }
    }
}
