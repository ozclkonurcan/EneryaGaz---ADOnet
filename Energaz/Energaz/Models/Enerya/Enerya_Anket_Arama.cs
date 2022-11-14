using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace EneryaGaz.Models.Enerya
{
    public class Enerya_Anket_Arama
    {
        public string MANDT { get; set; }
        public DateTime TARIH { get; set; }
        public string SOZHESAP { get; set; }
        public string SURECADI { get; set; }
        public string PERADSOYAD { get; set; }
        public string MUSADSOYAD { get; set; }
        public string CEPTEL { get; set; }
        public string USER { get; set; }
        public string LOKASYON { get; set; }
        public string SEHIR { get; set; }

    }
}
