using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PM_Helper
{
    class FileForTranslation
    {        
        public string Name { get; set; }
                
        public string SourceLanguage { get; set; }
                    
        public string SourceLanguageCode { get; set; }
               
        public string TargetLanguage { get; set; }
                       
        public string TargetLanguageCode { get; set; }
             
        public int Repetitions { get; set; }       
        
        public int Match100 { get; set; }
        
        public int Match95_99 { get; set; }      

        public int Match85_94 { get; set; }      
       
        public int Match75_84 { get; set; }
       
        public int MachineTrans { get; set; }      
        
        public int NoMatch { get; set; }

        public int NonTranslatable { get; set; }

        public int IceMatch { get; set; }

        private int total;

        public int XTMtotal
        {
            get { return total; }
        }

        public FileForTranslation(string name, string sourceLanguagee, string targetLanguage, int repetitions, int match100, int match95_99, int match85_94, int match75_84, int machineTrans, int noMatch, int nonTranslatable, int iceMatch, int total)
        {
            this.Name = name;
            this.SourceLanguage = sourceLanguagee;
            this.TargetLanguage = targetLanguage;
            this.Repetitions = repetitions;
            this.Match100 = match100;
            this.Match95_99 = match95_99;
            this.Match85_94 = match85_94;
            this.Match75_84 = match75_84;
            this.MachineTrans = machineTrans;
            this.NoMatch = noMatch;
            this.NonTranslatable = nonTranslatable;
            this.IceMatch = iceMatch;
            this.total = total;
        }

        public int TotalIncluding100()
        {
            int result = this.total - (this.NonTranslatable + this.IceMatch);
            return result;
        }

        public int TotalWithout100()
        {
            int result = this.total - (this.NonTranslatable + this.IceMatch + this.Match100);
            return result;
        }

    }
}
