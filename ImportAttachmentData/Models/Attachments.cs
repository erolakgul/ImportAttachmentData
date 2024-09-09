using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportAttachmentData.Models
{
   public class Attachments 
    {
        public string ItemCode { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string FileExtension { get; set; }
    }
}
