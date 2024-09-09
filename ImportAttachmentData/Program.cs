using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportAttachmentData
{
    class Program
    {
        // https://archive.sap.com/discussions/thread/3550868
        // https://blogs.sap.com/2014/07/28/connecting-an-add-on-to-sap-business-one/

        static void Main(string[] args)
        {
            //#################################################################
            string FilePath = @"C:\Users\erol.akgul\Desktop\";
            string FileName = "prod.csv"; //

            // Check Source File Exists
            if (CheckFileExists(FilePath, FileName))
            {

                SAPbobsCOM.Company oCompany;
                SAP_Bob SAPBOB = new SAP_Bob();
                oCompany = SAPBOB._getCompanyConnection();

                Console.WriteLine(string.Format("Bağlanıyor {0}.", oCompany.CompanyDB));
                // Check Connection is Opened
                if (oCompany.Connected == true)
                {

                    Console.WriteLine(string.Format("Bağlantı Onaylandı"));


                    Console.WriteLine(string.Format("Dosyadan veriler okunuyor {0}.", FileName));
                    // Get Data from File
                    var FileData = GetFileContent(FilePath, FileName);

                    if (FileData.Count > 0)
                    {

                        Console.WriteLine(string.Format("{0} kadar kayıt bulundu", FileData.Count));

                        // Need to Group by Item Codes as I only have one AbsEntry Record to Link into OITM Field
                        var ItemCodes = FileData.GroupBy(n => new { n.ItemCode })
                                                    .Select(g => new
                                                    {
                                                        g.Key.ItemCode
                                                    }).ToList();

                        foreach (var Item in ItemCodes)
                        {
                            Console.WriteLine(string.Format("{0} nolu malzeme kodu işleniyor...", Item.ItemCode));

                            // Get the Files That use the Item Code
                            var Files = FileData.Where(s => s.ItemCode.Equals(Item.ItemCode)).ToList();

                            Console.WriteLine(string.Format("{1} malzeme koduna ait {0} kadar dosya ekleniyor.", Files.Count, Item.ItemCode));

                            // Add the one or More Attachments to the Item Code
                            int intABSEntry = AddAttachments(oCompany, Files);

                            // If added, the and ABS entry is returned, this is used to "link" the files to the Item Code
                            if (intABSEntry > 0)
                            {
                                Console.WriteLine(string.Format("{0} malzeme koduna dosyalar eklendi....", Item.ItemCode));

                                UpdateItem(oCompany, Item.ItemCode, intABSEntry);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(string.Format("{0} dosyası için kayıt bulunamadı...", FileName));
                    }

                }
                else
                {
                    if (oCompany.Connect() != 0)
                    {
                        int errCode = 0;
                        string errMsg = "";
                        oCompany.GetLastError(out errCode, out errMsg);
                        Console.WriteLine(string.Format("Hata Kodu :" + errCode + " - Mesaj :" + errMsg));
                    }
                }
            }
            else
            {
                Console.WriteLine(string.Format("Dosya ismi: {0} Lokasyonunda bulunamadı : {1}.", FileName, FilePath));
            }
            Console.ReadKey(true);
        }

        static bool CheckFileExists(string FileLocation, string FileName)
        {
            if (System.IO.File.Exists(string.Concat(FileLocation, FileName)))
            {
                return true;
            }
            return false;
        }

        static IList<Models.Attachments> GetFileContent(string FilePath, string FileName)
        {
            List<Models.Attachments> Att = new List<Models.Attachments>();

            using (var data = new System.IO.StreamReader(string.Concat(FilePath, FileName)))
            {
                while (!data.EndOfStream)
                {
                    var LineItem = data.ReadLine();
                    var Fields = LineItem.Split(';');// ayrılmış kod tanımlaması , veya ; olabilir

                    Att.Add(new Models.Attachments { ItemCode = Fields[0].ToString(), FilePath = Fields[1].ToString(), FileName = Fields[2].ToString(), FileExtension = Fields[3].ToString() }); 
                }
            }
            return Att;
        }

        static void UpdateItem(SAPbobsCOM.Company oCompany, string ItemCode, int absEntry)
        {
            if (oCompany.Connected == true)
            {
                SAPbobsCOM.Items oSAPItem = null;
                oSAPItem = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                oSAPItem.GetByKey(ItemCode);

                oSAPItem.AttachmentEntry = absEntry;

                int StatusCode = oSAPItem.Update();

                if (StatusCode != 0)
                {
                    string Message = oCompany.GetLastErrorDescription();
                }
                else
                {
                    string ItemNumer;
                    oCompany.GetNewObjectCode(out ItemNumer);
                }
            }
        }

        static int AddAttachments(SAPbobsCOM.Company oCompany, List<Models.Attachments> Attachments)
        {
            int ABSEntry = -1;

            if (oCompany.Connected == true)
            {
                SAPbobsCOM.Attachments2 oSAPItemAttach = null;
                oSAPItemAttach = (SAPbobsCOM.Attachments2)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);

                int LineNo = 0;

                foreach (var Attachment in Attachments)
                {
                    oSAPItemAttach.Lines.Add();
                    oSAPItemAttach.Lines.SetCurrentLine(LineNo);
                    oSAPItemAttach.Lines.SourcePath = Attachment.FilePath;
                    oSAPItemAttach.Lines.FileName = Attachment.FileName;
                    oSAPItemAttach.Lines.FileExtension = Attachment.FileExtension;
                    oSAPItemAttach.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;

                    LineNo++;
                }

                ABSEntry = oSAPItemAttach.Add();

                if (ABSEntry != 0)
                {
                    string Message = oCompany.GetLastErrorDescription();
                }
                else
                {
                    string ItemNumer;
                    oCompany.GetNewObjectCode(out ItemNumer);
                    return int.Parse(ItemNumer);
                }
            }

            return ABSEntry;
        }
    }
}
