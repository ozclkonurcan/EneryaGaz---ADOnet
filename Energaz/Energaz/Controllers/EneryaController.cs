using ClosedXML.Excel;
using Energaz.VM_Models;
using EneryaGaz.Models.Enerya;
using Microsoft.Ajax.Utilities;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Energaz.Controllers
{
    public class EneryaController : Controller
    {
        private List<Enerya_Anket_Arama> anketList = new List<Enerya_Anket_Arama>();
        private List<Enerya_Personel_Bazli_Cagri> personelBazliCagriList = new List<Enerya_Personel_Bazli_Cagri>();
        public void EneryaAnketAramaFunc(string TARIH)
        {
            try
            {

                if (TARIH == null || TARIH == "")
                    TARIH = DateTime.UtcNow.ToString("yyyy-MM-dd");

                var client = new RestClient("http://egpoq00.enr.local:50000/XISOAPAdapter/MessageServlet?senderParty=&senderService=COMDATA_IVR&receiverParty=&receiverService=&interface=SI_CCAnketArama_Sync_OB&interfaceNamespace=http://energaz.com/ComdataCallCenterAnketArama");
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("SOAPAction", "http://sap.com/xi/WebService/soap1.1");
                request.AddHeader("Content-Type", "text/xml;charset=UTF-8");
                request.AddHeader("Authorization", "Basic UkZDVVNFUjphQTEyMzQ1Ng==");
                request.AddHeader("Cookie", "JSESSIONID=0K6u_4yjCDCZngpMxM-7V8_FMukThAESj4YA_SAPRhn5SsC-YaW0xTJMTNpYNpvH; JSESSIONMARKID=NMn5CQu3X5Apk_e955bHYrOL9NQRTLGwYzGhKPhgA; saplb_*=(J2EE8818420)8818450");

                var body = "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:urn=\"urn:sap-com:document:sap:rfc:functions\">" +
                    "<soapenv:Header/>" +
                    " <soapenv:Body>" +
                    "<urn:ZMD_WS_USER_PRS_BULK>" +
                    "<!--You may enter the following 4 items in any order-->" +
                    "<!--Optional:-->" +
                    "<BUKRS>3000</BUKRS>" +
                    "<!--Optional:-->" +
                    "<FLG></FLG>" +
                    "<!--Optional:-->" +
                    "<TARIH>" + TARIH + "</TARIH>" +
                    " <!--Optional:-->" +
                    "<GT_DATA>" +
                    "<!--Zero or more repetitions:-->" +
                    "<item>" +
                    "<!--Optional:-->" +
                    "<MANDT></MANDT>" +
                    "<!--Optional:-->" +
                    "<TARIH></TARIH>" +
                    "<!--Optional:-->" +
                    "<SOZHESAP></SOZHESAP>" +
                    "<!--Optional:-->" +
                    "<SURECADI></SURECADI>" +
                    "<!--Optional:-->" +
                    "<PERADSOYAD></PERADSOYAD>" +
                    "<!--Optional:-->" +
                    "<MUSADSOYAD></MUSADSOYAD>" +
                    "<!--Optional:-->" +
                    "<CEPTEL></CEPTEL>" +
                    "<!--Optional:-->" +
                    "<USER></USER>" +
                    "<!--Optional:-->" +
                    "<LOKASYON></LOKASYON>" +
                    "<!--Optional:-->" +
                    "<SEHIR></SEHIR>" +
                    "</item>" +
                    "</GT_DATA>" +
                    "</urn:ZMD_WS_USER_PRS_BULK>" +
                    "</soapenv:Body>" +
                    "</soapenv:Envelope>";


                request.AddParameter("text/xml;charset=UTF-8", body, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);

                List<Enerya_Anket_Arama> obList = new List<Enerya_Anket_Arama>();
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(response.Content);
                XmlNodeList elemList = doc.GetElementsByTagName("item");


                foreach (XmlNode chldNode in elemList)
                {
                    obList.Add(new Enerya_Anket_Arama
                    {
                        //Id = int.Parse(chldNode.InnerText),
                        MANDT = chldNode.ChildNodes[0].InnerText,
                        TARIH = DateTime.Parse(chldNode.ChildNodes[1].InnerText),
                        SOZHESAP = chldNode.ChildNodes[2].InnerText,
                        SURECADI = chldNode.ChildNodes[3].InnerText,
                        PERADSOYAD = chldNode.ChildNodes[4].InnerText,
                        MUSADSOYAD = chldNode.ChildNodes[5].InnerText,
                        CEPTEL = chldNode.ChildNodes[6].InnerText,
                        USER = chldNode.ChildNodes[7].InnerText,
                        LOKASYON = chldNode.ChildNodes[8].InnerText,
                        SEHIR = chldNode.ChildNodes[9].InnerText,


                    });

                    anketList.AddRange(obList);

                    var anketSorguResp = (from a in obList
                                          select new VM_Enerya_Anket_Arama()
                                          {
                                              MANDT = a.MANDT,
                                              TARIH = a.TARIH.ToString("dd/MM/yyyy"),
                                              SOZHESAP = a.SOZHESAP,
                                              SURECADI = a.SURECADI,
                                              PERADSOYAD = a.PERADSOYAD,
                                              MUSADSOYAD = a.MUSADSOYAD,
                                              CEPTEL = a.CEPTEL,
                                              USER = a.USER,
                                              LOKASYON = a.LOKASYON,
                                              SEHIR = a.SEHIR,
                                          }).ToList();


                    ViewBag.rest = anketSorguResp.DistinctBy(x => x.CEPTEL);

                }
            }
            catch (Exception ex)
            {

                Console.WriteLine("HATA" + ex.Message);
                ViewBag.HATA = "HATA"+ex.Message;
            }

            

        }




        public void EneryaPersonelBazliCagri(string TARIH)
        {
            try
            {
                var client = new RestClient("http://egpoq00.enr.local:50000/XISOAPAdapter/MessageServlet?senderParty=&senderService=COMDATA&receiverParty=&receiverService=&interface=SI_PersonelBazliCagri_OB&interfaceNamespace=http://energaz.com/CRM/PersonelBazliCagri");
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", "Basic UkZDVVNFUjphQTEyMzQ1Ng==");
                request.AddHeader("Content-Type", "application/xml");
                request.AddHeader("Cookie", "JSESSIONID=gfD4VB8hFbDmB37J9UAFfgHu2lEYhAESj4YA_SAPtsGI_JursRtQlDQUvDagFQd_; JSESSIONMARKID=Ke-yhAuYRVrg00VVkwJbaD79fJkLxs7aVwXRKPhgA; saplb_*=(J2EE8818420)8818450");

                var body =
                            "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:urn=\"urn:sap-com:document:sap:rfc:functions\">" +
                            "<soapenv:Header/>" +
                            "<soapenv:Body>" +
                            "<urn:ZCRM_PERSONEL_BAZLI_CAGRI_WS>" +
                            "<!--You may enter the following 2 items in any order-->" +
                            "<!--You may enter the following 2 items in any order-->" +
                            "<!--Optional:-->" +
                            "<IV_KANAL></IV_KANAL>" +
                            "<IV_KANAL></IV_KANAL>" +
                            "<IV_TARIH>" + TARIH + "</IV_TARIH>" +
                            "</urn:ZCRM_PERSONEL_BAZLI_CAGRI_WS>" +
                            "</soapenv:Body>" +
                            "</soapenv:Envelope>"
                            ;
                request.AddParameter("application/xml", body, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                List<Enerya_Personel_Bazli_Cagri> obList = new List<Enerya_Personel_Bazli_Cagri>();
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(response.Content);
                XmlNodeList elemList = doc.GetElementsByTagName("item");


                foreach (XmlNode chldNode in elemList)
                {
                    obList.Add(new Enerya_Personel_Bazli_Cagri
                    {
                        //Id = int.Parse(chldNode.InnerText),
                        ISLEM_NO = chldNode.ChildNodes[0].InnerText,
                        TARIH = DateTime.Parse(chldNode.ChildNodes[1].InnerText),
                        SAAT = chldNode.ChildNodes[2].InnerText,
                        GRUP = chldNode.ChildNodes[3].InnerText,
                        AGRUP = chldNode.ChildNodes[4].InnerText,
                        ORGANIZASYON = chldNode.ChildNodes[5].InnerText,
                        KULLANICI = chldNode.ChildNodes[6].InnerText,
                        MUHATAP = chldNode.ChildNodes[7].InnerText,
                        AD_SOYAD = chldNode.ChildNodes[8].InnerText,
                        TELEFON = chldNode.ChildNodes[9].InnerText,
                        CAGRI_TEL = chldNode.ChildNodes[10].InnerText,


                    });

                    personelBazliCagriList.AddRange(obList);

                    var anketSorguResp = (from a in obList
                                          select new Enerya_Personel_Bazli_Cagri()
                                          {
                                              ISLEM_NO = a.ISLEM_NO,
                                              TARIH = Convert.ToDateTime(a.TARIH.ToString("dd.MM.yyyy")),
                                              SAAT = a.SAAT,
                                              GRUP = a.GRUP,
                                              AGRUP = a.AGRUP,
                                              ORGANIZASYON = a.ORGANIZASYON,
                                              KULLANICI = a.KULLANICI,
                                              MUHATAP = a.MUHATAP,
                                              AD_SOYAD = a.AD_SOYAD,
                                              TELEFON = a.TELEFON,
                                              CAGRI_TEL = a.CAGRI_TEL,
                                          }).ToList();



                    ViewBag.rest = anketSorguResp.DistinctBy(x => x.TELEFON);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("HATA" + ex.Message);
                ViewBag.HATA = "HATA" + ex.Message;
            }
        }



        
        public ActionResult AnketSorgu(string TARIH)
        {
            if (Session["kullaniciDurum"] != null)
            {
                try
                {
                    if (TARIH != null)
                    {

                        EneryaAnketAramaFunc(TARIH);
                        return View();
                    }
                    else
                    {

                        EneryaAnketAramaFunc(TARIH);
                        return View();
                    }
                }
                catch (Exception)
                {

                    return View();
                }
            }
            else
            {
                TempData["InfoMessage"] = "Giriş yapmadan sayfalar arası geçiş yapamazsın !!";
                return RedirectToAction("personelKayitLogin", "Home");
            }



        }

        public ActionResult PersonelBazliCagri(string TARIH)
        {
            if (Session["kullaniciDurum"] != null)
            {

                try
                {
                    if (TARIH != null)
                    {

                        EneryaPersonelBazliCagri(TARIH);
                        return View();
                    }
                    else
                    {

                        EneryaPersonelBazliCagri(TARIH);
                        return View();
                    }
                }
                catch (Exception)
                {
                    return View();
                }

            }
            else
            {
                TempData["InfoMessage"] = "Giriş yapmadan sayfalar arası geçiş yapamazsın !!";
                return RedirectToAction("personelKayitLogin", "Home");
            }
        }


        [HttpPost]
        public ActionResult ExportToCSV(string TARIH)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("anketAramaMusteri");
                    var currentRow = 1;
                    worksheet.Cell(currentRow, 1).Value = "Müsteri Adı";
                    worksheet.Cell(currentRow, 2).Value = "Adresi";
                    worksheet.Cell(currentRow, 3).Value = "İlçe";
                    worksheet.Cell(currentRow, 4).Value = "Şehir";
                    worksheet.Cell(currentRow, 5).Value = "Telefon 1";
                    worksheet.Cell(currentRow, 6).Value = "Telefon 2";
                    worksheet.Cell(currentRow, 7).Value = "Faks";
                    worksheet.Cell(currentRow, 8).Value = "GSM";
                    worksheet.Cell(currentRow, 9).Value = "Yetkili adı";
                    worksheet.Cell(currentRow, 10).Value = "Bilgi 1";
                    worksheet.Cell(currentRow, 11).Value = "Bilgi 2";
                    worksheet.Cell(currentRow, 12).Value = "Bilgi 3";
                    worksheet.Cell(currentRow, 13).Value = "Bilgi 4";



                    //TARİH AYARI EXCELE ÇEVİRİLMEK ÜZERE AYARLANACKA

                    EneryaAnketAramaFunc(TARIH);

                    foreach (var item in anketList.DistinctBy(x => x.CEPTEL))
                    {
                        currentRow++;
                        worksheet.Cell(currentRow, 1).Value = item.MUSADSOYAD;
                        if (item.CEPTEL.StartsWith("+9"))
                        {
                            worksheet.Cell(currentRow, 5).SetValue(item.CEPTEL.Substring(2));
                        }
                        else
                        {
                            worksheet.Cell(currentRow, 5).SetValue(item.CEPTEL);
                        }
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheet.sheet", "anketSorguMusteri.xls");
                    }
                }
            }
            catch
            {
                return null;
            }


        }



        public ActionResult PersonelBazliExportToCSV(string TARIH)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("personelBazliCagri");
                    var currentRow = 1;
                    worksheet.Cell(currentRow, 1).Value = "Müsteri Adı";
                    worksheet.Cell(currentRow, 2).Value = "Adresi";
                    worksheet.Cell(currentRow, 3).Value = "İlçe";
                    worksheet.Cell(currentRow, 4).Value = "Şehir";
                    worksheet.Cell(currentRow, 5).Value = "Telefon 1";
                    worksheet.Cell(currentRow, 6).Value = "Telefon 2";
                    worksheet.Cell(currentRow, 7).Value = "Faks";
                    worksheet.Cell(currentRow, 8).Value = "GSM";
                    worksheet.Cell(currentRow, 9).Value = "Yetkili adı";
                    worksheet.Cell(currentRow, 10).Value = "Bilgi 1";
                    worksheet.Cell(currentRow, 11).Value = "Bilgi 2";
                    worksheet.Cell(currentRow, 12).Value = "Bilgi 3";
                    worksheet.Cell(currentRow, 13).Value = "Bilgi 4";



                    //TARİH AYARI EXCELE ÇEVİRİLMEK ÜZERE AYARLANACKA

                    EneryaPersonelBazliCagri(TARIH);

                    foreach (var item in personelBazliCagriList.DistinctBy(x => x.TELEFON))
                    {
                        currentRow++;
                        worksheet.Cell(currentRow, 1).Value = item.AD_SOYAD;

                        if (item.TELEFON.StartsWith("+9"))
                        {
                            worksheet.Cell(currentRow, 5).SetValue(item.TELEFON.Substring(2));
                        }

                        else if (!item.TELEFON.StartsWith("0"))
                        {
                            worksheet.Cell(currentRow, 5).SetValue("0" + item.TELEFON);
                        }
                        else
                        {
                            worksheet.Cell(currentRow, 5).SetValue(item.TELEFON);
                        }

                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheet.sheet", "personelBazliCagri.xls");
                    }
                }

            }
            catch (Exception)
            {

                return null;
            }
           }


        public ActionResult Index()
        {
            return View();
        }



        public ActionResult emptExcelFile()
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("anketAramaMusteri");
                    var currentRow = 1;
                    worksheet.Cell(currentRow, 1).Value = "Müsteri Adı";
                    worksheet.Cell(currentRow, 2).Value = "Adresi";
                    worksheet.Cell(currentRow, 3).Value = "İlçe";
                    worksheet.Cell(currentRow, 4).Value = "Şehir";
                    worksheet.Cell(currentRow, 5).Value = "Telefon 1";
                    worksheet.Cell(currentRow, 6).Value = "Telefon 2";
                    worksheet.Cell(currentRow, 7).Value = "Faks";
                    worksheet.Cell(currentRow, 8).Value = "GSM";
                    worksheet.Cell(currentRow, 9).Value = "Yetkili adı";
                    worksheet.Cell(currentRow, 10).Value = "Bilgi 1";
                    worksheet.Cell(currentRow, 11).Value = "Bilgi 2";
                    worksheet.Cell(currentRow, 12).Value = "Bilgi 3";
                    worksheet.Cell(currentRow, 13).Value = "Bilgi 4";



                    //TARİH AYARI EXCELE ÇEVİRİLMEK ÜZERE AYARLANACKA


                    foreach (var item in anketList.DistinctBy(x => x.CEPTEL))
                    {
                        currentRow++;
                        worksheet.Cell(currentRow, 1).Value = item.MUSADSOYAD;
                        if (item.CEPTEL.StartsWith("+9"))
                        {
                            worksheet.Cell(currentRow, 5).SetValue(item.CEPTEL.Substring(2));
                        }
                        else
                        {
                            worksheet.Cell(currentRow, 5).SetValue(item.CEPTEL);
                        }
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheet.sheet", "anketSorguMusteri.xls");
                    }
                }

            }
            catch (Exception)
            {

                return null;
            }
       }




        public ActionResult emptExcelFilePersonel()
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("personelBazliCagri");
                    var currentRow = 1;
                    worksheet.Cell(currentRow, 1).Value = "Müsteri Adı";
                    worksheet.Cell(currentRow, 2).Value = "Adresi";
                    worksheet.Cell(currentRow, 3).Value = "İlçe";
                    worksheet.Cell(currentRow, 4).Value = "Şehir";
                    worksheet.Cell(currentRow, 5).Value = "Telefon 1";
                    worksheet.Cell(currentRow, 6).Value = "Telefon 2";
                    worksheet.Cell(currentRow, 7).Value = "Faks";
                    worksheet.Cell(currentRow, 8).Value = "GSM";
                    worksheet.Cell(currentRow, 9).Value = "Yetkili adı";
                    worksheet.Cell(currentRow, 10).Value = "Bilgi 1";
                    worksheet.Cell(currentRow, 11).Value = "Bilgi 2";
                    worksheet.Cell(currentRow, 12).Value = "Bilgi 3";
                    worksheet.Cell(currentRow, 13).Value = "Bilgi 4";



                    //TARİH AYARI EXCELE ÇEVİRİLMEK ÜZERE AYARLANACKA


                    foreach (var item in anketList.DistinctBy(x => x.CEPTEL))
                    {
                        currentRow++;
                        worksheet.Cell(currentRow, 1).Value = item.MUSADSOYAD;
                        if (item.CEPTEL.StartsWith("+9"))
                        {
                            worksheet.Cell(currentRow, 5).SetValue(item.CEPTEL.Substring(2));
                        }
                        else
                        {
                            worksheet.Cell(currentRow, 5).SetValue(item.CEPTEL);
                        }
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheet.sheet", "personelBazliCagri.xls");
                    }
                }

            }
            catch (Exception)
            {

                return null;
            }
           }


    }
}