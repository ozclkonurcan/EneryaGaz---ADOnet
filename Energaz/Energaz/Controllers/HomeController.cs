using Energaz.Models;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Energaz.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            if (Session["kullaniciDurum"] != null)
            {
                try
                {
                    return View();

                }
                catch (Exception ex)
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

        public ActionResult About()
        {
            if (Session["kullaniciDurum"] != null)
            {
                try
                {
                    return View();

                }
                catch (Exception ex)
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

        public ActionResult Contact()
        {
            if (Session["kullaniciDurum"] != null)
            {
                try
                {
                    return View();

                }
                catch (Exception ex)
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




        //LOGİN CONTROLLER
        public ActionResult personelKayitLogin(YetkiliPersonel yetkiliPersonel)
        {
            try
            {

                NameValueCollection section =  (NameValueCollection)ConfigurationManager.GetSection("MyDictionary");






                List<YetkiliPersonel> personelGiris = new List<YetkiliPersonel>()
            {
                new YetkiliPersonel(){id=1,k_ad="corumgaz",k_sifre="crm_20220609",k_durum="Yetkili"},
                new YetkiliPersonel(){id=2,k_ad="ozclkonur",k_sifre="onr_159456",k_durum="Yetkili"},
                new YetkiliPersonel(){id=3,k_ad="Anketadmin",k_sifre="159456.a0",k_durum="Yetkili"},

            };



                var den = section[yetkiliPersonel.k_ad];
                if (den != null && den == yetkiliPersonel.k_sifre)
                {




                    var resp = (from a in personelGiris select a).ToList();

                    var giris = resp.Any(x => x.k_ad.Trim() == yetkiliPersonel.k_ad.Trim() && x.k_sifre.Trim() == yetkiliPersonel.k_sifre.Trim());

                    if (giris)
                    {
                        var respData = (from a in personelGiris where a.k_ad.Trim() == yetkiliPersonel.k_ad.Trim() && a.k_sifre.Trim() == yetkiliPersonel.k_sifre.Trim() select a).FirstOrDefault();

                        Session["K_ID"] = respData.id;
                        Session["kullaniciAdi"] = respData.k_ad;
                        Session["kullaniciDurum"] = respData.k_durum;


                        TempData["SuccessMessage"] = "Giriş Başarılı";
                        return RedirectToAction("AnketSorgu", "Enerya");
                    }
                    else
                    {
                        TempData["ErrorMessage"] = "Giriş Başarısız";
                        return RedirectToAction("");
                    }

                }
                else
                {

                    return View();
                }


            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = "Beklenmeyen bir hata meydana geldi tekrar giriş yapmayı deneyin \n" + ex.Message;
                return RedirectToAction("");
            }
        }

        public ActionResult logout(YetkiliPersonel yetkiliPersonel)
        {
            try
            {

                int K_ID = (int)Session["K_ID"];
                Session.Abandon();
                TempData["InfoMessage"] = "Çıkış başarılı";
                return RedirectToAction("personelKayitLogin", "Home");

            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = "Beklenmeyen bir hata meydana geldi tekrar giriş yapmayı deneyin \n" + ex.Message;
                return RedirectToAction("");
            }
        }


        //LOGİN CONTROLLER

    }
}