using AutomacaoGoogleAcademico.BLL;
using AutomacaoGoogleAcademicoDB.DAL;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScieloEzequiel.BLL
{
    class Automacao
    {
        private string Caminho;
        private string sPesquisa;
        private string dDataPesq;
        private string Controle;
        public bool ParaThread;
        private AccessDB DB = new AccessDB();

        public Automacao(string Caminho_, string sPesquisa_, string Controle_, AccessDB DB_)
        {
            DB = DB_;
            Caminho = Caminho_;
            sPesquisa = sPesquisa_;
            Controle = Controle_;
            dDataPesq = DateTime.Now.ToString("dd/mm/yyy hh:mm:ss");
        }
        public void PreparaPesquisa()
        {

            int ControlePast = 0;
            int P = 1;
            int ControlePaginas = 0;
            int QtdPesquisas = 1;
            string AuxPesq = "";
            string CaminhoPesq = "";
            string Url = "";
            string ControlePesq = "";
            Boolean MaisPags = false;
            IWebElement element1 = null;
            IWebDriver driver = null;


            try
            {
                ChromeOptions options = new ChromeOptions();

                //=================================================================================
                //Configura pasta de download
                //=================================================================================
                options.AddUserProfilePreference("download.default_directory", Caminho);

                options.AddUserProfilePreference("download.prompt_for_download", false);

                options.AddUserProfilePreference("download.directory_upgrade", true);

                options.AddUserProfilePreference("plugins.plugins_disabled", "Chrome PDF Viewer");

                options.AddUserProfilePreference("plugins.always_open_pdf_externally", true);
                //=================================================================================
                //=================================================================================

                driver = new ChromeDriver(options); // instância uma nova sessão do ChromeDriver navegador usado para testes unitários 

                IniciaPesquisaScielo(driver);

                ControlePaginas = ObtemNumeroPags(driver);

                for (int i = 0; i < ControlePaginas; i++)
                {
                    var Pesquisas = driver.FindElements(By.XPath(("//td[@width='485']/table/tbody/tr/td[2]/div/a[2]"))); //Pega os elemento referente aos links exibidos na pagina

                    QtdPesquisas = 0;
                    QtdPesquisas = Pesquisas.Count();

                    for (int j = 0; j < QtdPesquisas - 1; j++)
                    {
                        element1 = Pesquisas[j];

                        if (element1.Text == "text in portuguese")
                        {

                            element1.Click();

                            CapturaPesquisa(driver);

                            ControlePesq = ConsultaCodPesquisa();

                            CaminhoPesq = CriaPastaPesquisa(ControlePesq, Caminho, sPesquisa);// Cria uma pasta para cada link dentro de uma pasta raiz para o titulo da pesquisa

                            DownloadRefs(driver, CaminhoPesq);

                            driver.SwitchTo().Window(driver.WindowHandles[0]);

                            driver.Navigate().Back();
                            //IniciaPesquisaScielo(driver);

                            Pesquisas = driver.FindElements(By.XPath(("//td[@width='485']/table/tbody/tr/td[2]/div/a[2]"))); //Pega os elemento referente aos links exibidos na pagina

                        }
                    }

                    P++;

                    if (ControlePaginas > 1)
                    {
                        driver.FindElement(By.XPath("//input[@name='Page" + P + "']")).Click();
                    }

                }



            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + ", Classe: Automacao, Método: PreparaPesquisa");
            }
            finally
            {
                driver.Close(); // fecha pagina do ChromeDriver
                driver.Quit();  // fecha prompt do ChromeDriver
            }
        }

        private string ConsultaCodPesquisa()
        {
            string Qry;
            OleDbDataReader DR;

            try
            {
                Qry = "select top 1 nCodPesquisa  from TbPesquisas order by nCodPesquisa desc";

                using (DR = DB.DR(Qry))
                {
                    if (DR.HasRows == true)
                    {
                        DR.Read();

                        return DR["nCodPesquisa"].ToString();
                    }
                    else
                    {
                        return "erro";
                    }

                }

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: ConsultaCodPesquisa");
            }
        }

        public string CriaPastaPesquisa(string Controle, string Caminho, string sPesquisa)
        {
            try
            {
                Directory.CreateDirectory(Caminho + "\\" + Controle + "." + sPesquisa);

                return Caminho + "\\" + Controle + "." + sPesquisa;
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: CriaPastaPesquisa");
            }
        }

        private void IniciaPesquisaScielo(IWebDriver driver)
        {
            try
            {

                string[] Pesquisa = sPesquisa.Split(' ');

                driver.Navigate().GoToUrl("http://www.scielo.br/cgi-bin/wxis.exe/iah/?IsisScript=iah/iah.xis&base=article%5Edlibrary&format=iso.pft&lang=i"); //Navega na url especificada
                EsperaElemento("config", driver);

                driver.FindElement(By.XPath("//table[@align='center']/tbody/tr[2]/td[3]/input")).SendKeys(Pesquisa[0].ToString()); // Efetua pesquisa informada pelo usuario

                if (Pesquisa.Count() > 1)
                {
                    driver.FindElement(By.XPath("//table[@align='center']/tbody/tr[3]/td[3]/input")).SendKeys(Pesquisa[1].ToString()); // Efetua pesquisa informada pelo usuario
                }

                driver.FindElement(By.XPath("//input[@src='/iah/I/image/pesq.gif']")).Click();
                EsperaElemento("References", driver);

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: IniciaPesquisaScielo");
            }
        }

        private void CapturaPesquisa(IWebDriver driver)
        {
            string Resumo = "";
            string LinkPDF = "";
            string LinkSub = "";
            string Qry = "";
            int idx = 0;
            IWebElement element1 = null;

            try
            {


                Resumo = driver.FindElement(By.XPath("//div[@class='content']")).Text;

                
                Resumo = Resumo.Replace("RESUMO", "₱ RESUMO ₱");
                Resumo = Resumo.Replace("INTRODUÇÃO", "₱ INTRODUÇÃO ₱");
                Resumo = Resumo.Replace("Palavras-chave", "₱ Palavras-chave ₱");
                Resumo = Resumo.Replace("ARTICLES", "₱ ARTICLES ₱");
                Resumo = Resumo.Replace("SUMMARY", "₱ SUMMARY ₱");
                Resumo = Resumo.Replace("ABSTRACT", "₱ ABSTRACT ₱");
                Resumo = Resumo.Replace("INTRODUCTION", "₱ INTRODUCTION ₱");
                Resumo = Resumo.Replace("RESUMEN", "₱ RESUMEN ₱");

                string[] VetResumo = Resumo.Split('₱');

                for (int i = 0; i < VetResumo.Count(); i++)
                {
                    if (VetResumo[i].ToString().Trim() == "RESUMO")
                    {
                        Resumo = VetResumo[i + 1].ToString().Replace("\r", " ");
                        break;
                    }
                }

                LinkPDF = driver.Url.ToString();


                driver.FindElement(By.XPath("//a[@style='text-decoration: none;']")).Click();

                driver.SwitchTo().Window(driver.WindowHandles[1]);

                EsperaElemento("webServices", driver);

                LinkSub = driver.Url.ToString();

                Qry = " Insert into TbPesquisas (sResumo,sLinkPDF,sLinkSubPag,sPesquisa)  " +
                      " values('" + Resumo + "', '" + LinkPDF.Replace("'", "1!2@3#4$") + "' " +
                      " , '" + LinkSub.Replace("'", "1!2@3#4$") + "', '" + sPesquisa + "')";

                DB.ExecutaQry(Qry);

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: CapturaPesquisa");
            }
        }

        private void DownloadRefs(IWebDriver driver, string CaminhoPesq)
        {
            Arquivo Arq = new Arquivo();

            try
            {

                Thread.Sleep(500);
                driver.FindElement(By.XPath("//div[@class='webServices']/ul/li/a")).Click();//Export to BibTex
                EsperaDownload("*.bib", driver);
                Arq.MoveArq(Caminho, CaminhoPesq, "bib");
                driver.FindElement(By.XPath("//div[@class='webServices']/ul/li[2]/a")).Click(); //Export to Reference Manager
                EsperaDownload("*.ris", driver);
                Arq.MoveArq(Caminho, CaminhoPesq, "ris");
                driver.FindElement(By.XPath("//div[@class='webServices']/ul/li[3]/a")).Click(); //Export to Pro Cite
                EsperaDownload("*.txt", driver);
                Arq.MoveArq(Caminho, CaminhoPesq, "txt");
                driver.FindElement(By.XPath("//div[@class='webServices']/ul/li[4]/a")).Click(); //Export to End Note
                EsperaDownload("*.enw", driver);
                Arq.MoveArq(Caminho, CaminhoPesq, "enw");
                //driver.FindElement(By.XPath("//div[@class='webServices']/ul/li[5]")).Click(); //Export to Refworks
                //EsperaDownload("*.ref", driver);

                driver.Close();

                FechaAbasSecundarias(driver);

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: DownloadRefs");
            }
        }

        private void FechaAbasSecundarias(IWebDriver driver)
        {
            try
            {
                List<String> tabs = new List<String>(driver.WindowHandles);

                for (int i = 1; i < tabs.Count; i++)
                {
                    driver.SwitchTo().Window(driver.WindowHandles[i]);

                    driver.Close();
                }

                driver.SwitchTo().Window(driver.WindowHandles[0]);

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: FechaAbasSecundarias");
            }
        }

        private int ObtemNumeroPags(IWebDriver driver)
        {
            string aux = "";
            int NumeroPags = 0;
            string[] Vet;

            try
            {
                aux = driver.FindElement(By.XPath("//table[@width='600']/tbody/tr/td/font/sup/b/i")).Text;

                Vet = aux.Split(' ');

                NumeroPags = Int32.Parse(Vet[3]);

                return NumeroPags;

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: ObtemNumeroPag");
            }
        }



        public string EsperaElemento(string Elemento, IWebDriver driver)
        {
            int qtd = 0;
            int TimeOut = 0;
            string Bib = "";

            try
            {

                while (qtd == 0)
                {
                    Thread.Sleep(500);

                    if (driver.PageSource.Contains(Elemento) == true)
                    {
                        return "ok";
                    }
                    else if (TimeOut == 120)
                    {
                        throw new Exception("Tempo de espera excedido, tela inesperada ou problemas com o navegador.");
                    }

                    TimeOut++;
                }

                return "ok";

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: EsperaElemento");
            }
        }

        public string EsperaDownload(string TipoArq, IWebDriver driver)
        {
            int qtd = 0;
            int TimeOut = 0;

            try
            {
                var PDF = new DirectoryInfo(Caminho).GetFiles(TipoArq);//"*.pdf"

                while (qtd == 0)
                {


                    PDF = new DirectoryInfo(Caminho).GetFiles(TipoArq);

                    qtd = PDF.Length;

                    Thread.Sleep(1500);

                    if (TimeOut == 120)
                    {
                        //if (DialogResult.Yes == MessageBox.Show("O download do PDF está demorando ou ocorreu algum erro, deseja aguardar?", "", MessageBoxButtons.YesNo))
                        //{
                        //    TimeOut = 0;
                        //}
                        //else
                        //{
                        return "erro";
                        //}

                    }

                    TimeOut++;
                }


                return "ok";

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: EsperaDownload");
            }
        }
    }
}
