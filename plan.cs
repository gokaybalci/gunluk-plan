using System;
using System.IO;
using System.Net;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;

namespace GunlukPlanGUI
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void download_daily_plan()
        {
            string kademe = kademe_entry.Text;
            string hafta = hafta_entry.Text;
            string oisim = oisim_entry.Text;
            string misim = misim_entry.Text;

            // Gives an error if there is a missing input
            if (kademe == "" || hafta == "" || oisim == "" || misim == "")
            {
                MessageBox.Show("Lütfen eksik alanları doldurunuz.", "Hata!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string kademe_url = "https://www.ingilizceciyiz.com/" + kademe + "-sinif-ingilizce-gunluk-plan/";

            // Requests URL and get response object
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(kademe_url);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            // Parse text obtained
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.Load(response.GetResponseStream());

            string xpath = $"//*[contains(text(), '{hafta}. Hafta') and not(.//a)]";
            HtmlNodeCollection eslesen_haftalar = htmlDoc.DocumentNode.SelectNodes(xpath);

            if (eslesen_haftalar != null)
            {
                foreach (HtmlNode link in eslesen_haftalar)
                {
                    if (link.InnerHtml.Contains(".docx"))
                    {
                        HttpWebRequest docxRequest = (HttpWebRequest)WebRequest.Create(link.InnerHtml);
                        HttpWebResponse docxResponse = (HttpWebResponse)docxRequest.GetResponse();
                        Stream docxStream = docxResponse.GetResponseStream();
                        string docxFilePath = kademe + " Sınıf " + hafta + ". Hafta.docx";
                        using (FileStream fileStream = new FileStream(docxFilePath, FileMode.Create, FileAccess.Write))
                        {
                            docxStream.CopyTo(fileStream);
                        }
                        break; // exit the loop after the first valid link is downloaded
                    }
                }
            }

            // Load the document and change teacher and principal name
            string documentPath = kademe + " Sınıf " + hafta + ". Hafta.docx";
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(documentPath, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                Body body = mainPart.Document.Body;

                int dotCount = 0; // Counter for number of occurrences of 'dots' in given documents
                foreach (Paragraph para in body.Elements<Paragraph>())
                {
                    string paraText = para.InnerText;
                    if (paraText.Contains("…"))
                    {
                        dotCount++;
                        if (dotCount == 1)
                        {
                            para.InnerXml = para.InnerXml.Replace("…", oisim);
                        }
                        else if (dotCount == 2)
                        {
                            para.InnerXml = para.InnerXml.Replace("…", misim);
                        }
                    }
                }
                mainPart.Document.Save();
            }

            status_label.Text = "Dosyanız hazır.";
        }

        private void execute_button_Click(object sender, EventArgs e)
        {
            download_daily_plan();
        }
    }
}
