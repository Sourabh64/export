using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DewsPdfPlugIn
{
    public class ExportToPdf: IDewsExport
    {
        public void Export(IDictionary<string, string> projectDetails, IDictionary<string, string> Metrics, IDictionary<string, Dictionary<string, string>> projectMetricValues, string filepath)
        {
            //string filepath = @"D:\output.pdf";
            string imgpath = @"D:\siemens.jpg";

            System.IO.FileStream fs = new FileStream(filepath, FileMode.Create, FileAccess.Write, FileShare.None);
                Document document = new Document();
            document.SetPageSize(iTextSharp.text.PageSize.A4);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();

            var hd = new PdfPTable(3);
            
            hd.SetWidths(new float[] { 30f, 30f, 30f });
            hd.WidthPercentage = 100;
            iTextSharp.text.Font hf1 = FontFactory.GetFont("Times Roman", 10, BaseColor.BLACK);
            iTextSharp.text.Font hf = FontFactory.GetFont("Times Roman", 6, BaseColor.BLACK);
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imgpath);
            jpg.ScaleToFit(80f, 80f);

            PdfPCell hdcell1 = new PdfPCell();
            hdcell1.Border = 0;
            hdcell1.AddElement(jpg);
            hdcell1.PaddingTop = 5f;
            hd.AddCell(hdcell1);

            PdfPCell hdcell2 = new PdfPCell();
            hdcell2.Border = 0;
            hdcell2.AddElement(new Chunk("DEWS Report",hf1));
            hdcell2.PaddingLeft = 50f;
            hd.AddCell(hdcell2);

            DateTime dt = DateTime.Now;
            PdfPCell hdcell3 = new PdfPCell();
            hdcell3.Border = 0;
            hdcell3.AddElement(new Chunk(dt.ToString(),hf));
            hdcell3.HorizontalAlignment = Element.ALIGN_RIGHT;
            hdcell3.PaddingLeft = 115f;
            hdcell3.PaddingTop = 10f;
            hd.AddCell(hdcell3);
            document.Add(hd);

            iTextSharp.text.Font tf = FontFactory.GetFont("Times Roman", 6, BaseColor.BLACK);
            iTextSharp.text.Font vf = FontFactory.GetFont("Times Roman", 5, BaseColor.BLACK);
    
            //details of the project
            var pt = new PdfPTable(projectDetails.Count);
            pt.SetWidths(new float[] { 250f, 250f, 250f, 250f });
            pt.WidthPercentage = 100;
           // pt.SpacingBefore=1f;
            var tcells = new List<PdfPCell>();
            var vcells = new List<PdfPCell>();

            foreach (var pi in projectDetails)
            {
                var tcell = new PdfPCell();
                tcell.BackgroundColor = new iTextSharp.text.Color(150, 150, 150);
                tcell.AddElement(new Chunk(pi.Key, tf));
                tcell.PaddingLeft = 45f;
                tcell.FixedHeight = 15f;
                tcells.Add(tcell);

                var vcell = new PdfPCell();
                if (pi.Value == " ")
                {
                    vcell.BackgroundColor = new iTextSharp.text.Color(0, 128, 0);
                    vcell.AddElement(new Chunk(pi.Value, vf));
                    vcell.FixedHeight = 15f;
                    vcells.Add(vcell);
                }
                else
                {
                    vcell.BackgroundColor = new iTextSharp.text.Color(150, 150, 150);
                    vcell.AddElement(new Chunk(pi.Value, vf));
                    vcell.PaddingLeft = 55f;
                    vcell.FixedHeight = 15f;
                    vcells.Add(vcell);
                }
            }

            var pRow = new PdfPRow(tcells.ToArray());
            var vRow = new PdfPRow(vcells.ToArray());

            pt.Rows.Add(pRow);
            pt.Rows.Add(vRow);
            document.Add(pt);

            iTextSharp.text.Font tf1 = FontFactory.GetFont("Times Roman", 6, BaseColor.BLACK);
            iTextSharp.text.Font vf1 = FontFactory.GetFont("Times Roman", 5, BaseColor.BLACK);

            //goal
            var mt = new PdfPTable(Metrics.Count + 1);
            mt.SetWidths(new float[] { 80f, 30f, 30f, 30f, 30f, 30f, 30f, 30f, 30f, 30f, 160f, 160f });
            mt.WidthPercentage = 100;

            var mcells = new List<PdfPCell>();
            var mvcells = new List<PdfPCell>();

            var mcell1 = new PdfPCell();
            mcell1.BackgroundColor = new iTextSharp.text.Color(169, 169, 169);
            mcell1.AddElement(new Chunk(" ", tf1));
            mcell1.HorizontalAlignment = Element.ALIGN_MIDDLE;
            mcell1.VerticalAlignment = Element.ALIGN_MIDDLE;
            mcell1.Rotation = -270;
            mcells.Add(mcell1);

            var mvcell1 = new PdfPCell();
            mvcell1.BackgroundColor = new iTextSharp.text.Color(169, 169, 169);
            mvcell1.AddElement(new Chunk("Goal", tf1));
            mvcell1.PaddingLeft = 5f;
            mvcells.Add(mvcell1);


            foreach (var met in Metrics)
            {
                var mcell = new PdfPCell();
                if (met.Key == "High Risks" || met.Key == "Comments")
                {
                    mcell.BackgroundColor = new iTextSharp.text.Color(169, 169, 169);
                    mcell.AddElement(new Chunk(met.Key, tf1));
                    mcell.PaddingLeft = 50f;
                    mcell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                    mcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    //mcell.FixedHeight = 20f;
                    mcells.Add(mcell);
                }
                else
                {

                    mcell.BackgroundColor = new iTextSharp.text.Color(169, 169, 169);
                    mcell.AddElement(new Chunk(met.Key, tf1));
                    mcell.Rotation = -270;
                    mcell.PaddingLeft = 5f;
                    mcells.Add(mcell);
                }

                var mvcell = new PdfPCell();
                mvcell.BackgroundColor = new iTextSharp.text.Color(190, 190, 190);
                mvcell.AddElement(new Chunk(met.Value, vf1));
                mvcell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                mvcell.VerticalAlignment = Element.ALIGN_MIDDLE;
                mvcell.FixedHeight = 15f;
                mvcell.PaddingLeft = 4f;
                //mvcell.PaddingTop = 5f;
                mvcells.Add(mvcell);
            }

            var mRow = new PdfPRow(mcells.ToArray());
            var mvRow = new PdfPRow(mvcells.ToArray());

            mt.Rows.Add(mRow);
            mt.Rows.Add(mvRow);

            //values
            foreach (var pm in projectMetricValues)
            {
                var pmcells = new List<PdfPCell>();

                var pmcell = new PdfPCell();
                pmcell.BackgroundColor = new iTextSharp.text.Color(169, 169, 169);
                pmcell.AddElement(new Chunk(pm.Key, tf1));
                pmcell.HorizontalAlignment = Element.ALIGN_CENTER;
                pmcell.VerticalAlignment = Element.ALIGN_CENTER;
                pmcell.PaddingLeft = 5f;
                // pmcell.FixedHeight = 15f;
                pmcells.Add(pmcell);

                foreach (var val in pm.Value)
                {
                    var pmval = new PdfPCell();
                    pmval.BackgroundColor = new iTextSharp.text.Color(255, 255, 255);
                    pmval.AddElement(new Chunk(val.Value, vf1));
                    pmval.HorizontalAlignment = Element.ALIGN_MIDDLE;
                    pmval.VerticalAlignment = Element.ALIGN_MIDDLE;
                    pmval.FixedHeight = 15f;
                    pmval.PaddingLeft = 10f;
                    pmcells.Add(pmval);
                }

                var pmRow = new PdfPRow(pmcells.ToArray());
                mt.Rows.Add(pmRow);
            }
 
            document.Add(mt);
            document.Close();
        }
    }
}
