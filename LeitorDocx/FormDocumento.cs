using Microsoft.Office.Interop.Word;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Collections.Specialized.BitVector32;

namespace LeitorDocx
{
    public partial class FormDocumento : Form
    {
        private Microsoft.Office.Interop.Word.Application wordApplication;
        private Document wordDocument;

        public FormDocumento()
        {
            InitializeComponent();
            //InitializeDocument();
        }

        private void InitializeDocument()
        {
            string directory = AppDomain.CurrentDomain.BaseDirectory;
            wordApplication = new Microsoft.Office.Interop.Word.Application();
            wordDocument = wordApplication.Documents.Open(directory + @"\FICHA CADASTRAL.docx");

            btnImprimir.Enabled = true;
            btnPdf.Enabled = true;
        }

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            PreencherDocumento();

            wordDocument.PrintOut();

            MessageBox.Show("Impressão concluída!");

            FinalizarDocumento();
            InitializeDocument();
        }

        private void btnSalvarPDF_Click(object sender, EventArgs e)
        {
            PreencherDocumento();

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Documento do Word|*.pdf";
            saveFileDialog.FileName = "Documento_Preenchido.pdf";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = saveFileDialog.FileName;
                object saveAsPath = savePath;
                object saveFormat = WdSaveFormat.wdFormatPDF;
                wordDocument.SaveAs2(ref saveAsPath, ref saveFormat);

                MessageBox.Show("Documento salvo com sucesso!");

                FinalizarDocumento();
                InitializeDocument();
            }
        }

        private void PreencherDocumento()
        {
            object replaceAll = WdReplace.wdReplaceAll;

            //[NOME]
            string textoNome = txtNome.Text;
            Find findNome = wordApplication.Selection.Find;
            findNome.ClearFormatting();
            findNome.Text = "[NOME]";
            findNome.Replacement.ClearFormatting();
            findNome.Replacement.Text = textoNome;
            findNome.Execute(Replace: ref replaceAll);

            //[CPF]
            string textoCpf = txtCpf.Text;
            Find findCpf = wordApplication.Selection.Find;
            findCpf.ClearFormatting();
            findCpf.Text = "[CPF]";
            findCpf.Replacement.ClearFormatting();
            findCpf.Replacement.Text = textoCpf;
            findCpf.Execute(Replace: ref replaceAll);

            //[RG]
            string textoRg = txtRg.Text;
            Find findRg = wordApplication.Selection.Find;
            findRg.ClearFormatting();
            findRg.Text = "[RG]";
            findRg.Replacement.ClearFormatting();
            findRg.Replacement.Text = textoRg;
            findRg.Execute(Replace: ref replaceAll);

            //[SEXO]
            string textoSexo = txtSexo.Text;
            Find findSexo = wordApplication.Selection.Find;
            findSexo.ClearFormatting();
            findSexo.Text = "[SEXO]";
            findSexo.Replacement.ClearFormatting();
            findSexo.Replacement.Text = textoSexo;
            findSexo.Execute(Replace: ref replaceAll);
        }

        private void FormDocumento_FormClosing(object sender, FormClosingEventArgs e)
        {
            FinalizarDocumento();
        }

        private void FinalizarDocumento()
        {
            //wordDocument.Close(false);
            //wordApplication.Quit();
        }

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            foreach (Control textbox in this.Controls)
            {
                if (textbox is TextBox)
                {
                    ((TextBox)textbox).Text = String.Empty;
                }
                else if (textbox is MaskedTextBox)
                {
                    ((MaskedTextBox)textbox).Text = String.Empty;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region PARAMETROS GLOBAIS DO WORD

            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc";

            #endregion

            #region INSTANCIA DA APLICAÇÃO E CRIAÇÃO DO DOCUMENTO EM BRANCO NA MEMORIA

            Microsoft.Office.Interop.Word._Application oWord;
            Microsoft.Office.Interop.Word._Document oDoc;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);

            #endregion

            #region FLAG QUE DETERMINA SE O PROGRAMA WORD SERÁ ABERTO AO CRIAR O ARQUIVO EM MEMORIA

            //SOMENTE DEIXE TRUE PARA TESTAR. 
            oWord.Visible = true;

            #endregion

            #region MARGEMS DO DOCUMENTO

            oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.5f); //SUPERIOR
            oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.5f); //INFERIOR
            oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.5f); //ESQUERDA
            oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.5f); //DIREITA

            #endregion

            #region ESTILO DA FONTE DO DOCUMENTO

            //CUIDADO, AO MUDAR A FONTE AS TABELAS E PARAGRAFOS PODEM QUEBRAR DE LINHA DEVIDO AO NOVO TAMANHO
            oDoc.Content.Font.Name = "Calibri";

            #endregion

            #region BORDA DO DOCUMENTO

            Borders borders = oDoc.Sections[1].Borders;
            borders.Enable = 1; 
            borders.DistanceFromTop = 30; 
            borders.DistanceFromBottom = 24; 
            borders.DistanceFromLeft = 26; 
            borders.DistanceFromRight = 26; 
            borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle; 
            borders.OutsideLineWidth = WdLineWidth.wdLineWidth300pt;

            #endregion

            #region CORPO DO DOCUMENTO

            #region TÍTULO FICHA CADASTRAL - PESSOA FISICA

            Microsoft.Office.Interop.Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "FICHA CADASTRAL - PESSOA FISICA";
            oPara1.Range.Font.Size = 22;
            oPara1.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceBefore = 40;
            oPara1.Format.SpaceAfter = 24;
            oPara1.Range.InsertParagraphAfter();

            #endregion

            #region 1ª TABELA

            Microsoft.Office.Interop.Word.Table oTable;
            Microsoft.Office.Interop.Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 2, 2, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 0;
            oTable.Range.ParagraphFormat.SpaceBefore = 0;
            oTable.Cell(1, 1).Range.Text = "REPRESENTAÇÃO:";
            oTable.Cell(2, 1).Range.Text = ""; // [REPRESENTAÇÃO]
            oTable.Cell(1, 2).Range.Text = "CNPJ:";
            oTable.Cell(2, 2).Range.Text = "";// [CNPJ]
            oTable.Rows[1].Range.Font.Bold = 1;
            int r, c;
            for (r = 1; r <= 2; r++)
            {
                for (c = 1; c <= 2; c++)
                {
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                }

                oTable.Rows[r].Range.Font.Size = 10;
            }

            #endregion

            #region TÍTULO DADOS PESSOAIS

            Microsoft.Office.Interop.Word.Paragraph oPara2;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara2.Range.Text = "DADOS PESSOAIS";
            oPara2.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            oPara2.Range.Font.Size = 11;
            oPara2.Range.Font.Bold = 1;
            oPara2.Format.SpaceBefore = 20;
            oPara2.Format.SpaceAfter = 6;
            oPara2.Range.InsertParagraphAfter();

            #endregion

            #region 2ª TABELA

            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 8, 4, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceBefore = 0;
            oTable.Range.ParagraphFormat.SpaceAfter = 0;
            oTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            oTable.Columns[1].SetWidth(60, WdRulerStyle.wdAdjustProportional);
            oTable.Columns[3].SetWidth(80, WdRulerStyle.wdAdjustProportional);

            oTable.Cell(1, 1).Range.Text = "NOME:";
            oTable.Cell(1, 2).Range.Text = ""; // [NOME]
            oTable.Cell(2, 1).Range.Text = "CPF:";
            oTable.Cell(2, 2).Range.Text = ""; // [CPF]
            oTable.Cell(3, 1).Range.Text = "SEXO:";
            oTable.Cell(3, 2).Range.Text = ""; // [SEXO]
            oTable.Cell(4, 1).Range.Text = "NATURAL:";
            oTable.Cell(4, 2).Range.Text = ""; // [NATURAL]
            oTable.Cell(5, 1).Range.Text = "EST. CIVIL:";
            oTable.Cell(5, 2).Range.Text = ""; // [EST. CIVIL]
            oTable.Cell(6, 1).Range.Text = "ORGÃO EXP:";
            oTable.Cell(6, 2).Range.Text = ""; // [ORGÃO EXP]
            oTable.Cell(7, 1).Range.Text = "CONTATO:";
            oTable.Cell(7, 2).Range.Text = ""; // [CONTATO]
            oTable.Cell(8, 1).Range.Text = "PROFISSÃO:";
            oTable.Cell(8, 2).Range.Text = ""; // [PROFISSÃO]

            oTable.Cell(2, 3).Range.Text = "RG:";
            oTable.Cell(2, 4).Range.Text = ""; // [RG]
            oTable.Cell(3, 3).Range.Text = "NASCIMENTO:";
            oTable.Cell(3, 4).Range.Text = ""; // [NASCIMENTO]
            oTable.Cell(4, 3).Range.Text = "NACIONALIDADE:";
            oTable.Cell(4, 4).Range.Text = ""; // [NACIONALIDADE]
            oTable.Cell(5, 3).Range.Text = "DATA EXP:";
            oTable.Cell(5, 4).Range.Text = ""; // [DATA EXP]
            oTable.Cell(6, 3).Range.Text = "xxxxxxxxxxxxxx";
            oTable.Cell(6, 4).Range.Text = "xxxxxxxxxxxxxx";
            oTable.Cell(7, 3).Range.Text = "EMAIL:";
            oTable.Cell(7, 4).Range.Text = ""; // [EMAIL]
            oTable.Cell(8, 3).Range.Text = "RENDA:";
            oTable.Cell(8, 4).Range.Text = ""; // [RENDA]

            for (r = 1; r <= 8; r++)
            {
                for (c = 1; c <= 4; c++)
                {
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                }

                oTable.Rows[r].Range.Font.Size = 10;
            }

            oTable.Rows[1].Cells[2].Merge(oTable.Rows[1].Cells[3]);
            oTable.Rows[1].Cells[2].Merge(oTable.Rows[1].Cells[3]);

            #endregion

            #region TÍTULO DADOS RESIDENCIAIS

            Microsoft.Office.Interop.Word.Paragraph oPara3;
            object oRng2 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara3 = oDoc.Content.Paragraphs.Add(ref oRng2);
            oPara3.Range.Text = "DADOS RESIDENCIAIS";
            oPara3.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            oPara3.Range.Font.Size = 11;
            oPara3.Range.Font.Bold = 1;
            oPara3.Format.SpaceBefore = 20;
            oPara3.Format.SpaceAfter = 6;
            oPara3.Range.InsertParagraphAfter();

            #endregion

            #region 3ª TABELA

            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 3, 4, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceBefore = 0;
            oTable.Range.ParagraphFormat.SpaceAfter = 0;
            oTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            oTable.Columns[1].SetWidth(80, WdRulerStyle.wdAdjustProportional);
            oTable.Columns[2].SetWidth(200, WdRulerStyle.wdAdjustProportional);
            oTable.Columns[3].SetWidth(45, WdRulerStyle.wdAdjustProportional);

            oTable.Cell(1, 1).Range.Text = "ENDEREÇO:";
            oTable.Cell(1, 2).Range.Text = ""; // [ENDEREÇO]
            oTable.Cell(2, 1).Range.Text = "COMPLEMENTO:";
            oTable.Cell(2, 2).Range.Text = ""; // [COMPLEMENTO]
            oTable.Cell(3, 1).Range.Text = "CIDADE:";
            oTable.Cell(3, 2).Range.Text = ""; // [CIDADE]

            oTable.Cell(1, 3).Range.Text = "BAIRRO:";
            oTable.Cell(1, 4).Range.Text = ""; // [BAIRRO]
            oTable.Cell(2, 3).Range.Text = "CEP:";
            oTable.Cell(2, 4).Range.Text = ""; // [CEP]
            oTable.Cell(3, 3).Range.Text = "ESTADO:";
            oTable.Cell(3, 4).Range.Text = ""; // [ESTADO]

            for (r = 1; r <= 3; r++)
            {
                for (c = 1; c <= 4; c++)
                {
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                }

                oTable.Rows[r].Range.Font.Size = 10;
            }

            #endregion

            #region TÍTULO DADOS DA PROPOSTA

            Microsoft.Office.Interop.Word.Paragraph oPara4;
            object oRng3 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng3);
            oPara4.Range.Text = "DADOS DA PROPOSTA";
            oPara4.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            oPara4.Range.Font.Size = 11;
            oPara4.Range.Font.Bold = 1;
            oPara4.Format.SpaceBefore = 20;
            oPara4.Format.SpaceAfter = 6;
            oPara4.Range.InsertParagraphAfter();

            #endregion

            #region 4ª TABELA

            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 4, 4, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceBefore = 0;
            oTable.Range.ParagraphFormat.SpaceAfter = 0;
            oTable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            oTable.Columns[1].SetWidth(80, WdRulerStyle.wdAdjustProportional);
            oTable.Columns[3].SetWidth(75, WdRulerStyle.wdAdjustProportional);

            oTable.Cell(1, 1).Range.Text = "TIPO PRODUTO:";
            oTable.Cell(1, 2).Range.Text = ""; // [TIPO PRODUTO]
            oTable.Cell(2, 1).Range.Text = "TABELA:";
            oTable.Cell(2, 2).Range.Text = ""; // [TABELA]
            oTable.Cell(4, 1).Range.Text = "VENDEDOR:";
            oTable.Cell(4, 2).Range.Text = ""; // [VENDEDOR]

            oTable.Cell(1, 3).Range.Text = "CRÉDITO R$:";
            oTable.Cell(1, 4).Range.Text = ""; // [CREDITO  R$]
            oTable.Cell(2, 3).Range.Text = "PARCELA R$:";
            oTable.Cell(2, 4).Range.Text = ""; // [PARCELA R$]
            oTable.Cell(3, 3).Range.Text = "ADESÃO R$:";
            oTable.Cell(3, 4).Range.Text = ""; // [ADESÃO   R$]
            oTable.Cell(4, 3).Range.Text = "TOTAL PAGO R$:";
            oTable.Cell(4, 4).Range.Text = ""; // [TOTAL PAGO R$]

            for (r = 1; r <= 4; r++)
            {
                for (c = 1; c <= 4; c++)
                {
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(r, c).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                }
                oTable.Rows[r].Range.Font.Size = 10;
            }

            #endregion

            #region MARCA D'AGUA

            //string marcaDagua = AppDomain.CurrentDomain.BaseDirectory + @"marca.png";
            if (!File.Exists("marca.png"))
            {
                object obj = Properties.Resources.marca;
                System.Drawing.Bitmap rs = (System.Drawing.Bitmap)(obj);
                rs.Save("marca.png");
            }

            string marcaDagua = Path.GetFullPath("marca.png");
            Microsoft.Office.Interop.Word.Range myRange = oWord.Selection.Range.GoTo(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage, Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage, 1);
            Microsoft.Office.Interop.Word.Shape myShape = oDoc.Shapes.AddPicture(marcaDagua, false, true, 0, 0, oDoc.Application.CentimetersToPoints((float)13), oDoc.Application.CentimetersToPoints((float)13), myRange);
            myShape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapTight;
            myShape.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            myShape.Left = 130;
            myShape.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            myShape.Top = 180;
            myShape.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapBehind;
            myShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward);

            #endregion

            #endregion

            //oDoc.PrintOut();
            //SaveFileDialog saveFileDialog = new SaveFileDialog();
            //saveFileDialog.Filter = "Documento do Word|*.docx";
            //saveFileDialog.FileName = "Documento_Preenchido.docx";

            //if (saveFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    string savePath = saveFileDialog.FileName;
            //    object saveAsPath = savePath;
            //    object saveFormat = WdSaveFormat.wdFormatDocumentDefault;
            //    wordDocument.SaveAs2(ref saveAsPath, ref saveFormat);

            //    MessageBox.Show("Documento salvo com sucesso!");

            //    FinalizarDocumento();
            //    InitializeDocument();
            //}
        }
    }
}
