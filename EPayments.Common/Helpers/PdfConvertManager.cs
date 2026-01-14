using iTextSharp.text.pdf;
using SautinSoft;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WkHtmlToPdfDotNet;
using static System.Resources.ResXFileRef;

namespace EPayments.Common.Helpers
{
    public class PdfConvertManager
    {
        static readonly string PDF_METAMORPHOSIS_SERIAL = "10024747414";
        static readonly string PDF_SECURITY_PASSWORD = "epayments@abbaty@encryption";
        private static readonly SynchronizedConverter converter = new SynchronizedConverter(new PdfTools());

        static readonly object rtfLocker = new object();
        static readonly object txtLocker = new object();
        static readonly object htmlLocker = new object();
        static readonly object docxLocker = new object();

        public static byte[] Convert(string input, ref string mimeType)
        {
            byte[] pdf;
            switch (mimeType)
            {
                case MimeTypeFileExtension.MIME_TEXT_RTF:
                    {
                        pdf = ConvertRtfToPdf(input);
                        if (pdf != null)
                            mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                case MimeTypeFileExtension.MIME_TEXT_PLAIN:
                    {
                        pdf = ConvertTxtToPdf(input);
                        if (pdf != null)
                            mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                case MimeTypeFileExtension.MIME_TEXT_HTML:
                    {
                        pdf = ConvertHtmlToPdf(input);
                        if (pdf != null)
                            mimeType = MimeTypeFileExtension.MIME_APPLICATION_PDF;
                        break;
                    }
                default:
                    {
                        return UTF8Encoding.UTF8.GetBytes(input);
                    }
            }

            return EncryptPdf(pdf);
        }

        private static byte[] ConvertHtmlToPdf(string input)
        {
            var pdf = new HtmlToPdfDocument
            {
                GlobalSettings = {
                    ColorMode = ColorMode.Grayscale,
                    Orientation = WkHtmlToPdfDotNet.Orientation.Portrait,
                    PaperSize = PaperKind.A4,
                },
                Objects = {
                    new ObjectSettings {
                        PagesCount = true,
                        HtmlContent = input,
                        WebSettings = {
                            DefaultEncoding = "utf-8",
                            LoadImages = false,
                            EnableJavascript = false
                        },
                    }
                }
            };

            return converter.Convert(pdf);
        }

        private static byte[] ConvertTxtToPdf(string input)
        {
            try
            {
                lock (txtLocker)
                {
                    PdfMetamorphosis pdfConverter = new PdfMetamorphosis();
                    pdfConverter.SetSerial(PDF_METAMORPHOSIS_SERIAL);
                    byte[] pdf = pdfConverter.RtfToPdfConvertByte(input);

                    return pdf;
                }
            }
            catch
            {
                return null;
            }
        }

        private static byte[] ConvertRtfToPdf(string inputString)
        {
            try
            {
                byte[] input = UTF8Encoding.UTF8.GetBytes(inputString);
                lock (rtfLocker)
                {
                    MemoryStream rtfMemoryStream = new MemoryStream(input);
                    MemoryStream pdfMemoryStream = new MemoryStream();

                    PdfMetamorphosis pdfConverter = new PdfMetamorphosis();
                    pdfConverter.SetSerial(PDF_METAMORPHOSIS_SERIAL);
                    RichTextBox rtfParser = new RichTextBox();

                    rtfParser.LoadFile(rtfMemoryStream, RichTextBoxStreamType.RichText);
                    rtfParser.SaveFile(pdfMemoryStream, RichTextBoxStreamType.RichText);

                    byte[] rtf = pdfMemoryStream.ToArray();
                    byte[] pdf = pdfConverter.RtfToPdfConvertByte(rtf);

                    return pdf;
                }
            }
            catch
            {
                return null;
            }
        }

        private static byte[] EncryptPdf(byte[] input)
        {
            using (MemoryStream inputStream = new MemoryStream(input))
            {
                using (MemoryStream outputStream = new MemoryStream())
                {
                    try
                    {
                        PdfReader reader = new PdfReader(inputStream);
                        PdfEncryptor.Encrypt(reader, outputStream, true, null, PDF_SECURITY_PASSWORD,
                            PdfWriter.ALLOW_PRINTING &
                            ~PdfWriter.ALLOW_COPY &
                            ~PdfWriter.ALLOW_ASSEMBLY &
                            PdfWriter.ALLOW_DEGRADED_PRINTING &
                            ~PdfWriter.ALLOW_FILL_IN &
                            ~PdfWriter.ALLOW_MODIFY_ANNOTATIONS &
                            ~PdfWriter.ALLOW_MODIFY_CONTENTS &
                            ~PdfWriter.ALLOW_SCREENREADERS);

                        return outputStream.ToArray();
                    }
                    catch
                    {
                        return input;
                    }
                }
            }
        }
    }
}
