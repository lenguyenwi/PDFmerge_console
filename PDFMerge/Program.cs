using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.collection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFMerge
{
    class Program
    {       
        static void Main(string[] args)
        {
            string[] fileNames;
            string outFile;
            if (args.Count() > 0)
            {//wenn gar keine Argumente, beenden
                if (args[0].Trim().ToLower() == "-merge")
                {
                    List<string> fileNamesList = new List<string>();
                    for (int i = 1; i <= args.Count() - 3; i++)
                    {
                        fileNamesList.Add(args[i]);
                    }
                    if (fileNamesList.Count() >= 1)
                    {//nur wenn mindestens ein Pfad angegeben wurde
                        bool alleDateiPFadeExistieren = true;
                        foreach (string tmp_pfad in fileNamesList)
                        {
                            if (!File.Exists(tmp_pfad))
                            {
                                alleDateiPFadeExistieren = false;
                                break;
                            }
                        }

                        if (alleDateiPFadeExistieren)
                        {
                            fileNames = fileNamesList.ToArray();

                            outFile = args[args.Count() - 1];
                            CombineMultiplePDFs(fileNames, outFile);
                        }
                    }                    
                }
                else if (args[0].Trim().ToLower() == "-mergeall")
                {
                    List<string> fileNameAllsList = new List<string>();
                    DirectoryInfo d = new DirectoryInfo(args[1]);//Assuming Test is your Folder
                    FileInfo[] Files = d.GetFiles("*.pdf").OrderBy(p => p.CreationTime).ToArray(); ; //Getting Text files
                    string str = "";
                    foreach (FileInfo file in Files)
                    {
                        str = args[1] + "\\" + file.Name.ToString();
                        fileNameAllsList.Add(str);
                    }
                    string[] fileNamesAll = fileNameAllsList.ToArray();
                    string outFileAll = args[3];
                    if (fileNameAllsList.Count > 0)
                    {
                        CombineMultiplePDFs(fileNamesAll, outFileAll);
                    }
                }
                else if (args[0].Trim().ToLower() == "-portfolio")
                {
                    List<string> fileNamesList = new List<string>();
                    for (int i = 1; i <= args.Count() - 3; i++)
                    {
                        fileNamesList.Add(args[i]);
                    }
                    if (fileNamesList.Count() >= 1)
                    {//nur wenn mindestens ein Pfad angegeben wurde
                        bool alleDateiPFadeExistieren = true;
                        foreach (string tmp_pfad in fileNamesList)
                        {

                            if (!File.Exists(tmp_pfad))
                            {
                                alleDateiPFadeExistieren = false;
                                break;
                            }
                        }

                        if (alleDateiPFadeExistieren)
                        {
                            fileNames = fileNamesList.ToArray();

                            outFile = args[args.Count() - 1];
                            FolderWriter fw = new FolderWriter();
                            fw.Write(new FileStream(outFile, FileMode.Create), fileNames);
                        }
                    }
                }
            }
        }

        public static void CombineMultiplePDFs(string[] fileNames, string outFile)
        {
            // step 1: creation of a document-object
            Document document = new Document();

            // step 2: we create a writer that listens to the document
            PdfCopy writer = new PdfCopy(document, new FileStream(outFile, FileMode.Create));
            if (writer == null)
            {
                return;
            }
            // step 3: we open the document
            document.Open();
            foreach (string fileName in fileNames)
            {
                // we create a reader for a certain document
                PdfReader reader = new PdfReader(fileName);
                reader.ConsolidateNamedDestinations();

                // step 4: we add content
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    PdfImportedPage page = writer.GetImportedPage(reader, i);
                    writer.AddPage(page);
                }

                PRAcroForm form = reader.AcroForm;
                if (form != null)
                {
                    //writer.CopyAcroForm(reader);
                    writer.CopyDocumentFields(reader);
                }
                reader.Close();
            }
            // step 5: we close the document and writer
            writer.Close();
            document.Close();
        }
    }

    public class FolderWriter
    {
        private readonly string[] keys = new[] {"Type","File"};

        public void Write(Stream stream, string[] pfade)
        {
            using (Document document = new Document())
            {
                PdfWriter writer = PdfWriter.GetInstance(document, stream);

                document.Open();
                document.Add(new Paragraph(" "));

                PdfIndirectReference parentFolderObjectReference = writer.PdfIndirectReference;

                PdfCollection collection = new PdfCollection(PdfCollection.DETAILS);
                PdfCollectionSchema schema = CollectionSchema();
                collection.Schema = schema;
                collection.Sort = new PdfCollectionSort(keys);
                collection.Put(new PdfName("Vorlagen BA"), parentFolderObjectReference);
                writer.Collection = collection;

                PdfFileSpecification fs;
                PdfCollectionItem item;

                int nummer = 1;

                foreach (string pfad in pfade)
                {
                    String Filename = Path.GetFileName(pfad);
                    fs = PdfFileSpecification.FileEmbedded(writer, pfad,string.Format("{0} - {1}", nummer, Filename), null);
                    item = new PdfCollectionItem(schema);
                    item.AddItem("Type", "pdf");
                    fs.AddCollectionItem(item);
                    fs.AddDescription(GetDescription(Filename), false);
                    writer.AddFileAttachment(fs);

                    nummer++;
                }

                document.Close();
            }
        }

        private static string GetDescription(string fileName)
        {
            return fileName.Replace(".pdf","").Replace("_"," ");
        }
        
        private static PdfCollectionSchema CollectionSchema()
        {
            PdfCollectionSchema schema = new PdfCollectionSchema();
            PdfCollectionField type = new PdfCollectionField("File type", PdfCollectionField.TEXT);
            type.Order = 0;
            schema.AddField("Type", type);
            PdfCollectionField filename = new PdfCollectionField("File", PdfCollectionField.FILENAME);
            filename.Order = 1;
            schema.AddField("File", filename);
            return schema;
        }
    }

}
