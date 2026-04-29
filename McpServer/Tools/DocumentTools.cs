using System.ComponentModel;
using System.Text;
using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using KAPITypes;
using ModelContextProtocol.Server;

namespace McpKompas.Tools;

[McpServerToolType]
public sealed class DocumentTools
{
    // ---------------------------------------------------------------
    //  Create documents (via API7 for proper visibility)
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_create_drawing")]
    [Description("Create a new 2D drawing document (.cdw) in Kompas 3D and optionally save it to disk.")]
    public static string CreateDrawing(
        [Description("Full path where the drawing will be saved, e.g. C:\\Work\\detail.cdw. Leave empty to create an unsaved document.")]
        string file_path = "",
        [Description("Sheet format: 0=A0, 1=A1, 2=A2, 3=A3, 4=A4.")]
        int sheet_format = 3,
        [Description("Landscape orientation (true) or portrait (false).")]
        bool landscape = true,
        [Description("Document author name.")]
        string author = "",
        [Description("Document comment/description.")]
        string comment = "")
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";

        // Create document via API7 so it is visible in the Kompas window
        var app7 = KompasConnector.GetApplication7();
        if (app7 != null)
        {
            var doc7 = app7.Documents.Add(DocumentTypeEnum.ksDocumentDrawing, true);
            if (doc7 != null)
            {
                doc7.Active = true;
                KompasConnector.BringToFront();

                // Now configure sheet format via API5 on the active document
                var doc = (ksDocument2D?)k.ActiveDocument2D();
                if (doc != null)
                {
                    ksDocumentParam? docPar = (ksDocumentParam?)k.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
                    if (docPar != null)
                    {
                        doc.ksGetObjParam(doc.reference, docPar, ldefin2d.ALLPARAM);
                        docPar.comment = comment;
                        docPar.author = author;

                        ksSheetPar? shPar = (ksSheetPar?)docPar.GetLayoutParam();
                        if (shPar != null)
                        {
                            shPar.shtType = (short)1;
                            ksStandartSheet? stPar = (ksStandartSheet?)shPar.GetSheetParam();
                            if (stPar != null)
                            {
                                stPar.format = (short)sheet_format;
                                stPar.multiply = (short)1;
                                stPar.direct = landscape;
                            }
                        }
                        doc.ksSetObjParam(doc.reference, docPar, ldefin2d.ALLPARAM);
                    }

                    if (!string.IsNullOrEmpty(file_path) && !doc.ksSaveDocument(file_path))
                        return $"ERROR: Drawing created in memory but Kompas could not save it to: {file_path}";
                }

                UndoManager.Clear();
                string saved = string.IsNullOrEmpty(file_path) ? "unsaved" : file_path;
                return $"Drawing created. Format: A{sheet_format}, {(landscape ? "landscape" : "portrait")}. File: {saved}";
            }
        }

        // Fallback to API5 if API7 fails
        ksDocument2D? doc5 = (ksDocument2D?)k.Document2D();
        if (doc5 == null)
            return "ERROR: Could not create Document2D object.";

        ksDocumentParam? docPar5 = (ksDocumentParam?)k.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
        if (docPar5 == null)
            return "ERROR: Could not get DocumentParam struct.";

        docPar5.fileName = file_path;
        docPar5.comment = comment;
        docPar5.author = author;
        docPar5.regime = 0;
        docPar5.type = (short)DocType.lt_DocSheetStandart;

        ksSheetPar? shPar5 = (ksSheetPar?)docPar5.GetLayoutParam();
        if (shPar5 != null)
        {
            shPar5.shtType = (short)1;
            ksStandartSheet? stPar5 = (ksStandartSheet?)shPar5.GetSheetParam();
            if (stPar5 != null)
            {
                stPar5.format = (short)sheet_format;
                stPar5.multiply = (short)1;
                stPar5.direct = landscape;
            }
        }

        doc5.ksCreateDocument(docPar5);
        KompasConnector.BringToFront();
        UndoManager.Clear();
        string saved5 = string.IsNullOrEmpty(file_path) ? "unsaved" : file_path;
        return $"Drawing created (API5 fallback). Format: A{sheet_format}, {(landscape ? "landscape" : "portrait")}. File: {saved5}";
    }

    [McpServerTool(Name = "kompas_create_fragment")]
    [Description("Create a new 2D fragment document (.frw) in Kompas 3D.")]
    public static string CreateFragment(
        [Description("Full path where the fragment will be saved, e.g. C:\\Work\\frag.frw. Leave empty for unsaved.")]
        string file_path = "",
        [Description("Document author name.")]
        string author = "",
        [Description("Document comment/description.")]
        string comment = "")
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";

        var app7 = KompasConnector.GetApplication7();
        if (app7 != null)
        {
            var doc7 = app7.Documents.Add(DocumentTypeEnum.ksDocumentFragment, true);
            if (doc7 != null)
            {
                doc7.Active = true;
                KompasConnector.BringToFront();

                if (!string.IsNullOrEmpty(file_path))
                {
                    var doc = (ksDocument2D?)k.ActiveDocument2D();
                    if (doc != null && !doc.ksSaveDocument(file_path))
                        return $"ERROR: Fragment created in memory but Kompas could not save it to: {file_path}";
                }

                UndoManager.Clear();
                string saved = string.IsNullOrEmpty(file_path) ? "unsaved" : file_path;
                return $"Fragment created. File: {saved}";
            }
        }

        // Fallback
        ksDocument2D? doc5 = (ksDocument2D?)k.Document2D();
        if (doc5 == null)
            return "ERROR: Could not create Document2D object.";

        ksDocumentParam? docPar = (ksDocumentParam?)k.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
        if (docPar == null)
            return "ERROR: Could not get DocumentParam struct.";

        docPar.fileName = file_path;
        docPar.comment = comment;
        docPar.author = author;
        docPar.regime = 0;
        docPar.type = (short)DocType.lt_DocFragment;

        doc5.ksCreateDocument(docPar);
        KompasConnector.BringToFront();
        UndoManager.Clear();
        string saved5 = string.IsNullOrEmpty(file_path) ? "unsaved" : file_path;
        return $"Fragment created (API5 fallback). File: {saved5}";
    }

    [McpServerTool(Name = "kompas_create_3d_part")]
    [Description("Create a new 3D part document (.m3d) in Kompas 3D.")]
    public static string Create3dPart(
        [Description("Full path where the part will be saved, e.g. C:\\Work\\part.m3d. Leave empty for unsaved.")]
        string file_path = "",
        [Description("Document author name.")]
        string author = "",
        [Description("Document comment/description.")]
        string comment = "")
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";

        var app7 = KompasConnector.GetApplication7();
        if (app7 != null)
        {
            var doc7 = app7.Documents.Add(DocumentTypeEnum.ksDocumentPart, true);
            if (doc7 != null)
            {
                doc7.Active = true;
                KompasConnector.BringToFront();

                // Set author/comment via API5
                var doc = (ksDocument3D?)k.ActiveDocument3D();
                if (doc != null)
                {
                    if (!string.IsNullOrEmpty(author)) doc.author = author;
                    if (!string.IsNullOrEmpty(comment)) doc.comment = comment;
                    doc.UpdateDocumentParam();

                    if (!string.IsNullOrEmpty(file_path))
                    {
                        if (!doc.SaveAs(file_path))
                            return $"ERROR: 3D part created in memory but Kompas could not save it to: {file_path}";
                        UndoManager.Clear();
                        return $"3D part created and saved: {file_path}";
                    }
                }
                UndoManager.Clear();
                return "3D part created (unsaved).";
            }
        }

        // Fallback
        ksDocument3D? doc5 = (ksDocument3D?)k.Document3D();
        if (doc5 == null)
            return "ERROR: Could not create Document3D object.";

        doc5.Create(true, true);
        if (!string.IsNullOrEmpty(author)) doc5.author = author;
        if (!string.IsNullOrEmpty(comment)) doc5.comment = comment;
        doc5.UpdateDocumentParam();
        KompasConnector.BringToFront();

        if (!string.IsNullOrEmpty(file_path))
        {
            if (!doc5.SaveAs(file_path))
                return $"ERROR: 3D part created in memory but Kompas could not save it to: {file_path}";
            UndoManager.Clear();
            return $"3D part created and saved (API5 fallback): {file_path}";
        }
        UndoManager.Clear();
        return "3D part created (unsaved, API5 fallback).";
    }

    // ---------------------------------------------------------------
    //  Open / Save / Close
    // ---------------------------------------------------------------

    [McpServerTool(Name = "kompas_open_document")]
    [Description("Open an existing Kompas document from disk. Supported formats: .cdw, .frw (2D), .m3d, .a3d (3D).")]
    public static string OpenDocument(
        [Description("Full path to the file to open.")]
        string file_path,
        [Description("Open in read-only mode.")]
        bool read_only = false)
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";

        if (!File.Exists(file_path))
            return $"ERROR: File not found: {file_path}";

        // Use API7 for opening — ensures the document is visible
        var app7 = KompasConnector.GetApplication7();
        if (app7 != null)
        {
            var doc7 = app7.Documents.Open(file_path, true, read_only);
            if (doc7 != null)
            {
                doc7.Active = true;
                KompasConnector.BringToFront();
                // Refs from any previously tracked document are no longer valid in the new one
                UndoManager.Clear();
                return $"Opened document: {file_path}{(read_only ? " (read-only)" : "")}";
            }
        }

        // Fallback to API5
        string ext = Path.GetExtension(file_path).ToLowerInvariant();

        if (ext is ".m3d" or ".a3d")
        {
            ksDocument3D? doc3d = (ksDocument3D?)k.Document3D();
            if (doc3d == null)
                return "ERROR: Could not create Document3D object.";
            bool ok = doc3d.Open(file_path, read_only);
            if (ok)
            {
                KompasConnector.BringToFront();
                UndoManager.Clear();
            }
            return ok
                ? $"Opened 3D document: {file_path}{(read_only ? " (read-only)" : "")}"
                : $"ERROR: Kompas could not open: {file_path}";
        }
        else
        {
            ksDocument2D? doc2d = (ksDocument2D?)k.Document2D();
            if (doc2d == null)
                return "ERROR: Could not create Document2D object.";
            bool ok = doc2d.ksOpenDocument(file_path, read_only);
            if (ok)
            {
                KompasConnector.BringToFront();
                UndoManager.Clear();
            }
            return ok
                ? $"Opened 2D document: {file_path}{(read_only ? " (read-only)" : "")}"
                : $"ERROR: Kompas could not open: {file_path}";
        }
    }

    [McpServerTool(Name = "kompas_save_document")]
    [Description("Save the active Kompas document. Optionally provide a new path to save-as.")]
    public static string SaveDocument(
        [Description("New file path to save as. Leave empty to save in place.")]
        string file_path = "")
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";

        ksDocument2D? doc2d = (ksDocument2D?)k.ActiveDocument2D();
        if (doc2d != null)
        {
            // ksSaveDocument returns true on success, false on failure (bad path, no permission, etc.)
            if (!doc2d.ksSaveDocument(file_path))
                return string.IsNullOrEmpty(file_path)
                    ? "ERROR: Kompas could not save the 2D document."
                    : $"ERROR: Kompas could not save 2D document to: {file_path}";
            return string.IsNullOrEmpty(file_path) ? "2D document saved." : $"2D document saved as: {file_path}";
        }

        ksDocument3D? doc3d = (ksDocument3D?)k.ActiveDocument3D();
        if (doc3d != null)
        {
            if (string.IsNullOrEmpty(file_path))
            {
                if (!doc3d.Save())
                    return "ERROR: Kompas could not save the 3D document.";
                return "3D document saved.";
            }
            if (!doc3d.SaveAs(file_path))
                return $"ERROR: Kompas could not save 3D document to: {file_path}";
            return $"3D document saved as: {file_path}";
        }

        return "ERROR: No active document found.";
    }

    [McpServerTool(Name = "kompas_close_document")]
    [Description("Close the active Kompas document.")]
    public static string CloseDocument(
        [Description("Save changes before closing.")]
        bool save = true)
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";

        ksDocument2D? doc2d = (ksDocument2D?)k.ActiveDocument2D();
        if (doc2d != null)
        {
            if (save && !doc2d.ksSaveDocument(string.Empty))
                return "ERROR: Save before close failed; document not closed. " +
                       "Call kompas_save_document with an explicit path or pass save=false.";
            doc2d.ksCloseDocument();
            // Tracked refs belonged to the now-closed doc — drop them
            UndoManager.Clear();
            return "2D document closed.";
        }

        ksDocument3D? doc3d = (ksDocument3D?)k.ActiveDocument3D();
        if (doc3d != null)
        {
            if (save && !doc3d.Save())
                return "ERROR: Save before close failed; document not closed. " +
                       "Call kompas_save_document with an explicit path or pass save=false.";
            doc3d.close();
            UndoManager.Clear();
            return "3D document closed.";
        }

        return "ERROR: No active document to close.";
    }

    [McpServerTool(Name = "kompas_get_document_info")]
    [Description("Get information about the currently active Kompas document (type, file, author, comment).")]
    public static string GetDocumentInfo()
    {
        var k = KompasConnector.Instance;
        if (k == null) return "ERROR: Not connected to Kompas. Call kompas_connect first.";
        var sb = new StringBuilder();

        ksDocument2D? doc2d = (ksDocument2D?)k.ActiveDocument2D();
        if (doc2d != null)
        {
            ksDocumentParam? par = (ksDocumentParam?)k.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);
            if (par != null)
            {
                doc2d.ksGetObjParam(doc2d.reference, par, ldefin2d.ALLPARAM);
                sb.AppendLine("Type: 2D Document");
                sb.AppendLine($"File: {par.fileName}");
                sb.AppendLine($"Author: {par.author}");
                sb.AppendLine($"Comment: {par.comment}");
            }
            return sb.ToString().Trim();
        }

        ksDocument3D? doc3d = (ksDocument3D?)k.ActiveDocument3D();
        if (doc3d != null)
        {
            sb.AppendLine("Type: 3D Document");
            sb.AppendLine($"File: {doc3d.fileName}");
            sb.AppendLine($"Author: {doc3d.author}");
            sb.AppendLine($"Comment: {doc3d.comment}");
            return sb.ToString().Trim();
        }

        return "No active document.";
    }
}
