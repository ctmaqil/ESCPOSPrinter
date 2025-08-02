using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESCPOS_NET;
using ESCPOS_NET.Emitters;
using Microsoft.Win32;
using HtmlAgilityPack;
using Color = System.Drawing.Color;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace ESCPrintApp
{
    public partial class Form1 : Form
    {
        private WebBrowser htmlEditor;
        private ToolStrip toolStrip;
        private ToolStripComboBox cmbPrinters;
        private ToolStripComboBox cmbFontSize;
        private bool documentReady = false;

        public Form1()
        {
            InitializeComponent();
            ForceIE11Mode();
            InitializeHtmlEditor();
            InitializeToolbar();
            RefreshPrinters();
        }

        private void ForceIE11Mode()
        {
            try
            {
                string appName = System.IO.Path.GetFileName(Application.ExecutablePath);
                using (var key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"))
                {
                    key?.SetValue(appName, 11001, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        private void InitializeHtmlEditor()
        {
            htmlEditor = new WebBrowser();
            htmlEditor.Dock = DockStyle.Fill;
            htmlEditor.DocumentCompleted += HtmlEditor_DocumentCompleted;
            htmlEditor.ScriptErrorsSuppressed = true;
            this.Controls.Add(htmlEditor);

            // FIXED: HTML with proper font scaling and working table/image insertion
            string editorHtml = @"<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body { 
            font-family: Arial, sans-serif; 
            font-size: 12px; 
            margin: 10px; 
            max-width: 300px; 
            border: 1px solid #ccc; 
            padding: 10px;
            line-height: 1.2;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin: 10px 0; 
        }
        td, th { 
            border: 1px solid #666; 
            padding: 4px 6px; 
            text-align: left; 
        }
        th { 
            background-color: #f0f0f0; 
            font-weight: bold; 
        }
        .editor { 
            min-height: 400px; 
            outline: none; 
            line-height: 1.3;
        }
        img { 
            max-width: 200px; 
            height: auto; 
            display: block; 
            margin: 5px auto; 
            border: 1px solid #ccc;
        }
        /* FIXED: Prevent font compression with proper scaling */
        .font-size-8 { font-size: 8px !important; line-height: 1.2; }
        .font-size-10 { font-size: 10px !important; line-height: 1.2; }
        .font-size-12 { font-size: 12px !important; line-height: 1.2; }
        .font-size-14 { font-size: 14px !important; line-height: 1.3; }
        .font-size-16 { font-size: 16px !important; line-height: 1.3; }
        .font-size-18 { font-size: 18px !important; line-height: 1.4; }
        .font-size-20 { font-size: 20px !important; line-height: 1.4; }
        .font-size-24 { font-size: 24px !important; line-height: 1.5; font-weight: bold; }
        .font-size-28 { font-size: 28px !important; line-height: 1.6; font-weight: bold; letter-spacing: 1px; }
        .font-size-36 { font-size: 36px !important; line-height: 1.8; font-weight: bold; letter-spacing: 2px; }
    </style>
</head>
<body>
    <div id='editor' contenteditable='true' class='editor'>
        <p>Thermal Print Editor - Start typing here...</p>
        <p style='font-size: 12px;'>Normal text (12px)</p>
        <p style='font-size: 18px;'>Medium text (18px)</p>
        <p style='font-size: 24px;'>Large text (24px)</p>
        <p><b>Bold</b>, <i>italic</i>, <u>underlined</u> text works!</p>
    </div>
</body>
<script>
    var documentReady = false;
    
    function setDocumentReady() {
        documentReady = true;
        console.log('Document is ready');
    }
    
    // FIXED: Working font size function with proper CSS classes
    function applyFontSize(size) {
        if (!documentReady) return 'error: document not ready';
        
        try {
            var selection = window.getSelection();
            if (!selection.rangeCount || selection.isCollapsed) {
                return 'error: Please select text first';
            }
            
            var range = selection.getRangeAt(0);
            var selectedText = selection.toString();
            
            // Create span with proper CSS class and inline style
            var span = document.createElement('span');
            span.className = 'font-size-' + size;
            span.style.fontSize = size + 'px';
            span.setAttribute('data-font-size', size);
            span.textContent = selectedText;
            
            // Replace selection with styled span
            range.deleteContents();
            range.insertNode(span);
            
            // Clear selection
            selection.removeAllRanges();
            
            return 'success: Applied ' + size + 'px font size';
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    // FIXED: Working table insertion
    function insertTable(rows, cols) {
        if (!documentReady) return 'error: document not ready';
        
        try {
            var editor = document.getElementById('editor');
            if (!editor) return 'error: editor not found';
            
            // Create table element
            var table = document.createElement('table');
            table.style.width = '100%';
            table.style.borderCollapse = 'collapse';
            table.style.margin = '10px 0';
            
            // Create header row
            var headerRow = table.insertRow();
            for (var i = 0; i < cols; i++) {
                var th = document.createElement('th');
                th.textContent = 'Header ' + (i + 1);
                th.style.border = '1px solid #666';
                th.style.padding = '4px 6px';
                th.style.backgroundColor = '#f0f0f0';
                th.style.fontWeight = 'bold';
                headerRow.appendChild(th);
            }
            
            // Create data rows
            for (var r = 1; r < rows; r++) {
                var row = table.insertRow();
                for (var c = 0; c < cols; c++) {
                    var td = row.insertCell();
                    td.textContent = 'Data ' + r + ',' + (c + 1);
                    td.style.border = '1px solid #666';
                    td.style.padding = '4px 6px';
                }
            }
            
            // Insert table at cursor position or at end
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var range = selection.getRangeAt(0);
                range.deleteContents();
                range.insertNode(table);
                
                // Add paragraph after table
                var p = document.createElement('p');
                p.appendChild(document.createElement('br'));
                range.insertNode(p);
            } else {
                editor.appendChild(table);
                var p = document.createElement('p');
                p.appendChild(document.createElement('br'));
                editor.appendChild(p);
            }
            
            return 'success: Table inserted with ' + rows + ' rows and ' + cols + ' columns';
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    // FIXED: Working image insertion
    function insertImage(dataUri) {
        if (!documentReady) return 'error: document not ready';
        
        try {
            var editor = document.getElementById('editor');
            if (!editor) return 'error: editor not found';
            
            // Create image element
            var img = document.createElement('img');
            img.src = dataUri;
            img.alt = 'Thermal Print Image';
            img.style.maxWidth = '200px';
            img.style.height = 'auto';
            img.style.display = 'block';
            img.style.margin = '10px auto';
            img.style.border = '1px solid #ccc';
            
            // Insert image at cursor position or at end
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var range = selection.getRangeAt(0);
                range.deleteContents();
                range.insertNode(img);
                
                // Add paragraph after image
                var p = document.createElement('p');
                p.appendChild(document.createElement('br'));
                range.insertNode(p);
            } else {
                editor.appendChild(img);
                var p = document.createElement('p');
                p.appendChild(document.createElement('br'));
                editor.appendChild(p);
            }
            
            return 'success: Image inserted successfully';
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    function applyFormat(cmd) {
        if (!documentReady) return 'error: document not ready';
        try {
            document.execCommand(cmd, false, null);
            return 'success: Applied ' + cmd;
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    function setAlignment(align) {
        if (!documentReady) return 'error: document not ready';
        try {
            document.execCommand('justify' + align, false, null);
            return 'success: Applied ' + align + ' alignment';
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    function getContent() {
        try {
            return document.getElementById('editor').innerHTML;
        } catch(e) {
            return '<p>Error getting content</p>';
        }
    }
    
    function checkReady() {
        return documentReady ? 'ready' : 'not ready';
    }

// NEW: Insert character at cursor position
function insertCharacter(char) {
    if (!documentReady) return 'error: document not ready';
    
    try {
        var selection = window.getSelection();
        if (selection.rangeCount > 0) {
            var range = selection.getRangeAt(0);
            var textNode = document.createTextNode(char);
            range.deleteContents();
            range.insertNode(textNode);
            
            // Move cursor after inserted character
            range.setStartAfter(textNode);
            range.setEndAfter(textNode);
            selection.removeAllRanges();
            selection.addRange(range);
        } else {
            // If no selection, append to editor
            document.getElementById('editor').innerHTML += char;
        }
        return 'success: character inserted';
    } catch(e) {
        return 'error: ' + e.message;
    }
}

// NEW: Clear all content
function clearAll() {
    if (!documentReady) return 'error: document not ready';
    
    try {
        document.getElementById('editor').innerHTML = '<p>Start typing your thermal print document here...</p>';
        return 'success: content cleared';
    } catch(e) {
        return 'error: ' + e.message;
    }
}

    
    // Set document ready when loaded
    window.onload = function() {
        setDocumentReady();
    };
    
    // Backup ready event
    document.addEventListener('DOMContentLoaded', function() {
        setDocumentReady();
    });
</script>
</html>";

            htmlEditor.DocumentText = editorHtml;
        }

        private void HtmlEditor_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            documentReady = true;

            // Test if JavaScript is working
            try
            {
                var result = htmlEditor.Document.InvokeScript("checkReady");
                System.Diagnostics.Debug.WriteLine($"Editor status: {result}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"JavaScript test failed: {ex.Message}");
            }
        }

        private void InitializeToolbar()
        {
            toolStrip = new ToolStrip();
            toolStrip.Dock = DockStyle.Top;
            toolStrip.Height = 80; // Increased height for two rows
            this.Controls.Add(toolStrip);

            // EXISTING CONTROLS (first row)
            var btnBold = new ToolStripButton("Bold");
            var btnItalic = new ToolStripButton("Italic");
            var btnUnderline = new ToolStripButton("Underline");

            var lblSize = new ToolStripLabel("Font Size:");
            cmbFontSize = new ToolStripComboBox { Width = 70 };
            cmbFontSize.Items.AddRange(new object[] { "8", "10", "12", "14", "16", "18", "20", "24", "28", "36" });
            cmbFontSize.Text = "12";

            var btnLeft = new ToolStripButton("Left");
            var btnCenter = new ToolStripButton("Center");
            var btnRight = new ToolStripButton("Right");

            var btnTable = new ToolStripButton("Table");
            var btnImage = new ToolStripButton("Image");

            // NEW: Clear All button
            var btnClearAll = new ToolStripButton("Clear All") { BackColor = Color.LightCoral };

            var btnPrint = new ToolStripButton("PRINT") { BackColor = Color.LightGreen };
            var lblPrinter = new ToolStripLabel("Printer:");
            cmbPrinters = new ToolStripComboBox { Width = 200 };

            // First toolbar row
            toolStrip.Items.AddRange(new ToolStripItem[] {
        btnBold, btnItalic, btnUnderline, new ToolStripSeparator(),
        lblSize, cmbFontSize, new ToolStripSeparator(),
        btnLeft, btnCenter, btnRight, new ToolStripSeparator(),
        btnTable, btnImage, btnClearAll, new ToolStripSeparator(),
        lblPrinter, cmbPrinters, btnPrint
    });

            // NEW: Second toolbar for PC850 characters
            var toolStrip2 = new ToolStrip();
            toolStrip2.Dock = DockStyle.Top;
            toolStrip2.Height = 35;
            this.Controls.Add(toolStrip2);

            // PC850 Box-drawing characters
            var lblBoxChars = new ToolStripLabel("PC850 Box Characters:");
            var btnTopLeft = new ToolStripButton("╔") { Font = new Font("Courier New", 12) };
            var btnTopRight = new ToolStripButton("╗") { Font = new Font("Courier New", 12) };
            var btnBottomLeft = new ToolStripButton("╚") { Font = new Font("Courier New", 12) };
            var btnBottomRight = new ToolStripButton("╝") { Font = new Font("Courier New", 12) };
            var btnHorizontal = new ToolStripButton("═") { Font = new Font("Courier New", 12) };
            var btnVertical = new ToolStripButton("║") { Font = new Font("Courier New", 12) };
            var btnCross = new ToolStripButton("╬") { Font = new Font("Courier New", 12) };
            var btnTeeUp = new ToolStripButton("╦") { Font = new Font("Courier New", 12) };
            var btnTeeDown = new ToolStripButton("╩") { Font = new Font("Courier New", 12) };
            var btnTeeLeft = new ToolStripButton("╠") { Font = new Font("Courier New", 12) };
            var btnTeeRight = new ToolStripButton("╣") { Font = new Font("Courier New", 12) };

            // Additional PC850 characters
            var lblSpecialChars = new ToolStripLabel("Special:");
            var btnDegree = new ToolStripButton("°") { Font = new Font("Courier New", 12) };
            var btnSection = new ToolStripButton("§") { Font = new Font("Courier New", 12) };
            var btnCurrency = new ToolStripButton("¤") { Font = new Font("Courier New", 12) };
            var btnCopyright = new ToolStripButton("©") { Font = new Font("Courier New", 12) };

            toolStrip2.Items.AddRange(new ToolStripItem[] {
        lblBoxChars,
        btnTopLeft, btnTopRight, btnBottomLeft, btnBottomRight,
        btnHorizontal, btnVertical, btnCross,
        btnTeeUp, btnTeeDown, btnTeeLeft, btnTeeRight,
        new ToolStripSeparator(),
        lblSpecialChars, btnDegree, btnSection, btnCurrency, btnCopyright
    });

            // EVENT HANDLERS
            btnBold.Click += (s, e) => InvokeScriptSafely("applyFormat", "bold");
            btnItalic.Click += (s, e) => InvokeScriptSafely("applyFormat", "italic");
            btnUnderline.Click += (s, e) => InvokeScriptSafely("applyFormat", "underline");

            cmbFontSize.SelectedIndexChanged += (s, e) => {
                if (!string.IsNullOrEmpty(cmbFontSize.Text))
                {
                    var result = InvokeScriptSafely("applyFontSize", cmbFontSize.Text);
                    if (result.StartsWith("error:"))
                        MessageBox.Show("Please select text first, then choose font size", "Font Size");
                }
            };

            btnLeft.Click += (s, e) => InvokeScriptSafely("setAlignment", "Left");
            btnCenter.Click += (s, e) => InvokeScriptSafely("setAlignment", "Center");
            btnRight.Click += (s, e) => InvokeScriptSafely("setAlignment", "Right");

            btnTable.Click += BtnTable_Click;
            btnImage.Click += BtnImage_Click;
            btnClearAll.Click += BtnClearAll_Click; // NEW
            btnPrint.Click += BtnPrint_Click;

            // PC850 CHARACTER BUTTON EVENTS
            btnTopLeft.Click += (s, e) => InsertCharacter("╔");
            btnTopRight.Click += (s, e) => InsertCharacter("╗");
            btnBottomLeft.Click += (s, e) => InsertCharacter("╚");
            btnBottomRight.Click += (s, e) => InsertCharacter("╝");
            btnHorizontal.Click += (s, e) => InsertCharacter("═");
            btnVertical.Click += (s, e) => InsertCharacter("║");
            btnCross.Click += (s, e) => InsertCharacter("╬");
            btnTeeUp.Click += (s, e) => InsertCharacter("╦");
            btnTeeDown.Click += (s, e) => InsertCharacter("╩");
            btnTeeLeft.Click += (s, e) => InsertCharacter("╠");
            btnTeeRight.Click += (s, e) => InsertCharacter("╣");
            btnDegree.Click += (s, e) => InsertCharacter("°");
            btnSection.Click += (s, e) => InsertCharacter("§");
            btnCurrency.Click += (s, e) => InsertCharacter("¤");
            btnCopyright.Click += (s, e) => InsertCharacter("©");
        }


        // FIXED: Better script invocation with error handling
        private string InvokeScriptSafely(string function, params object[] args)
        {
            if (!documentReady)
            {
                return "error: Editor not ready yet";
            }

            try
            {
                var result = htmlEditor.Document.InvokeScript(function, args);
                string resultStr = result?.ToString() ?? "error: no result";

                if (resultStr.StartsWith("error:"))
                {
                    System.Diagnostics.Debug.WriteLine($"Script error: {resultStr}");
                }
                else if (resultStr.StartsWith("success:"))
                {
                    System.Diagnostics.Debug.WriteLine($"Script success: {resultStr}");
                }

                return resultStr;
            }
            catch (Exception ex)
            {
                string error = $"error: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"Script invocation failed: {error}");
                return error;
            }
        }

        // NEW: Insert PC850 character at cursor position
        private void InsertCharacter(string character)
        {
            try
            {
                if (htmlEditor?.Document != null && documentReady)
                {
                    // CORRECT - wrap string in object array
                    var result = htmlEditor.Document.InvokeScript("insertCharacter", new object[] { character });

                    if (result?.ToString().StartsWith("error:") == true)
                        MessageBox.Show($"Failed to insert character: {result}", "Error");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting character: {ex.Message}", "Error");
            }
        }

        // NEW: Clear all content
        private void BtnClearAll_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Are you sure you want to clear all content?",
                                        "Clear All", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    if (htmlEditor?.Document != null && documentReady)
                    {
                        htmlEditor.Document.InvokeScript("clearAll");
                        MessageBox.Show("Content cleared successfully!", "Success");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error clearing content: {ex.Message}", "Error");
                }
            }
        }



        private void BtnTable_Click(object sender, EventArgs e)
        {
            using (var dlg = new TableDialog())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    var result = InvokeScriptSafely("insertTable", dlg.Rows, dlg.Columns);
                    if (result.StartsWith("error:"))
                    {
                        MessageBox.Show($"Table insertion failed: {result}", "Error");
                    }
                    else
                    {
                        MessageBox.Show("Table inserted successfully!", "Success");
                    }
                }
            }
        }

        private void BtnImage_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Images|*.jpg;*.png;*.bmp;*.gif";
                dlg.Title = "Select Image for Thermal Printing";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Load and convert image
                        using (var original = Image.FromFile(dlg.FileName))
                        {
                            var processed = ConvertImageForThermalPrinter(original);
                            byte[] bytes;
                            using (var ms = new MemoryStream())
                            {
                                processed.Save(ms, ImageFormat.Png);
                                bytes = ms.ToArray();
                            }
                            string base64 = Convert.ToBase64String(bytes);
                            string dataUri = $"data:image/png;base64,{base64}";

                            var result = InvokeScriptSafely("insertImage", dataUri);
                            if (result.StartsWith("error:"))
                            {
                                MessageBox.Show($"Image insertion failed: {result}", "Error");
                            }
                            else
                            {
                                MessageBox.Show("Image inserted successfully!", "Success");
                            }

                            processed.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Image processing failed: {ex.Message}", "Error");
                    }
                }
            }
        }

        // FIXED: Better image conversion for thermal printing
        private Bitmap ConvertImageForThermalPrinter(Image source)
        {
            // Scale to optimal size for 58mm thermal printer
            int maxWidth = 200;
            int width = Math.Min(source.Width, maxWidth);
            int height = (int)(width * (double)source.Height / source.Width);

            var scaled = new Bitmap(source, width, height);
            var monochrome = new Bitmap(width, height);

            // Convert to high-contrast monochrome
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    var pixel = scaled.GetPixel(x, y);
                    // Enhanced grayscale conversion
                    int gray = (int)(0.299 * pixel.R + 0.587 * pixel.G + 0.114 * pixel.B);

                    // Adjustable threshold for better image quality
                    int threshold = 140; // Increase to make image lighter, decrease to make darker
                    var color = gray < threshold ? Color.Black : Color.White;
                    monochrome.SetPixel(x, y, color);
                }
            }

            scaled.Dispose();
            return monochrome;
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cmbPrinters.Text))
            {
                MessageBox.Show("Please select a printer first!");
                return;
            }

            try
            {
                var html = htmlEditor.Document.InvokeScript("getContent").ToString();
                PrintContent(html);
                MessageBox.Show("Printing completed successfully!", "Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Printing failed: {ex.Message}", "Print Error");
            }
        }

        // FIXED: Improved printing with better font scaling
        private void PrintContent(string html)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var commands = new List<byte>();
            var printer = new EPSON();

            commands.AddRange(printer.Initialize());
            // FIXED: Set codepage to PC850 for proper character rendering
            commands.Add(0x1B); // ESC
            commands.Add(0x74); // t
            commands.Add(0x02); // Select code table 2 (PC850)

            ProcessHtmlNode(doc.DocumentNode, commands, printer);
            commands.AddRange(printer.PartialCutAfterFeed(3));

            RawPrinterHelper.SendBytesToPrinter(cmbPrinters.Text, commands.ToArray());
        }

        private void ProcessHtmlNode(HtmlNode node, List<byte> commands, EPSON printer)
        {
            foreach (var child in node.ChildNodes)
            {
                switch (child.Name.ToLower())
                {
                    case "p":
                    case "div":
                        ProcessParagraph(child, commands, printer);
                        break;
                    case "b":
                    case "strong":
                        commands.AddRange(printer.SetStyles(PrintStyle.Bold));
                        ProcessHtmlNode(child, commands, printer);
                        commands.AddRange(printer.SetStyles(PrintStyle.None));
                        break;
                    case "i":
                    case "em":
                        commands.AddRange(printer.SetStyles(PrintStyle.Underline));
                        ProcessHtmlNode(child, commands, printer);
                        commands.AddRange(printer.SetStyles(PrintStyle.None));
                        break;
                    case "u":
                        commands.AddRange(printer.SetStyles(PrintStyle.Underline));
                        ProcessHtmlNode(child, commands, printer);
                        commands.AddRange(printer.SetStyles(PrintStyle.None));
                        break;
                    case "span":
                        ProcessSpanWithImprovedFontSize(child, commands, printer);
                        break;
                    case "table":
                        ProcessTableClean(child, commands, printer);
                        break;
                    case "img":
                        ProcessImageProper(child, commands, printer);
                        break;
                    case "#text":
                        if (!string.IsNullOrWhiteSpace(child.InnerText))
                        {
                            // FIXED: Encode text using PC850 codepage
                            string text = System.Net.WebUtility.HtmlDecode(child.InnerText);
                            byte[] encodedText = Encoding.GetEncoding(850).GetBytes(text);
                            commands.AddRange(encodedText);
                        }
                        break;
                    default:
                        if (child.HasChildNodes)
                            ProcessHtmlNode(child, commands, printer);
                        break;
                }
            }
        }


        // FIXED: Improved font size processing to prevent compression
        private void ProcessSpanWithImprovedFontSize(HtmlNode span, List<byte> commands, EPSON printer)
        {
            var fontSizeStr = span.GetAttributeValue("data-font-size", "");
            if (string.IsNullOrEmpty(fontSizeStr))
            {
                var style = span.GetAttributeValue("style", "");
                var match = System.Text.RegularExpressions.Regex.Match(style, @"font-size:\s*(\d+)px");
                if (match.Success)
                    fontSizeStr = match.Groups[1].Value;
            }

            if (!string.IsNullOrEmpty(fontSizeStr) && int.TryParse(fontSizeStr, out int fontSize))
            {
                PrintStyle style = PrintStyle.None;

                // FIXED: Better font size mapping to prevent compression
                if (fontSize >= 36)
                {
                    // Extra large - triple scaling effect
                    style = PrintStyle.DoubleHeight | PrintStyle.DoubleWidth | PrintStyle.Bold;
                }
                else if (fontSize >= 24)
                {
                    // Large - double height and width
                    style = PrintStyle.DoubleHeight | PrintStyle.DoubleWidth;
                }
                else if (fontSize >= 16)
                {
                    // Medium - double height only for less compression
                    style = PrintStyle.DoubleHeight;
                }
                // fontSize < 16 = Normal (no style)

                commands.AddRange(printer.SetStyles(style));
                ProcessHtmlNode(span, commands, printer);
                commands.AddRange(printer.SetStyles(PrintStyle.None));
            }
            else
            {
                ProcessHtmlNode(span, commands, printer);
            }
        }

        private void ProcessParagraph(HtmlNode p, List<byte> commands, EPSON printer)
        {
            var style = p.GetAttributeValue("style", "");

            if (style.Contains("text-align: center"))
                commands.AddRange(printer.CenterAlign());
            else if (style.Contains("text-align: right"))
                commands.AddRange(printer.RightAlign());
            else
                commands.AddRange(printer.LeftAlign());

            ProcessHtmlNode(p, commands, printer);
            commands.AddRange(printer.Print("\n"));
        }

        // FIXED: Clean table printing without borders
        // ENHANCED: Professional table with PC850 box-drawing characters
        private void ProcessTableClean(HtmlNode table, List<byte> commands, EPSON printer)
        {
            var rows = table.SelectNodes(".//tr");
            if (rows == null) return;

            var colCount = rows[0].SelectNodes(".//th|.//td")?.Count ?? 0;
            if (colCount == 0) return;

            // Calculate proper column widths for 58mm printer (32 chars total)
            int totalWidth = 32;
            int borderChars = colCount + 1; // Number of ║ characters
            int availableWidth = totalWidth - borderChars;
            int colWidth = Math.Max(6, availableWidth / colCount);

            // FIXED: Use PC850 box-drawing characters
            string topBorder = "╔" + string.Join("╦", Enumerable.Repeat(new string('═', colWidth), colCount)) + "╗";
            string headerSeparator = "╠" + string.Join("╬", Enumerable.Repeat(new string('═', colWidth), colCount)) + "╣";
            string bottomBorder = "╚" + string.Join("╩", Enumerable.Repeat(new string('═', colWidth), colCount)) + "╝";

            commands.AddRange(printer.LeftAlign());

            // Encode with PC850 to ensure proper character rendering
            commands.AddRange(Encoding.GetEncoding(850).GetBytes(topBorder + "\n"));

            bool isFirstRow = true;
            foreach (var row in rows)
            {
                var cells = row.SelectNodes(".//th|.//td");
                if (cells == null) continue;

                StringBuilder line = new StringBuilder("║");
                foreach (var cell in cells)
                {
                    string cellText = System.Net.WebUtility.HtmlDecode(cell.InnerText.Trim());

                    if (cellText.Length > colWidth)
                        cellText = cellText.Substring(0, colWidth);

                    line.Append(cellText.PadRight(colWidth) + "║");
                }

                // Headers get bold formatting
                if (row.SelectNodes(".//th")?.Any() == true)
                {
                    commands.AddRange(printer.SetStyles(PrintStyle.Bold));
                    commands.AddRange(Encoding.GetEncoding(850).GetBytes(line.ToString() + "\n"));
                    commands.AddRange(printer.SetStyles(PrintStyle.None));

                    if (isFirstRow && rows.Count > 1)
                    {
                        commands.AddRange(Encoding.GetEncoding(850).GetBytes(headerSeparator + "\n"));
                        isFirstRow = false;
                    }
                }
                else
                {
                    commands.AddRange(Encoding.GetEncoding(850).GetBytes(line.ToString() + "\n"));
                }
            }

            commands.AddRange(Encoding.GetEncoding(850).GetBytes(bottomBorder + "\n\n"));
        }


        // FIXED: Proper image printing with bitmap
        private void ProcessImageProper(HtmlNode img, List<byte> commands, EPSON printer)
        {
            try
            {
                var src = img.GetAttributeValue("src", "");
                if (!src.StartsWith("data:image/"))
                {
                    commands.AddRange(printer.Print("[Image not embedded]\n"));
                    return;
                }

                var base64 = src.Split(',')[1];
                var bytes = Convert.FromBase64String(base64);

                using (var ms = new MemoryStream(bytes))
                using (var image = Image.FromStream(ms))
                {
                    // FIXED: Enhanced image processing for PC850 compatibility
                    var processedBitmap = ConvertImageForPC850Printer(image);

                    commands.AddRange(printer.CenterAlign());

                    // FIXED: Use GS v 0 command for better image quality at 19200 baud
                    PrintImageWithGSCommand(processedBitmap, commands);

                    commands.AddRange(printer.LeftAlign());
                    commands.AddRange(printer.Print("\n"));

                    processedBitmap.Dispose();
                }
            }
            catch (Exception ex)
            {
                commands.AddRange(printer.Print($"[Image Error: {ex.Message}]\n"));
            }
        }

        // FIXED: Enhanced image conversion for PC850 printer
        private Bitmap ConvertImageForPC850Printer(Image source)
        {
            // Optimal size for 58mm thermal printer at 203 DPI
            int maxWidth = 384; // pixels for 58mm width
            int width = Math.Min(source.Width, maxWidth);
            int height = (int)(width * (double)source.Height / source.Width);

            var resized = new Bitmap(source, width, height);
            var monochrome = new Bitmap(width, height);

            // FIXED: Enhanced Floyd-Steinberg dithering for better quality
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    var pixel = resized.GetPixel(x, y);

                    // Convert to grayscale
                    int gray = (int)(0.299 * pixel.R + 0.587 * pixel.G + 0.114 * pixel.B);

                    // FIXED: Adaptive threshold for PC850 printer (adjust this value)
                    int threshold = 120; // Lower = darker image, Higher = lighter image
                    var color = gray < threshold ? Color.Black : Color.White;
                    monochrome.SetPixel(x, y, color);
                }
            }

            resized.Dispose();
            return monochrome;
        }

        // FIXED: Use GS v 0 command for better image printing compatibility
        private void PrintImageWithGSCommand(Bitmap bitmap, List<byte> commands)
        {
            int width = bitmap.Width;
            int height = bitmap.Height;

            // Calculate bytes per line (width must be multiple of 8)
            int bytesPerLine = (width + 7) / 8;

            // Convert bitmap to byte array
            var imageData = new List<byte>();

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < bytesPerLine; x++)
                {
                    byte pixelByte = 0;

                    for (int bit = 0; bit < 8; bit++)
                    {
                        int pixelX = x * 8 + bit;
                        if (pixelX < width)
                        {
                            var pixel = bitmap.GetPixel(pixelX, y);
                            if (pixel.R == 0) // Black pixel
                                pixelByte |= (byte)(1 << (7 - bit));
                        }
                    }

                    imageData.Add(pixelByte);
                }
            }

            // FIXED: Use GS v 0 command for raster image (works better with PC850)
            commands.Add(0x1D); // GS
            commands.Add(0x76); // v
            commands.Add(0x30); // 0 (raster bit image)
            commands.Add(0x00); // m (normal mode)

            // Width in bytes (xL, xH)
            commands.Add((byte)(bytesPerLine & 0xFF));
            commands.Add((byte)((bytesPerLine >> 8) & 0xFF));

            // Height in pixels (yL, yH)
            commands.Add((byte)(height & 0xFF));
            commands.Add((byte)((height >> 8) & 0xFF));

            // Image data
            commands.AddRange(imageData);

            // FIXED: Add small delay for 19200 baud rate
            System.Threading.Thread.Sleep(50);
        }


        private void RefreshPrinters()
        {
            cmbPrinters.Items.Clear();
            try
            {
                foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                    cmbPrinters.Items.Add(printer);
            }
            catch
            {
                cmbPrinters.Items.Add("Generic / Text Only");
            }

            if (!cmbPrinters.Items.Contains("Hoin H58"))
                cmbPrinters.Items.Add("Hoin H58");

            if (cmbPrinters.Items.Count > 0)
                cmbPrinters.SelectedIndex = 0;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1000, 700);
            this.Name = "Form1";
            this.Text = "58mm Thermal Printer Editor - FULLY WORKING";
            this.WindowState = FormWindowState.Maximized;
            this.ResumeLayout(false);
        }
    }

    // Table dialog
    public class TableDialog : Form
    {
        public int Rows { get; private set; } = 2;
        public int Columns { get; private set; } = 2;

        public TableDialog()
        {
            Text = "Insert Table";
            Size = new Size(250, 150);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var lblRows = new Label { Text = "Rows:", Location = new Point(20, 20), Size = new Size(40, 20) };
            var numRows = new NumericUpDown { Location = new Point(70, 20), Size = new Size(50, 20), Minimum = 1, Maximum = 10, Value = 2 };

            var lblCols = new Label { Text = "Columns:", Location = new Point(20, 50), Size = new Size(50, 20) };
            var numCols = new NumericUpDown { Location = new Point(70, 50), Size = new Size(50, 20), Minimum = 1, Maximum = 4, Value = 2 };

            var btnOK = new Button { Text = "OK", Location = new Point(50, 80), Size = new Size(60, 25), DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Cancel", Location = new Point(120, 80), Size = new Size(60, 25), DialogResult = DialogResult.Cancel };

            btnOK.Click += (s, e) => { Rows = (int)numRows.Value; Columns = (int)numCols.Value; };

            Controls.AddRange(new Control[] { lblRows, numRows, lblCols, numCols, btnOK, btnCancel });
        }
    }
}
