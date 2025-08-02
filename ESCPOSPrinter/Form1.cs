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
        private ToolStrip toolStrip2;
        private ToolStripComboBox cmbPrinters;
        private ToolStripComboBox cmbFontSize;
        private NumericUpDown nudCanvasWidth;
        private SplitContainer splitContainer;
        private bool documentReady = false;

        public Form1()
        {
            InitializeComponent();
            ForceIE11Mode();
            InitializeResizableEditor();
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

        private void InitializeResizableEditor()
        {
            splitContainer = new SplitContainer();
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Orientation = Orientation.Horizontal;
            splitContainer.SplitterDistance = 400;
            splitContainer.Panel1MinSize = 200;
            splitContainer.Panel2MinSize = 100;
            splitContainer.SplitterWidth = 8;
            splitContainer.BackColor = Color.Gray;

            Panel controlPanel = new Panel();
            controlPanel.Dock = DockStyle.Fill;
            controlPanel.BackColor = Color.LightGray;
            controlPanel.Padding = new Padding(10);

            Label lblCanvasWidth = new Label();
            lblCanvasWidth.Text = "📏 Editor Width (pixels):";
            lblCanvasWidth.Location = new Point(10, 15);
            lblCanvasWidth.Size = new Size(150, 20);
            lblCanvasWidth.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            nudCanvasWidth = new NumericUpDown();
            nudCanvasWidth.Location = new Point(165, 12);
            nudCanvasWidth.Size = new Size(80, 25);
            nudCanvasWidth.Minimum = 150;
            nudCanvasWidth.Maximum = 500;
            nudCanvasWidth.Value = 240; // Starting width for 58mm thermal paper
            nudCanvasWidth.Increment = 10;
            nudCanvasWidth.ValueChanged += NudCanvasWidth_ValueChanged;

            Button btnTestPrint = new Button();
            btnTestPrint.Text = "🖨️ Test Print";
            btnTestPrint.Location = new Point(255, 10);
            btnTestPrint.Size = new Size(100, 30);
            btnTestPrint.BackColor = Color.LightBlue;
            btnTestPrint.Click += BtnTestPrint_Click;

            Label lblInstructions = new Label();
            lblInstructions.Text = "💡 Adjust editor width → Print test → Compare line breaks → Repeat until perfect match";
            lblInstructions.Location = new Point(10, 45);
            lblInstructions.Size = new Size(600, 40);
            lblInstructions.ForeColor = Color.DarkBlue;

            controlPanel.Controls.AddRange(new Control[] {
                lblCanvasWidth, nudCanvasWidth, btnTestPrint, lblInstructions
            });

            splitContainer.Panel2.Controls.Add(controlPanel);
            this.Controls.Add(splitContainer);
        }

        // FIXED: Immediate canvas width adjustment that actually works
        private void NudCanvasWidth_ValueChanged(object sender, EventArgs e)
        {
            if (documentReady)
            {
                SetEditorWidth((int)nudCanvasWidth.Value);
            }
        }

        // FIXED: Working method to set editor width with proper CSS injection
        private void SetEditorWidth(int widthPixels)
        {
            try
            {
                if (htmlEditor?.Document == null) return;

                // FIXED: Direct DOM manipulation approach
                var thermalPaper = htmlEditor.Document.GetElementById("thermal-container");
                var editor = htmlEditor.Document.GetElementById("editor");

                if (thermalPaper != null && editor != null)
                {
                    // Set container width
                    thermalPaper.Style = $"width: {widthPixels}px; min-height: 300px; background: white; margin: 0 auto; padding: 10px; border: 2px solid #333; box-shadow: 0 0 10px rgba(0,0,0,0.3); font-family: 'Courier New', monospace; font-size: 12px; line-height: 1.2;";

                    // Set editor width
                    int editorWidth = widthPixels - 20;
                    editor.Style = $"width: {editorWidth}px; min-height: 250px; outline: none; font-size: 12px; line-height: 1.2; font-family: 'Courier New', monospace; word-wrap: break-word; overflow-wrap: break-word; white-space: pre-wrap; padding: 5px; border: 1px dashed #ccc;";

                    System.Diagnostics.Debug.WriteLine($"Editor width set to: {editorWidth}px (container: {widthPixels}px)");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting editor width: {ex.Message}");
            }
        }

        private void BtnTestPrint_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(cmbPrinters.Text))
                {
                    MessageBox.Show("Please select a printer first!");
                    return;
                }

                string testContent = @"<p><b>Bold</b>, <i>italic</i>, <u>underlined</u> text works!</p>
                                     <p>This is a longer line of text to test where line breaks occur on your thermal printer versus the editor display area.</p>
                                     <p>ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890</p>
                                     <p>Width test: 123456789012345678901234567890123456789012345678901234567890</p>";

                PrintTestContent(testContent);
                MessageBox.Show($"Test print sent!\nCurrent editor width: {nudCanvasWidth.Value}px\n\nCompare line breaks: editor vs printed output.", "Test Print");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Test print failed: {ex.Message}", "Error");
            }
        }

        private void InitializeHtmlEditor()
        {
            htmlEditor = new WebBrowser();
            htmlEditor.Dock = DockStyle.Fill;
            htmlEditor.DocumentCompleted += HtmlEditor_DocumentCompleted;
            htmlEditor.ScriptErrorsSuppressed = true;
            splitContainer.Panel1.Controls.Add(htmlEditor);

            // FIXED: HTML with proper thermal printer font matching and fixed width container
            string editorHtml = @"<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body { 
            margin: 0;
            padding: 20px;
            background: #f0f0f0;
            display: flex;
            justify-content: center;
            font-family: 'Courier New', monospace;
        }
        /* FIXED: Thermal paper container with proper monospace font */
        #thermal-container {
            width: 240px;
            min-height: 300px;
            background: white;
            margin: 0 auto;
            padding: 10px;
            border: 2px solid #333;
            box-shadow: 0 0 10px rgba(0,0,0,0.3);
            font-family: 'Courier New', monospace;
            font-size: 12px;
            line-height: 1.2;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin: 10px 0; 
            font-size: 10px;
            font-family: 'Courier New', monospace;
        }
        td, th { 
            border: 1px solid #666; 
            padding: 2px 4px; 
            text-align: left;
            font-family: 'Courier New', monospace;
        }
        th { 
            background-color: #f0f0f0; 
            font-weight: bold; 
        }
        #editor { 
            width: 220px;
            min-height: 250px; 
            outline: none; 
            font-size: 12px;
            line-height: 1.2;
            font-family: 'Courier New', monospace;
            word-wrap: break-word;
            overflow-wrap: break-word;
            white-space: pre-wrap;
            padding: 5px;
            border: 1px dashed #ccc;
        }
        img { 
            max-width: 100%; 
            height: auto; 
            display: block; 
            margin: 5px auto; 
            border: 1px solid #ccc;
        }
        /* FIXED: Font size classes that match ESC/POS thermal printer scaling */
        .font-size-8 { font-size: 8px !important; line-height: 1.2; font-weight: normal; font-family: 'Courier New', monospace; }
        .font-size-10 { font-size: 10px !important; line-height: 1.2; font-weight: normal; font-family: 'Courier New', monospace; }
        .font-size-12 { font-size: 12px !important; line-height: 1.2; font-weight: normal; font-family: 'Courier New', monospace; }
        .font-size-14 { font-size: 14px !important; line-height: 1.3; font-weight: bold; font-family: 'Courier New', monospace; }
        .font-size-16 { font-size: 16px !important; line-height: 1.3; font-weight: bold; font-family: 'Courier New', monospace; }
        .font-size-18 { font-size: 18px !important; line-height: 1.3; font-weight: bold; font-family: 'Courier New', monospace; }
        .font-size-20 { font-size: 20px !important; line-height: 1.4; font-weight: bold; letter-spacing: 1px; font-family: 'Courier New', monospace; }
        .font-size-24 { font-size: 24px !important; line-height: 1.4; font-weight: bold; letter-spacing: 1px; font-family: 'Courier New', monospace; }
        .font-size-28 { font-size: 28px !important; line-height: 1.4; font-weight: bold; letter-spacing: 1px; font-family: 'Courier New', monospace; }
        .font-size-36 { font-size: 36px !important; line-height: 1.4; font-weight: bold; letter-spacing: 1px; font-family: 'Courier New', monospace; }
    </style>
</head>
<body>
    <div id='thermal-container'>
        <div id='editor' contenteditable='true'>
            <p>🎫 Thermal Print Editor - Courier New Font</p>
            <p><b>Bold</b>, <i>italic</i>, <u>underlined</u> text works!</p>
            <p>This line tests wrapping at editor width boundary.</p>
            <p>Adjust width until editor matches printed line breaks exactly.</p>
            <p style='text-align: center;'>Centered text example</p>
            <p style='text-align: right;'>Right-aligned text</p>
        </div>
    </div>
</body>
<script>
    var documentReady = false;
    
    function setDocumentReady() {
        documentReady = true;
        console.log('Document ready with Courier New font');
    }
    
    function applyFontSize(size) {
        if (!documentReady) return 'error: document not ready';
        
        try {
            var selection = window.getSelection();
            if (!selection.rangeCount || selection.isCollapsed) {
                return 'error: Please select text first';
            }
            
            var range = selection.getRangeAt(0);
            var selectedText = selection.toString();
            
            var span = document.createElement('span');
            span.className = 'font-size-' + size;
            span.setAttribute('data-font-size', size);
            span.textContent = selectedText;
            
            range.deleteContents();
            range.insertNode(span);
            selection.removeAllRanges();
            
            return 'success: Applied ' + size + 'px font size';
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    function insertTable(rows, cols) {
        if (!documentReady) return 'error: document not ready';
        
        try {
            var editor = document.getElementById('editor');
            if (!editor) return 'error: editor not found';
            
            var table = document.createElement('table');
            table.style.width = '100%';
            table.style.borderCollapse = 'collapse';
            table.style.margin = '10px 0';
            table.style.fontFamily = 'Courier New, monospace';
            table.setAttribute('data-table', 'true');
            
            var headerRow = table.insertRow();
            for (var i = 0; i < cols; i++) {
                var th = document.createElement('th');
                th.textContent = 'Header ' + (i + 1);
                th.style.border = '1px solid #666';
                th.style.padding = '2px 4px';
                th.style.backgroundColor = '#f0f0f0';
                th.style.fontWeight = 'bold';
                th.style.fontFamily = 'Courier New, monospace';
                headerRow.appendChild(th);
            }
            
            for (var r = 1; r < rows; r++) {
                var row = table.insertRow();
                for (var c = 0; c < cols; c++) {
                    var td = row.insertCell();
                    td.textContent = 'Data ' + r + ',' + (c + 1);
                    td.style.border = '1px solid #666';
                    td.style.padding = '2px 4px';
                    td.style.fontFamily = 'Courier New, monospace';
                }
            }
            
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var range = selection.getRangeAt(0);
                range.deleteContents();
                range.insertNode(table);
                
                var br = document.createElement('p');
                br.appendChild(document.createElement('br'));
                range.insertNode(br);
            } else {
                editor.appendChild(table);
                var br = document.createElement('p');
                br.appendChild(document.createElement('br'));
                editor.appendChild(br);
            }
            
            return 'success: Table inserted';
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    function insertImage(dataUri) {
        if (!documentReady) return 'error: document not ready';
        
        try {
            var img = document.createElement('img');
            img.src = dataUri;
            img.alt = 'Thermal Image';
            img.style.maxWidth = '100%';
            img.style.height = 'auto';
            img.style.display = 'block';
            img.style.margin = '10px auto';
            
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var range = selection.getRangeAt(0);
                range.deleteContents();
                range.insertNode(img);
            } else {
                document.getElementById('editor').appendChild(img);
            }
            
            return 'success: Image inserted';
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
            var command = 'justify' + align;
            document.execCommand(command, false, null);
            
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var range = selection.getRangeAt(0);
                var parentElement = range.commonAncestorContainer.parentElement;
                if (parentElement && (parentElement.tagName === 'P' || parentElement.tagName === 'DIV')) {
                    parentElement.style.textAlign = align.toLowerCase();
                }
            }
            
            return 'success: Applied ' + align + ' alignment';
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    function insertCharacter(char) {
        if (!documentReady) return 'error: document not ready';
        
        try {
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var range = selection.getRangeAt(0);
                var textNode = document.createTextNode(char);
                range.deleteContents();
                range.insertNode(textNode);
                range.setStartAfter(textNode);
                range.setEndAfter(textNode);
                selection.removeAllRanges();
                selection.addRange(range);
            } else {
                document.getElementById('editor').innerHTML += char;
            }
            return 'success: character inserted';
        } catch(e) {
            return 'error: ' + e.message;
        }
    }
    
    function clearAll() {
        if (!documentReady) return 'error: document not ready';
        
        try {
            document.getElementById('editor').innerHTML = '<p>Start typing your thermal print document here...</p>';
            return 'success: content cleared';
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
    
    window.onload = function() { setDocumentReady(); };
    document.addEventListener('DOMContentLoaded', function() { setDocumentReady(); });
</script>
</html>";

            htmlEditor.DocumentText = editorHtml;
        }

        private void HtmlEditor_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            documentReady = true;
            // Apply initial editor width
            SetEditorWidth((int)nudCanvasWidth.Value);
        }

        // [Include all your existing toolbar initialization and other methods from previous code]
        private void InitializeToolbar()
        {
            toolStrip = new ToolStrip();
            toolStrip.Dock = DockStyle.Top;
            toolStrip.Height = 50;
            this.Controls.Add(toolStrip);

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

            var btnSaveTemplate = new ToolStripButton("Save Template") { BackColor = Color.LightBlue };
            var btnLoadTemplate = new ToolStripButton("Load Template") { BackColor = Color.LightYellow };
            var btnClearAll = new ToolStripButton("Clear All") { BackColor = Color.LightCoral };

            var btnPrint = new ToolStripButton("PRINT") { BackColor = Color.LightGreen };
            var lblPrinter = new ToolStripLabel("Printer:");
            cmbPrinters = new ToolStripComboBox { Width = 200 };

            toolStrip.Items.AddRange(new ToolStripItem[] {
                btnBold, btnItalic, btnUnderline, new ToolStripSeparator(),
                lblSize, cmbFontSize, new ToolStripSeparator(),
                btnLeft, btnCenter, btnRight, new ToolStripSeparator(),
                btnTable, btnImage, new ToolStripSeparator(),
                btnSaveTemplate, btnLoadTemplate, btnClearAll, new ToolStripSeparator(),
                lblPrinter, cmbPrinters, btnPrint
            });

            toolStrip2 = new ToolStrip();
            toolStrip2.Dock = DockStyle.Top;
            toolStrip2.Height = 35;
            this.Controls.Add(toolStrip2);

            var lblBoxChars = new ToolStripLabel("PC850 Box:");
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

            var lblSpecial = new ToolStripLabel("Special:");
            var btnDegree = new ToolStripButton("°") { Font = new Font("Courier New", 12) };
            var btnSection = new ToolStripButton("§") { Font = new Font("Courier New", 12) };

            toolStrip2.Items.AddRange(new ToolStripItem[] {
                lblBoxChars, btnTopLeft, btnTopRight, btnBottomLeft, btnBottomRight,
                btnHorizontal, btnVertical, btnCross, btnTeeUp, btnTeeDown, btnTeeLeft, btnTeeRight,
                new ToolStripSeparator(), lblSpecial, btnDegree, btnSection
            });

            // Event handlers
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
            btnSaveTemplate.Click += BtnSaveTemplate_Click;
            btnLoadTemplate.Click += BtnLoadTemplate_Click;
            btnClearAll.Click += BtnClearAll_Click;
            btnPrint.Click += BtnPrint_Click;

            // PC850 character events
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
        }

        // [Include all remaining methods: InvokeScriptSafely, BtnTable_Click, BtnImage_Click, etc. 
        //  These remain the same as in previous code]

        private string InvokeScriptSafely(string function, params object[] args)
        {
            if (!documentReady) return "error: Editor not ready";

            try
            {
                var result = htmlEditor.Document.InvokeScript(function, args);
                string resultStr = result?.ToString() ?? "error: no result";

                if (resultStr.StartsWith("success:"))
                    System.Diagnostics.Debug.WriteLine($"Script: {resultStr}");

                return resultStr;
            }
            catch (Exception ex)
            {
                string error = $"error: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"Script failed: {error}");
                return error;
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
                        MessageBox.Show($"Table insertion failed: {result}", "Error");
                    else
                        MessageBox.Show("Table inserted successfully!", "Success");
                }
            }
        }

        private void BtnImage_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Images|*.jpg;*.png;*.bmp;*.gif";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
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
                                MessageBox.Show($"Image failed: {result}", "Error");
                            else
                                MessageBox.Show("Image inserted successfully!", "Success");

                            processed.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Image error: {ex.Message}", "Error");
                    }
                }
            }
        }

        private void BtnSaveTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                if (htmlEditor?.Document == null || !documentReady)
                {
                    MessageBox.Show("Editor not ready. Please wait.", "Save Template");
                    return;
                }

                string content = htmlEditor.Document.InvokeScript("getContent").ToString();
                if (string.IsNullOrWhiteSpace(content))
                {
                    MessageBox.Show("No content to save.", "Save Template");
                    return;
                }

                using (SaveFileDialog dlg = new SaveFileDialog())
                {
                    dlg.Filter = "Thermal Print Templates (*.tpt)|*.tpt|HTML Files (*.html)|*.html|All Files (*.*)|*.*";
                    dlg.DefaultExt = "tpt";
                    dlg.Title = "Save Template";
                    dlg.FileName = "ThermalTemplate_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        var template = new
                        {
                            Title = Path.GetFileNameWithoutExtension(dlg.FileName),
                            Created = DateTime.Now,
                            Content = content,
                            EditorWidth = (int)nudCanvasWidth.Value,
                            PrinterType = "Thermal 58mm",
                            Codepage = "PC850"
                        };

                        string jsonTemplate = Newtonsoft.Json.JsonConvert.SerializeObject(template, Newtonsoft.Json.Formatting.Indented);
                        File.WriteAllText(dlg.FileName, jsonTemplate);

                        MessageBox.Show($"Template saved successfully!\n\nFile: {Path.GetFileName(dlg.FileName)}\nEditor Width: {nudCanvasWidth.Value}px",
                                       "Save Template", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving template:\n{ex.Message}", "Save Template Error");
            }
        }

        private void BtnLoadTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                if (htmlEditor?.Document == null || !documentReady)
                {
                    MessageBox.Show("Editor not ready. Please wait.", "Load Template");
                    return;
                }

                using (OpenFileDialog dlg = new OpenFileDialog())
                {
                    dlg.Filter = "Thermal Print Templates (*.tpt)|*.tpt|HTML Files (*.html)|*.html|All Files (*.*)|*.*";
                    dlg.Title = "Load Template";

                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        string fileContent = File.ReadAllText(dlg.FileName);
                        string htmlContent = "";
                        int editorWidth = 240;

                        try
                        {
                            dynamic template = Newtonsoft.Json.JsonConvert.DeserializeObject(fileContent);
                            htmlContent = template.Content;
                            editorWidth = template.EditorWidth ?? 240;

                            MessageBox.Show($"Loading template: {template.Title}\nCreated: {template.Created}\nEditor Width: {editorWidth}px",
                                           "Template Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch
                        {
                            htmlContent = fileContent;
                        }

                        var result = MessageBox.Show("This will replace current content. Continue?",
                                                   "Load Template", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            nudCanvasWidth.Value = editorWidth;

                            htmlEditor.Document.InvokeScript("clearAll");
                            System.Threading.Thread.Sleep(100);

                            var editorDiv = htmlEditor.Document.GetElementById("editor");
                            if (editorDiv != null)
                            {
                                editorDiv.InnerHtml = htmlContent;
                                MessageBox.Show("Template loaded successfully!", "Load Template");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading template:\n{ex.Message}", "Load Template Error");
            }
        }

        private void InsertCharacter(string character)
        {
            try
            {
                if (htmlEditor?.Document != null && documentReady)
                {
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

        private void BtnClearAll_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Clear all content?", "Clear All", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                try
                {
                    htmlEditor.Document.InvokeScript("clearAll");
                    MessageBox.Show("Content cleared!", "Success");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error clearing: {ex.Message}", "Error");
                }
            }
        }

        private Bitmap ConvertImageForThermalPrinter(Image source)
        {
            int maxWidth = 200;
            int width = Math.Min(source.Width, maxWidth);
            int height = (int)(width * (double)source.Height / source.Width);

            var scaled = new Bitmap(source, width, height);
            var monochrome = new Bitmap(width, height);

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    var pixel = scaled.GetPixel(x, y);
                    int gray = (int)(0.299 * pixel.R + 0.587 * pixel.G + 0.114 * pixel.B);
                    int threshold = 140;
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
                PrintContentFixed(html);
                MessageBox.Show("Printing completed successfully!", "Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Printing failed: {ex.Message}", "Print Error");
            }
        }

        private void PrintTestContent(string testHtml)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(testHtml);

            var commands = new List<byte>();
            var printer = new EPSON();

            commands.AddRange(printer.Initialize());
            commands.Add(0x1B); commands.Add(0x74); commands.Add(0x02);
            commands.AddRange(printer.SetStyles(PrintStyle.None));
            commands.AddRange(printer.LeftAlign());

            ProcessHtmlNodeFixed(doc.DocumentNode, commands, printer);
            commands.AddRange(printer.PartialCutAfterFeed(3));

            RawPrinterHelper.SendBytesToPrinter(cmbPrinters.Text, commands.ToArray());
        }

        private void PrintContentFixed(string html)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var commands = new List<byte>();
            var printer = new EPSON();

            commands.AddRange(printer.Initialize());
            commands.Add(0x1B); commands.Add(0x74); commands.Add(0x02);
            commands.AddRange(printer.SetStyles(PrintStyle.None));
            commands.AddRange(printer.LeftAlign());

            ProcessHtmlNodeFixed(doc.DocumentNode, commands, printer);
            commands.AddRange(printer.PartialCutAfterFeed(3));

            RawPrinterHelper.SendBytesToPrinter(cmbPrinters.Text, commands.ToArray());
        }

        // [Include all ProcessHtmlNodeFixed, ProcessParagraphFixed, ProcessSpanFixed, 
        //  ProcessTableFixed, ProcessImageFixed, RefreshPrinters methods from previous code]

        private void ProcessHtmlNodeFixed(HtmlNode node, List<byte> commands, EPSON printer)
        {
            foreach (var child in node.ChildNodes)
            {
                try
                {
                    switch (child.Name.ToLower())
                    {
                        case "p":
                        case "div":
                            ProcessParagraphFixed(child, commands, printer);
                            break;
                        case "b":
                        case "strong":
                            commands.AddRange(printer.SetStyles(PrintStyle.Bold));
                            ProcessHtmlNodeFixed(child, commands, printer);
                            commands.AddRange(printer.SetStyles(PrintStyle.None));
                            break;
                        case "i":
                        case "em":
                            commands.AddRange(printer.SetStyles(PrintStyle.Underline));
                            ProcessHtmlNodeFixed(child, commands, printer);
                            commands.AddRange(printer.SetStyles(PrintStyle.None));
                            break;
                        case "u":
                            commands.AddRange(printer.SetStyles(PrintStyle.Underline));
                            ProcessHtmlNodeFixed(child, commands, printer);
                            commands.AddRange(printer.SetStyles(PrintStyle.None));
                            break;
                        case "span":
                            ProcessSpanFixed(child, commands, printer);
                            break;
                        case "table":
                            ProcessTableFixed(child, commands, printer);
                            break;
                        case "img":
                            ProcessImageFixed(child, commands, printer);
                            break;
                        case "#text":
                            if (!string.IsNullOrWhiteSpace(child.InnerText))
                            {
                                string text = System.Net.WebUtility.HtmlDecode(child.InnerText);
                                byte[] encodedText = Encoding.GetEncoding(850).GetBytes(text);
                                commands.AddRange(encodedText);
                            }
                            break;
                        default:
                            if (child.HasChildNodes)
                                ProcessHtmlNodeFixed(child, commands, printer);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error processing {child.Name}: {ex.Message}");
                    if (!string.IsNullOrWhiteSpace(child.InnerText))
                    {
                        string text = System.Net.WebUtility.HtmlDecode(child.InnerText);
                        commands.AddRange(Encoding.GetEncoding(850).GetBytes(text));
                    }
                }
            }
        }

        private void ProcessParagraphFixed(HtmlNode pNode, List<byte> commands, EPSON printer)
        {
            string style = pNode.GetAttributeValue("style", "").ToLower();
            string align = pNode.GetAttributeValue("align", "").ToLower();

            if (style.Contains("text-align: center") || style.Contains("text-align:center") || align == "center")
                commands.AddRange(printer.CenterAlign());
            else if (style.Contains("text-align: right") || style.Contains("text-align:right") || align == "right")
                commands.AddRange(printer.RightAlign());
            else
                commands.AddRange(printer.LeftAlign());

            ProcessHtmlNodeFixed(pNode, commands, printer);
            commands.AddRange(printer.Print("\n"));
        }

        private void ProcessSpanFixed(HtmlNode span, List<byte> commands, EPSON printer)
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
                PrintStyle printStyle = PrintStyle.None;

                if (fontSize >= 28)
                    printStyle = PrintStyle.DoubleHeight | PrintStyle.DoubleWidth | PrintStyle.Bold;
                else if (fontSize >= 20)
                    printStyle = PrintStyle.DoubleHeight | PrintStyle.DoubleWidth;
                else if (fontSize >= 16)
                    printStyle = PrintStyle.DoubleHeight;

                commands.AddRange(printer.SetStyles(printStyle));
                ProcessHtmlNodeFixed(span, commands, printer);
                commands.AddRange(printer.SetStyles(PrintStyle.None));
            }
            else
            {
                ProcessHtmlNodeFixed(span, commands, printer);
            }
        }

        private void ProcessTableFixed(HtmlNode table, List<byte> commands, EPSON printer)
        {
            try
            {
                var rows = table.SelectNodes(".//tr");
                if (rows == null || rows.Count == 0)
                {
                    commands.AddRange(Encoding.GetEncoding(850).GetBytes("[No table rows found]\n"));
                    return;
                }

                var firstRowCells = rows[0].SelectNodes(".//th|.//td");
                if (firstRowCells == null || firstRowCells.Count == 0)
                {
                    commands.AddRange(Encoding.GetEncoding(850).GetBytes("[No table columns found]\n"));
                    return;
                }

                int colCount = firstRowCells.Count;
                int totalWidth = 32;
                int borderChars = colCount + 1;
                int availableWidth = totalWidth - borderChars;
                int colWidth = Math.Max(4, availableWidth / colCount);

                string topBorder = "╔" + string.Join("╦", Enumerable.Repeat(new string('═', colWidth), colCount)) + "╗";
                string headerSeparator = "╠" + string.Join("╬", Enumerable.Repeat(new string('═', colWidth), colCount)) + "╣";
                string bottomBorder = "╚" + string.Join("╩", Enumerable.Repeat(new string('═', colWidth), colCount)) + "╝";

                commands.AddRange(printer.LeftAlign());
                commands.AddRange(Encoding.GetEncoding(850).GetBytes(topBorder + "\n"));

                bool isFirstRow = true;
                int rowCount = 0;

                foreach (var row in rows)
                {
                    try
                    {
                        var cells = row.SelectNodes(".//th|.//td");
                        if (cells == null || cells.Count == 0) continue;

                        rowCount++;

                        StringBuilder line = new StringBuilder("║");

                        for (int i = 0; i < colCount; i++)
                        {
                            string cellText = "";
                            if (i < cells.Count)
                            {
                                cellText = System.Net.WebUtility.HtmlDecode(cells[i].InnerText?.Trim() ?? "");
                            }

                            if (cellText.Length > colWidth)
                                cellText = cellText.Substring(0, colWidth);

                            line.Append(cellText.PadRight(colWidth));
                            line.Append("║");
                        }

                        string rowLine = line.ToString();
                        bool isHeaderRow = cells.Any(cell => cell.Name.Equals("th", StringComparison.OrdinalIgnoreCase));

                        if (isHeaderRow)
                        {
                            commands.AddRange(printer.SetStyles(PrintStyle.Bold));
                            commands.AddRange(Encoding.GetEncoding(850).GetBytes(rowLine + "\n"));
                            commands.AddRange(printer.SetStyles(PrintStyle.None));

                            if (isFirstRow && rows.Count > 1)
                            {
                                commands.AddRange(Encoding.GetEncoding(850).GetBytes(headerSeparator + "\n"));
                                isFirstRow = false;
                            }
                        }
                        else
                        {
                            commands.AddRange(Encoding.GetEncoding(850).GetBytes(rowLine + "\n"));
                        }
                    }
                    catch (Exception ex)
                    {
                        commands.AddRange(Encoding.GetEncoding(850).GetBytes($"║[Row Error: {ex.Message}]║\n"));
                    }
                }

                commands.AddRange(Encoding.GetEncoding(850).GetBytes(bottomBorder + "\n\n"));

                System.Diagnostics.Debug.WriteLine($"Table processed successfully: {rowCount} rows, {colCount} columns");
            }
            catch (Exception ex)
            {
                commands.AddRange(Encoding.GetEncoding(850).GetBytes($"[Table Error: {ex.Message}]\n"));
                System.Diagnostics.Debug.WriteLine($"Table processing error: {ex.Message}");
            }
        }

        private void ProcessImageFixed(HtmlNode img, List<byte> commands, EPSON printer)
        {
            try
            {
                var src = img.GetAttributeValue("src", "");
                if (!src.StartsWith("data:image/"))
                {
                    commands.AddRange(printer.Print("[No Image Data]\n"));
                    return;
                }

                var base64 = src.Split(',')[1];
                var bytes = Convert.FromBase64String(base64);

                using (var ms = new MemoryStream(bytes))
                using (var image = Image.FromStream(ms))
                {
                    var bitmap = (Bitmap)image;

                    commands.AddRange(printer.CenterAlign());

                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        var lineData = new List<byte>();

                        for (int x = 0; x < bitmap.Width; x += 8)
                        {
                            byte pixelByte = 0;
                            for (int bit = 0; bit < 8 && x + bit < bitmap.Width; bit++)
                            {
                                var pixel = bitmap.GetPixel(x + bit, y);
                                if (pixel.R == 0)
                                    pixelByte |= (byte)(1 << (7 - bit));
                            }
                            lineData.Add(pixelByte);
                        }

                        if (lineData.Any(b => b != 0))
                        {
                            commands.Add(0x1B); commands.Add(0x2A); commands.Add(0x00);
                            commands.Add((byte)(lineData.Count % 256));
                            commands.Add((byte)(lineData.Count / 256));
                            commands.AddRange(lineData);
                        }
                        commands.Add(0x0A);
                    }

                    commands.AddRange(printer.LeftAlign());
                    commands.AddRange(printer.Print("\n"));
                }
            }
            catch (Exception ex)
            {
                commands.AddRange(printer.Print($"[Image Error: {ex.Message}]\n"));
            }
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
            this.Text = "58mm Thermal Printer Editor - WORKING WIDTH ADJUSTMENT & MATCHING FONTS";
            this.WindowState = FormWindowState.Maximized;
            this.ResumeLayout(false);
        }
    }

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
