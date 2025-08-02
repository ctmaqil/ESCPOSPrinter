using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;

public abstract class DesignElement
{
    [Category("Position")]
    public Rectangle Bounds { get; set; }

    public abstract void Draw(Graphics g);
}




[DisplayName("Text")]
public class TextElement : DesignElement
{
    [Category("Text")]
    public string Text { get; set; } = "Sample Text";

    [Category("Appearance")]
    public Font Font { get; set; } = new Font("Arial", 12);

    [Category("Appearance")]
    public System.Drawing.Color Color { get; set; } = System.Drawing.Color.Black;

    [Category("Appearance")]
    [DisplayName("Text Alignment")]
    public StringAlignment Alignment { get; set; } = StringAlignment.Near;

    public override void Draw(Graphics g)
    {
        using (Brush brush = new SolidBrush(Color))
        {
            StringFormat format = new StringFormat();
            format.Alignment = Alignment;
            format.LineAlignment = StringAlignment.Center;
            format.FormatFlags = StringFormatFlags.NoWrap;

            // Ensure text fits within bounds
            Rectangle textBounds = Bounds;

            // Draw text with exact positioning
            g.DrawString(Text, Font, brush, textBounds, format);
        }
    }
}

[DisplayName("Image")]
public class ImageElement : DesignElement
{
    private Image _image;
    private string _filePath;

    [Category("Image")]
    [DisplayName("File Path")]
    [ReadOnly(true)]
    public string FilePath
    {
        get => _filePath;
        set => _filePath = value;
    }

    [Category("Image")]
    [Browsable(false)]
    public Image Image
    {
        get => _image;
        set => _image = value;
    }

    [Category("Appearance")]
    [DisplayName("Maintain Aspect Ratio")]
    public bool MaintainAspectRatio { get; set; } = true;

    public override void Draw(Graphics g)
    {
        if (_image != null)
        {
            // Set high-quality image rendering
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.CompositingQuality = CompositingQuality.HighQuality;

            // Draw image exactly as positioned on canvas
            g.DrawImage(_image, Bounds);
        }
        else
        {
            // Draw placeholder for missing image
            using (Brush brush = new SolidBrush(System.Drawing.Color.LightGray))
            {
                g.FillRectangle(brush, Bounds);
            }
            using (Pen pen = new Pen(System.Drawing.Color.Gray, 1))
            {
                g.DrawRectangle(pen, Bounds);
            }
            using (Brush textBrush = new SolidBrush(System.Drawing.Color.DarkGray))
            {
                StringFormat format = new StringFormat();
                format.Alignment = StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                g.DrawString("No Image", SystemFonts.DefaultFont, textBrush, Bounds, format);
            }
        }
    }
}



[DisplayName("Table")]
public class TableElement : DesignElement
{
    [Category("Table")]
    public int Rows { get; set; } = 2;

    [Category("Table")]
    public int Columns { get; set; } = 2;

    [Category("Appearance")]
    [DisplayName("Border Color")]
    public System.Drawing.Color BorderColor { get; set; } = System.Drawing.Color.Black; // Fixed: Use full namespace

    [Category("Appearance")]
    [DisplayName("Background Color")]
    public System.Drawing.Color BackgroundColor { get; set; } = System.Drawing.Color.White; // Fixed: Use full namespace

    [Category("Text")]
    public Font Font { get; set; } = new Font("Arial", 10);

    [Browsable(false)]
    public string[,] CellData { get; set; }

    public override void Draw(Graphics g)
    {
        if (CellData == null)
        {
            CellData = new string[Rows, Columns];
            for (int r = 0; r < Rows; r++)
            {
                for (int c = 0; c < Columns; c++)
                {
                    CellData[r, c] = $"Cell {r + 1},{c + 1}";
                }
            }
        }

        // Draw table background
        using (Brush bgBrush = new SolidBrush(BackgroundColor))
        {
            g.FillRectangle(bgBrush, Bounds);
        }

        // Calculate cell dimensions
        float cellWidth = (float)Bounds.Width / Columns;
        float cellHeight = (float)Bounds.Height / Rows;

        using (Pen borderPen = new Pen(BorderColor, 1))
        using (Brush textBrush = new SolidBrush(System.Drawing.Color.Black)) // Fixed: Use full namespace
        {
            // Draw grid lines and content
            for (int r = 0; r < Rows; r++)
            {
                for (int c = 0; c < Columns; c++)
                {
                    RectangleF cellRect = new RectangleF(
                        Bounds.X + c * cellWidth,
                        Bounds.Y + r * cellHeight,
                        cellWidth,
                        cellHeight);

                    // Draw cell border
                    g.DrawRectangle(borderPen, Rectangle.Round(cellRect));

                    // Draw cell text
                    if (CellData != null && r < CellData.GetLength(0) && c < CellData.GetLength(1))
                    {
                        StringFormat format = new StringFormat();
                        format.Alignment = StringAlignment.Center;
                        format.LineAlignment = StringAlignment.Center;

                        g.DrawString(CellData[r, c] ?? "", Font, textBrush, cellRect, format);
                    }
                }
            }
        }
    }
}
