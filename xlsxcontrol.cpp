#include "xlsxcontrol.h"


Xlsxcontrol::Xlsxcontrol()
{

}

void Xlsxcontrol::writeHorizontalAlignCell(Document &xlsx, const QString &cell, const QString &text, Format::HorizontalAlignment align)
{
    Format format;
    format.setHorizontalAlignment(align);
    format.setBorderStyle(Format::BorderThin);
    xlsx.write(cell, text, format);
}

void Xlsxcontrol::writeVerticalAlignCell(Document &xlsx, const QString &range, const QString &text, Format::VerticalAlignment align)
{
    Format format;
    format.setVerticalAlignment(align);
    format.setBorderStyle(Format::BorderThin);
    CellRange r(range);
    xlsx.write(r.firstRow(), r.firstColumn(), text);
    xlsx.mergeCells(r, format);
}

void Xlsxcontrol::writeBorderStyleCell(Document &xlsx, const QString &cell, const QString &text, Format::BorderStyle bs)
{
    Format format;
    format.setBorderStyle(bs);
    xlsx.write(cell, text, format);
}

void Xlsxcontrol::writeSolidFillCell(Document &xlsx, const QString &cell, const QColor &color)
{
    Format format;
    format.setPatternBackgroundColor(color);
    xlsx.write(cell, QVariant(), format);
}

void Xlsxcontrol::writePatternFillCell(Document &xlsx, const QString &cell, Format::FillPattern pattern, const QColor &color)
{
    Format format;
    format.setPatternForegroundColor(color);
    format.setFillPattern(pattern);
    xlsx.write(cell, QVariant(), format);
}

void Xlsxcontrol::writeBorderAndFontColorCell(Document &xlsx, const QString &cell, const QString &text, const QColor &color)
{
    Format format;
    format.setBorderStyle(Format::BorderThin);
    format.setBorderColor(color);
    format.setFontColor(color);
    xlsx.write(cell, text, format);
}

void Xlsxcontrol::writeFontNameCell(Document &xlsx, const QString &cell, const QString &text)
{
    Format format;
    format.setFontName(text);
    format.setFontSize(16);
    xlsx.write(cell, text, format);
}

void Xlsxcontrol::writeFontSizeCell(Document &xlsx, const QString &cell, int size)
{
    Format format;
    format.setFontSize(size);
    xlsx.write(cell, "Qt Xlsx", format);
}

void Xlsxcontrol::writeInternalNumFormatsCell(Document &xlsx, int row, double value, int numFmt)
{
    Format format;
    format.setNumberFormatIndex(numFmt);
    xlsx.write(row, 1, value);
    xlsx.write(row, 2, QString("Builtin NumFmt %1").arg(numFmt));
    xlsx.write(row, 3, value, format);
}

void Xlsxcontrol::writeCustomNumFormatsCell(Document &xlsx, int row, double value, const QString &numFmt)
{
    Format format;
    format.setNumberFormat(numFmt);
    xlsx.write(row, 1, value);
    xlsx.write(row, 2, numFmt);
    xlsx.write(row, 3, value, format);
}

void Xlsxcontrol::Saveas(Document &xlsx,QString &title)
{

    xlsx.saveAs(title);
}
