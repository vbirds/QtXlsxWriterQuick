#ifndef XLSXCONTROL_H
#define XLSXCONTROL_H

#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"
#include "xlsxworksheet.h"

QTXLSX_USE_NAMESPACE

class Xlsxcontrol
{

public:
    Xlsxcontrol();
    void writeHorizontalAlignCell(Document &xlsx, const QString &cell, const QString &text, Format::HorizontalAlignment align);
    void writeVerticalAlignCell(Document &xlsx, const QString &range, const QString &text, Format::VerticalAlignment align);
    void writeBorderStyleCell(Document &xlsx, const QString &cell, const QString &text, Format::BorderStyle bs);
    void writeSolidFillCell(Document &xlsx, const QString &cell, const QColor &color);
    void writePatternFillCell(Document &xlsx, const QString &cell, Format::FillPattern pattern, const QColor &color);
    void writeBorderAndFontColorCell(Document &xlsx, const QString &cell, const QString &text, const QColor &color);
    void writeFontNameCell(Document &xlsx, const QString &cell, const QString &text);
    void writeFontSizeCell(Document &xlsx, const QString &cell, int size);
    void writeInternalNumFormatsCell(Document &xlsx, int row, double value, int numFmt);
    void writeCustomNumFormatsCell(Document &xlsx, int row, double value, const QString &numFmt);
    void Saveas(Document &xlsx,QString &title);
private:
    Document xlsx;
};

#endif // XLSXCONTROL_H
