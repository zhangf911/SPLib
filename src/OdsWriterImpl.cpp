// File: OdsWriterImpl.cpp
// OdsWriterImpl implementation file
//

#include "splib.h"
#include <stdio.h>
#include "OdsWriterImpl.h"
#include "ZipArchive.h"
#include "ToUTF8.h"
#include "Strings.h"
#include "Formulas.h"
#include "Util.h"
#include "splibint.h"

namespace splib {

void OdsWriterImpl::write(Spreadsheet& spreadsheet, const _TCHAR* pathname) {
    ZipArchive ar(pathname);
    writeManifest(ar);
    writeContent(spreadsheet, ar);
    writeMeta(ar);
    writeMimetype(ar);
    writeSettings(ar);
    writeStyles(ar);
    ar.close();
}

void OdsWriterImpl::writeManifest(ZipArchive& ar) {
    ar.openEntry("META-INF/manifest.xml");
    ar << "<\?xml version=\"1.0\" encoding=\"UTF-8\"\?>\r\n"
          "<!DOCTYPE manifest:manifest PUBLIC \"-//OpenOffice.org//DTD Manifest 1.0//EN\" \"Manifest.dtd\">\r\n"
          "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">\r\n"
          "<manifest:file-entry manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\" manifest:full-path=\"/\"/>\r\n"
          "<manifest:file-entry manifest:media-type=\"text/xml\" manifest:full-path=\"content.xml\"/>\r\n"
          "<manifest:file-entry manifest:media-type=\"text/xml\" manifest:full-path=\"styles.xml\"/>\r\n"
          "<manifest:file-entry manifest:media-type=\"text/xml\" manifest:full-path=\"meta.xml\"/>\r\n"
          "<manifest:file-entry manifest:media-type=\"text/xml\" manifest:full-path=\"settings.xml\"/>\r\n"
          "</manifest:manifest>\r\n";
    ar.closeEntry();
}

void OdsWriterImpl::writeContent(Spreadsheet& sp, ZipArchive& ar) {
    ar.openEntry("content.xml");
    ar << "<\?xml version=\"1.0\" encoding=\"UTF-8\"\?>\r\n"
          "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" xmlns:ooow=\"http://openoffice.org/2004/writer\" xmlns:oooc=\"http://openoffice.org/2004/calc\" xmlns:dom=\"http://www.w3.org/2001/xml-events\" xmlns:xforms=\"http://www.w3.org/2002/xforms\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" office:version=\"1.0\">\r\n"
          "<office:scripts/>\r\n"
          "<office:font-face-decls>\r\n"
          "<style:font-face style:name=\"Andale Sans UI\" svg:font-family=\"&apos;Andale Sans UI&apos;\" style:font-pitch=\"variable\"/>\r\n"
          "<style:font-face style:name=\"Tahoma\" svg:font-family=\"Tahoma\" style:font-pitch=\"variable\"/>\r\n"
          "<style:font-face style:name=\"Albany\" svg:font-family=\"Albany\" style:font-family-generic=\"swiss\" style:font-pitch=\"variable\"/>\r\n"
          "</office:font-face-decls>\r\n"
          "<office:automatic-styles>\r\n";
    int nextColumnStyle;
    writeColumnStyles(sp, nextColumnStyle, ar);
    int nextRowStyle;
    writeRowStyles(sp, nextRowStyle, ar);
    writeCellStyles(ar);
    ar << "<style:style style:name=\"ta1\" style:family=\"table\" style:master-page-name=\"Default\">\r\n"
          "<style:table-properties table:display=\"true\" style:writing-mode=\"lr-tb\"/>\r\n"
          "</style:style>\r\n"
          "</office:automatic-styles>\r\n"
          "<office:body>\r\n"
          "<office:spreadsheet>\r\n";
    for (int i = 0; i < sp.tableCount(); i++) {
        writeTable(sp.table(i), nextColumnStyle, nextRowStyle, ar);
    }
    ar << "</office:spreadsheet>\r\n"
          "</office:body>\r\n"
          "</office:document-content>\r\n";
    ar.closeEntry();
}

void OdsWriterImpl::writeMeta(ZipArchive& ar) {
    ar.openEntry("meta.xml");
    ar << "<\?xml version=\"1.0\" encoding=\"UTF-8\"\?>\r\n"
          "<office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" office:version=\"1.0\">\r\n"
          "<office:meta>\r\n"
          "<meta:user-defined meta:name=\"Info 1\"/>\r\n"
          "<meta:user-defined meta:name=\"Info 2\"/>\r\n"
          "<meta:user-defined meta:name=\"Info 3\"/>\r\n"
          "<meta:user-defined meta:name=\"Info 4\"/>\r\n"
          "</office:meta>\r\n"
          "</office:document-meta>\r\n";
    ar.closeEntry();
}

void OdsWriterImpl::writeMimetype(ZipArchive& ar) {
    ar.openEntry("mimetype");
    ar << "application/vnd.oasis.opendocument.spreadsheet";
    ar.closeEntry();
}

void OdsWriterImpl::writeSettings(ZipArchive& ar) {
    ar.openEntry("settings.xml");
    ar << "<\?xml version=\"1.0\" encoding=\"UTF-8\"\?>\r\n"
          "<office:document-settings xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:config=\"urn:oasis:names:tc:opendocument:xmlns:config:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" office:version=\"1.0\">\r\n"
          "<office:settings>\r\n"
          "<config:config-item-set config:name=\"ooo:view-settings\">\r\n"
          "<config:config-item config:name=\"VisibleAreaTop\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"VisibleAreaLeft\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"VisibleAreaWidth\" config:type=\"int\">2258</config:config-item>\r\n"
          "<config:config-item config:name=\"VisibleAreaHeight\" config:type=\"int\">451</config:config-item>\r\n"
          "<config:config-item-map-indexed config:name=\"Views\">\r\n"
          "<config:config-item-map-entry>\r\n"
          "<config:config-item config:name=\"ViewId\" config:type=\"string\">View1</config:config-item>\r\n"
          "<config:config-item-map-named config:name=\"Tables\">\r\n"
          "<config:config-item-map-entry config:name=\"Sheet1\">\r\n"
          "<config:config-item config:name=\"CursorPositionX\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"CursorPositionY\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"HorizontalSplitMode\" config:type=\"short\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"VerticalSplitMode\" config:type=\"short\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"HorizontalSplitPosition\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"VerticalSplitPosition\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"ActiveSplitRange\" config:type=\"short\">2</config:config-item>\r\n"
          "<config:config-item config:name=\"PositionLeft\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"PositionRight\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"PositionTop\" config:type=\"int\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"PositionBottom\" config:type=\"int\">0</config:config-item>\r\n"
          "</config:config-item-map-entry>\r\n"
          "</config:config-item-map-named>\r\n"
          "<config:config-item config:name=\"ActiveTable\" config:type=\"string\">Sheet1</config:config-item>\r\n"
          "<config:config-item config:name=\"HorizontalScrollbarWidth\" config:type=\"int\">600</config:config-item>\r\n"
          "<config:config-item config:name=\"ZoomType\" config:type=\"short\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"ZoomValue\" config:type=\"int\">100</config:config-item>\r\n"
          "<config:config-item config:name=\"PageViewZoomValue\" config:type=\"int\">60</config:config-item>\r\n"
          "<config:config-item config:name=\"ShowPageBreakPreview\" config:type=\"boolean\">false</config:config-item>\r\n"
          "<config:config-item config:name=\"ShowZeroValues\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"ShowNotes\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"ShowGrid\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"GridColor\" config:type=\"long\">12632256</config:config-item>\r\n"
          "<config:config-item config:name=\"ShowPageBreaks\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"HasColumnRowHeaders\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"HasSheetTabs\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"IsOutlineSymbolsSet\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"IsSnapToRaster\" config:type=\"boolean\">false</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterIsVisible\" config:type=\"boolean\">false</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterResolutionX\" config:type=\"int\">1000</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterResolutionY\" config:type=\"int\">1000</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterSubdivisionX\" config:type=\"int\">1</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterSubdivisionY\" config:type=\"int\">1</config:config-item>\r\n"
          "<config:config-item config:name=\"IsRasterAxisSynchronized\" config:type=\"boolean\">true</config:config-item>\r\n"
          "</config:config-item-map-entry>\r\n"
          "</config:config-item-map-indexed>\r\n"
          "</config:config-item-set>\r\n"
          "<config:config-item-set config:name=\"ooo:configuration-settings\">\r\n"
          "<config:config-item config:name=\"ShowZeroValues\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"ShowNotes\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"ShowGrid\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"GridColor\" config:type=\"long\">12632256</config:config-item>\r\n"
          "<config:config-item config:name=\"ShowPageBreaks\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"LinkUpdateMode\" config:type=\"short\">3</config:config-item>\r\n"
          "<config:config-item config:name=\"HasColumnRowHeaders\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"HasSheetTabs\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"IsOutlineSymbolsSet\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"IsSnapToRaster\" config:type=\"boolean\">false</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterIsVisible\" config:type=\"boolean\">false</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterResolutionX\" config:type=\"int\">1000</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterResolutionY\" config:type=\"int\">1000</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterSubdivisionX\" config:type=\"int\">1</config:config-item>\r\n"
          "<config:config-item config:name=\"RasterSubdivisionY\" config:type=\"int\">1</config:config-item>\r\n"
          "<config:config-item config:name=\"IsRasterAxisSynchronized\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"AutoCalculate\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"ApplyUserData\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"CharacterCompressionType\" config:type=\"short\">0</config:config-item>\r\n"
          "<config:config-item config:name=\"IsKernAsianPunctuation\" config:type=\"boolean\">false</config:config-item>\r\n"
          "<config:config-item config:name=\"SaveVersionOnClose\" config:type=\"boolean\">false</config:config-item>\r\n"
          "<config:config-item config:name=\"UpdateFromTemplate\" config:type=\"boolean\">false</config:config-item>\r\n"
          "<config:config-item config:name=\"AllowPrintJobCancel\" config:type=\"boolean\">true</config:config-item>\r\n"
          "<config:config-item config:name=\"LoadReadonly\" config:type=\"boolean\">false</config:config-item>\r\n"
          "</config:config-item-set>\r\n"
          "</office:settings>\r\n"
          "</office:document-settings>\r\n";
    ar.closeEntry();
}

void OdsWriterImpl::writeStyles(ZipArchive& ar) {
    ar.openEntry("styles.xml");
    ar << "<\?xml version=\"1.0\" encoding=\"UTF-8\"\?>\r\n"
          "<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" xmlns:math=\"http://www.w3.org/1998/Math/MathML\" xmlns:form=\"urn:oasis:names:tc:opendocument:xmlns:form:1.0\" xmlns:script=\"urn:oasis:names:tc:opendocument:xmlns:script:1.0\" xmlns:ooo=\"http://openoffice.org/2004/office\" xmlns:ooow=\"http://openoffice.org/2004/writer\" xmlns:oooc=\"http://openoffice.org/2004/calc\" xmlns:dom=\"http://www.w3.org/2001/xml-events\" office:version=\"1.0\">\r\n"
          "<office:font-face-decls>\r\n"
          "<style:font-face style:name=\"Andale Sans UI\" svg:font-family=\"&apos;Andale Sans UI&apos;\" style:font-pitch=\"variable\"/>\r\n"
          "<style:font-face style:name=\"Tahoma\" svg:font-family=\"Tahoma\" style:font-pitch=\"variable\"/>\r\n"
          "<style:font-face style:name=\"Albany\" svg:font-family=\"Albany\" style:font-family-generic=\"swiss\" style:font-pitch=\"variable\"/>\r\n"
          "</office:font-face-decls>\r\n"
          "<office:styles>\r\n"
          "<style:default-style style:family=\"table-cell\">\r\n"
          "<style:table-cell-properties style:decimal-places=\"2\"/>\r\n"
          "<style:paragraph-properties style:tab-stop-distance=\"1.25cm\"/>\r\n"
          "<style:text-properties style:font-name=\"Albany\" fo:language=\"en\" fo:country=\"US\" style:font-name-asian=\"Andale Sans UI\" style:language-asian=\"none\" style:country-asian=\"none\" style:font-name-complex=\"Tahoma\" style:language-complex=\"none\" style:country-complex=\"none\"/>\r\n"
          "</style:default-style>\r\n"
          "<number:number-style style:name=\"N0\">\r\n"
          "<number:number number:min-integer-digits=\"1\"/>\r\n"
          "</number:number-style>\r\n"
          "<number:currency-style style:name=\"N104P0\" style:volatile=\"true\">\r\n"
          "<number:number number:decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/>\r\n"
          "</number:currency-style>\r\n"
          "<number:currency-style style:name=\"N104\">\r\n"
          "<style:text-properties fo:color=\"#ff0000\"/>\r\n"
          "<number:text>-</number:text>\r\n"
          "<number:number number:decimal-places=\"2\" number:min-integer-digits=\"1\" number:grouping=\"true\"/>\r\n"
          "<style:map style:condition=\"value()&gt;=0\" style:apply-style-name=\"N104P0\"/>\r\n"
          "</number:currency-style>\r\n"
          "<style:style style:name=\"Default\" style:family=\"table-cell\"/>\r\n"
          "<style:style style:name=\"Result\" style:family=\"table-cell\" style:parent-style-name=\"Default\">\r\n"
          "<style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\"/>\r\n"
          "</style:style>\r\n"
          "<style:style style:name=\"Result2\" style:family=\"table-cell\" style:parent-style-name=\"Result\" style:data-style-name=\"N104\"/>\r\n"
          "<style:style style:name=\"Heading\" style:family=\"table-cell\" style:parent-style-name=\"Default\">\r\n"
          "<style:table-cell-properties style:text-align-source=\"fix\" style:repeat-content=\"false\"/>\r\n"
          "<style:paragraph-properties fo:text-align=\"center\"/>\r\n"
          "<style:text-properties fo:font-size=\"16pt\" fo:font-style=\"italic\" fo:font-weight=\"bold\"/>\r\n"
          "</style:style>\r\n"
          "<style:style style:name=\"Heading1\" style:family=\"table-cell\" style:parent-style-name=\"Heading\">\r\n"
          "<style:table-cell-properties style:rotation-angle=\"90\"/>\r\n"
          "</style:style>\r\n"
          "</office:styles>\r\n"
          "<office:automatic-styles>\r\n"
          "<style:page-layout style:name=\"pm1\">\r\n"
          "<style:page-layout-properties style:writing-mode=\"lr-tb\"/>\r\n"
          "<style:header-style>\r\n"
          "<style:header-footer-properties fo:min-height=\"0.751cm\" fo:margin-left=\"0cm\" fo:margin-right=\"0cm\" fo:margin-bottom=\"0.25cm\"/>\r\n"
          "</style:header-style>\r\n"
          "<style:footer-style>\r\n"
          "<style:header-footer-properties fo:min-height=\"0.751cm\" fo:margin-left=\"0cm\" fo:margin-right=\"0cm\" fo:margin-top=\"0.25cm\"/>\r\n"
          "</style:footer-style>\r\n"
          "</style:page-layout>\r\n"
          "<style:page-layout style:name=\"pm2\">\r\n"
          "<style:page-layout-properties style:writing-mode=\"lr-tb\"/>\r\n"
          "<style:header-style>\r\n"
          "<style:header-footer-properties fo:min-height=\"0.751cm\" fo:margin-left=\"0cm\" fo:margin-right=\"0cm\" fo:margin-bottom=\"0.25cm\" fo:border=\"0.088cm solid #000000\" fo:padding=\"0.018cm\" fo:background-color=\"#c0c0c0\">\r\n"
          "<style:background-image/>\r\n"
          "</style:header-footer-properties>\r\n"
          "</style:header-style>\r\n"
          "<style:footer-style>\r\n"
          "<style:header-footer-properties fo:min-height=\"0.751cm\" fo:margin-left=\"0cm\" fo:margin-right=\"0cm\" fo:margin-top=\"0.25cm\" fo:border=\"0.088cm solid #000000\" fo:padding=\"0.018cm\" fo:background-color=\"#c0c0c0\">\r\n"
          "<style:background-image/>\r\n"
          "</style:header-footer-properties>\r\n"
          "</style:footer-style>\r\n"
          "</style:page-layout>\r\n"
          "</office:automatic-styles>\r\n"
          "<office:master-styles>\r\n"
          "<style:master-page style:name=\"Default\" style:page-layout-name=\"pm1\">\r\n"
          "<style:header>\r\n"
          "<text:p>\r\n"
          "<text:sheet-name>\?\?\?</text:sheet-name>\r\n"
          "</text:p>\r\n"
          "</style:header>\r\n"
          "<style:header-left style:display=\"false\"/>\r\n"
          "<style:footer>\r\n"
          "<text:p>Page <text:page-number>1</text:page-number>\r\n"
          "</text:p>\r\n"
          "</style:footer>\r\n"
          "<style:footer-left style:display=\"false\"/>\r\n"
          "</style:master-page>\r\n"
          "<style:master-page style:name=\"Report\" style:page-layout-name=\"pm2\">\r\n"
          "<style:header>\r\n"
          "<style:region-left>\r\n"
          "<text:p>\r\n"
          "<text:sheet-name>\?\?\?</text:sheet-name> (<text:title>\?\?\?</text:title>)</text:p>\r\n"
          "</style:region-left>\r\n"
          "<style:region-right>\r\n"
          "<text:p>\r\n"
          "</text:p>\r\n"
          "</style:region-right>\r\n"
          "</style:header>\r\n"
          "<style:header-left style:display=\"false\"/>\r\n"
          "<style:footer>\r\n"
          "<text:p>Page <text:page-number>1</text:page-number> / <text:page-count>99</text:page-count>\r\n"
          "</text:p>\r\n"
          "</style:footer>\r\n"
          "<style:footer-left style:display=\"false\"/>\r\n"
          "</style:master-page>\r\n"
          "</office:master-styles>\r\n"
          "</office:document-styles>\r\n";
    ar.closeEntry();
}

void OdsWriterImpl::writeTable(Table& table, int& nextColumnStyle,
                               int& nextRowStyle, ZipArchive& ar) {
    ToUTF8 tableNameUtf8(Strings::xmlize(table.getName()).c_str());
    ar << "<table:table table:name=\"" << tableNameUtf8.get() << "\" table:style-name=\"ta1\" table:print=\"false\">\r\n";
    // columns
    writeColumns(table, nextColumnStyle, ar);
    // rows and cells
    int lastRowIndex = -1;
    Rows::Iterator* rowIt = table.rows().iterator();
    while (rowIt->hasNext()) {
        Rows::Entry entry = rowIt->next();
        int rowIndex = entry.index();
        Row& row = entry.object();
        Cells& cells = row.cells();
        double height = row.getHeight();
        if (cells.size() == 0 && height < 0) {
            continue;
        }
        writeEmptyRows(rowIndex - lastRowIndex - 1, ar);
        lastRowIndex = rowIndex;
        int style = 1;
        if (height >= 0) {
            style = nextRowStyle++;
        }
        ar << "<table:table-row table:style-name=\"ro" << style << "\">\r\n";
        int lastCellIndex = -1;
        Cells::Iterator* cellIt = cells.iterator();
        while (cellIt->hasNext()) {
            Cells::Entry entry = cellIt->next();
            int cellIndex = entry.index();
            Cell& cell = entry.object();
            if (cell.getType() == Cell::NONE) {
                continue;
            }
            writeEmptyCells(cellIndex - lastCellIndex - 1, ar);
            lastCellIndex = cellIndex;
            writeCell(cell, ar);
        }
        delete cellIt;
        ar << "</table:table-row>\r\n";
    }
    delete rowIt;
    ar << "</table:table>\r\n";
}

void OdsWriterImpl::writeRowStyles(Spreadsheet& sp, int& nextRowStyle,
                                   ZipArchive& ar) {
    ar << "<style:style style:name=\"ro1\" style:family=\"table-row\">\r\n"
          "<style:table-row-properties style:row-height=\"0.453cm\" fo:break-before=\"auto\" style:use-optimal-row-height=\"true\"/>\r\n"
          "</style:style>\r\n";
    int k = 2;
    for (int i = 0; i < sp.tableCount(); i++) {
        Table& table = sp.table(i);
        Rows::Iterator* j = table.rows().iterator();
        while (j->hasNext()) {
            double height = j->next().object().getHeight();
            if (height >= 0) {
                ar << "<style:style style:name=\"ro" << k++ << "\" style:family=\"table-row\">\r\n"
                      "<style:table-row-properties style:row-height=\"" << height << "pt\" fo:break-before=\"auto\" style:use-optimal-row-height=\"false\"/>\r\n"
                      "</style:style>\r\n";
            }
        }
        delete j;
    }
    nextRowStyle = 2;
}

void OdsWriterImpl::writeColumnStyles(Spreadsheet& sp, int& nextColumnStyle,
                                      ZipArchive& ar) {
    ar << "<style:style style:name=\"co1\" style:family=\"table-column\">\r\n"
          "<style:table-column-properties fo:break-before=\"auto\" style:column-width=\"2.267cm\"/>\r\n"
          "</style:style>\r\n";
    int k = 2;
    for (int i = 0; i < sp.tableCount(); i++) {
        Table& table = sp.table(i);
        Columns::Iterator* j = table.columns().iterator();
        while (j->hasNext()) {
            double width = j->next().object().getWidth();
            if (width >= 0) {
                ar << "<style:style style:name=\"co" << k++ << "\" style:family=\"table-column\">\r\n"
                      "<style:table-column-properties fo:break-before=\"auto\" style:column-width=\"" << width << "pt\"/>\r\n"
                      "</style:style>\r\n";
            }
        }
        delete j;
    }
    nextColumnStyle = 2;
}

void OdsWriterImpl::writeColumns(Table& table, int& nextColumnStyle,
                                 ZipArchive& ar) {
    int lastIndex = -1;
    Columns::Iterator* i = table.columns().iterator();
    while (i->hasNext()) {
        Columns::Entry entry = i->next();
        int index = entry.index();
        double width = entry.object().getWidth();
        if (width >= 0) {
            writeEmptyColumns(index - lastIndex - 1, ar);
            lastIndex = index;
            ar << "<table:table-column table:style-name=\"co" << nextColumnStyle++ << "\" table:default-cell-style-name=\"Default\"/>\r\n";
        }
    }
    delete i;
}

void OdsWriterImpl::writeEmptyColumns(int columns, ZipArchive& ar) {
    if (columns == 1) {
        ar << "<table:table-column table:style-name=\"co1\" table:default-cell-style-name=\"Default\"/>\r\n";
    } else if (columns > 1) {
        ar << "<table:table-column table:style-name=\"co1\" table:number-columns-repeated=\"" << columns << "\" table:default-cell-style-name=\"Default\"/>\r\n";
    }
}

void OdsWriterImpl::writeEmptyRows(int rows, ZipArchive& ar) {
    if (rows == 1) {
        ar << "<table:table-row table:style-name=\"ro1\"/>\r\n";
    } else if (rows > 1) {
        ar << "<table:table-row table:style-name=\"ro1\" table:number-rows-repeated=\"" << rows << "\"/>\r\n";
    }
}

void OdsWriterImpl::writeCell(Cell& cell, ZipArchive& ar) {
    ar << "<table:table-cell ";
    int style = cellStyleIndex(cell);
    if (style != 0) {
        ar << "table:style-name=\"ce" << style << "\" ";
    }
    if (cell.getType() == Cell::TEXT) {
        ToUTF8 textUtf8(Strings::xmlize(cell.getText()).c_str());
        ar << "office:value-type=\"string\">\r\n"
              "<text:p>" << textUtf8.get() << "</text:p>\r\n";
    } else if (cell.getType() == Cell::LONG) {
        ar << "office:value-type=\"float\" office:value=\"" << cell.getLong() << "\">\r\n"
              "<text:p>" << cell.getLong() << "</text:p>\r\n";
    } else if (cell.getType() == Cell::DOUBLE) {
        ar << "office:value-type=\"float\" office:value=\"" << cell.getDouble() << "\">\r\n"
              "<text:p>" << cell.getDouble() << "</text:p>\r\n";
    } else if (cell.getType() == Cell::DATE) {
        ar << "office:value-type=\"date\" office:date-value=\"" << date(cell.getDate()).c_str() << "\">\r\n";
    } else if (cell.getType() == Cell::TIME) {
        ar << "office:value-type=\"time\" office:time-value=\"" << time(cell.getTime()).c_str() << "\">\r\n";
    } else if (cell.getType() == Cell::FORMULA) {
        ar << "table:formula=\"" << formula(cell.getFormula()).c_str() << "\" office:value-type=\"float\">";
    }
    ar << "</table:table-cell>\r\n";
}

void OdsWriterImpl::writeEmptyCells(int cells, ZipArchive& ar) {
    if (cells == 1) {
        ar << "<table:table-cell/>\r\n";
    } else if (cells > 1) {
        ar << "<table:table-cell table:number-columns-repeated=\"" << cells << "\"/>\r\n";
    }
}

std::basic_string<char> OdsWriterImpl::date(const Date& date) {
    char s[11];
    int year = date.getYear();
    int month = date.getMonth();
    int day = date.getDay();
#pragma warning (disable : 4996)
    ::sprintf(s, "%04d-%02d-%02d", year, month, day);
#pragma warning (default : 4996)
    return std::basic_string<char>(s);
}

std::basic_string<char> OdsWriterImpl::time(const Time& time) {
    char s[12];
    int hours = time.getHours();
    int minutes = time.getMinutes();
    int seconds = time.getSeconds();    
#pragma warning (disable : 4996)
    ::sprintf(s, "PT%02dH%02dM%02dS", hours, minutes, seconds);
#pragma warning (default : 4996)
    return std::basic_string<char>(s);
}

std::basic_string<char> OdsWriterImpl::formula(const _TCHAR* formula) {
    std::basic_string<char> res;
    int c1, r1, c2, r2;
    Formulas::parse(formula, c1, r1, c2, r2);
    res += "oooc:=SUM([.";
    res += Util::buildLocation(c1, r1);
    res += ":.";
    res += Util::buildLocation(c2, r2);
    res += "])";
    return res;
}

void OdsWriterImpl::writeCellStyles(ZipArchive& ar) {
    ar << "<number:date-style style:name=\"N37\" number:automatic-order=\"true\">\r\n"
          "<number:month number:style=\"long\"/>\r\n"
          "<number:text>/</number:text>\r\n"
          "<number:day number:style=\"long\"/>\r\n"
          "<number:text>/</number:text>\r\n"
          "<number:year/>\r\n"
          "</number:date-style>\r\n"
          "<number:time-style style:name=\"N43\">\r\n"
          "<number:hours number:style=\"long\"/>\r\n"
          "<number:text>:</number:text>\r\n"
          "<number:minutes number:style=\"long\"/>\r\n"
          "<number:text>:</number:text>\r\n"
          "<number:seconds number:style=\"long\"/>\r\n"
          "<number:text> </number:text>\r\n"
          "<number:am-pm/>\r\n"
          "</number:time-style>\r\n";
    Cell::HAlignment hAligns[] = {Cell::HADEFAULT, Cell::LEFT, Cell::CENTER,
        Cell::RIGHT, Cell::JUSTIFIED, Cell::FILLED};
    Cell::VAlignment vAligns[] = {Cell::VADEFAULT, Cell::TOP, Cell::MIDDLE,
        Cell::BOTTOM};
    int style = 0;
    for (int i = 0; i < 3; i++) { // 0 - general, 1 - date, 2 - time
        for (int j = 0; j < 6; j++) {
            for (int k = 0; k < 4; k++, style++) {
                if (style == 0) {
                    continue;
                }
                const char* dataStyle = "";
                switch (i) {
                    case 1: dataStyle = " style:data-style-name=\"N37\""; break;
                    case 2: dataStyle = " style:data-style-name=\"N43\""; break;
                }
                ar << "<style:style style:name=\"ce" << style << "\" style:family=\"table-cell\" style:parent-style-name=\"Default\""
                   << dataStyle << ">\r\n";

                Cell::HAlignment hAlign = hAligns[j];
                Cell::VAlignment vAlign = vAligns[k];

                ar << "<style:table-cell-properties";
                if (hAlign != Cell::HADEFAULT) {
                    ar << " style:text-align-source=\"fix\" style:repeat-content=\""
                       << (hAlign == Cell::FILLED ? "true" : "false") << "\"";
                }
                if (vAlign != Cell::VADEFAULT) {
                    const char* sAlign = "bottom";
                    switch (vAlign) {
                        case Cell::TOP:    sAlign = "top"; break;
                        case Cell::MIDDLE: sAlign = "middle"; break;
                        case Cell::BOTTOM: sAlign = "bottom"; break;
                    }
                    ar << " style:vertical-align=\"" << sAlign << "\"";
                }
                ar << "/>\r\n";

                if (hAlign != Cell::HADEFAULT) {
                    const char* sAlign = "end";
                    switch (hAlign) {
                        case Cell::LEFT:      sAlign  = "start";   break;
                        case Cell::CENTER:    sAlign  = "center";  break;
                        case Cell::RIGHT:     sAlign  = "end";     break;
                        case Cell::JUSTIFIED: sAlign  = "justify"; break;
                        case Cell::FILLED:    sAlign  = "start";   break;
                    }
                    ar << "<style:paragraph-properties fo:text-align=\"" << sAlign << "\"/>\r\n";
                }

                ar << "</style:style>\r\n";
            }
        }
    }
}

int OdsWriterImpl::cellStyleIndex(Cell& cell) {
    int i = 0;
    switch (cell.getType()) {
        case Cell::DATE:        i = 1; break;
        case Cell::TIME:        i = 2; break;
    }
    int j = 0;
    switch (cell.getHAlignment()) {
        case Cell::LEFT:        j = 1; break;
        case Cell::CENTER:      j = 2; break;
        case Cell::RIGHT:       j = 3; break;
        case Cell::JUSTIFIED:   j = 4; break;
        case Cell::FILLED:      j = 5; break;
    }
    int k = 0;
    switch (cell.getVAlignment()) {
        case Cell::TOP:         k = 1; break;
        case Cell::MIDDLE:      k = 2; break;
        case Cell::BOTTOM:      k = 3; break;
    }
    return i * 4 * 6 + j * 4 + k;
}

}
