// File: XlsxWriterImpl.cpp
// XlsxWriterImpl implementation file
//

#include "splib.h"
#include "XlsxWriterImpl.h"
#include "ZipArchive.h"
#include "ToUTF8.h"
#include "Strings.h"
#include <sstream>
#include "ExcelUtil.h"
#include "Util.h"
#include "splibint.h"

namespace splib {

void XlsxWriterImpl::write(Spreadsheet& sp, const _TCHAR* pathname) {
    ZipArchive ar(pathname);
    writeContentTypes(sp, ar);
    writeRels(ar);
    writeAppDocProps(sp, ar);
    writeCoreDocProps(ar);
    writeStyles(ar);
    writeWorkbookRels(sp, ar);
    writeWorkbook(sp, ar);
    writeTheme(ar);
    for (int i = 0; i < sp.tableCount(); i++) {
        writeSheet(sp.table(i), i + 1, ar);
    }
    ar.close();
}

void XlsxWriterImpl::writeContentTypes(Spreadsheet& sp, ZipArchive& ar) {
    ar.openEntry("[Content_Types].xml");
    ar <<
        "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>\r\n"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n"
        "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\r\n"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\r\n"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\r\n"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\r\n"
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\r\n";
    for (int i = 0; i < sp.tableCount(); i++) {        
        ar <<
            "<Override PartName=\"/xl/worksheets/sheet" << i + 1 << ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\r\n";
    }
    ar <<
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\r\n"
        "</Types>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeRels(ZipArchive& ar) {
    ar.openEntry("_rels/.rels");
    ar <<
        "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>\r\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\r\n"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\r\n"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\r\n"
        "</Relationships>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeAppDocProps(Spreadsheet& sp, ZipArchive& ar) {
    ar.openEntry("docProps/app.xml");
    ar <<
        "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>\r\n"
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\r\n"
        "<Application>Microsoft Excel</Application>\r\n"
        "<DocSecurity>0</DocSecurity>\r\n"
        "<ScaleCrop>false</ScaleCrop>\r\n"
        "<HeadingPairs>\r\n"
        "<vt:vector size=\"2\" baseType=\"variant\">\r\n"
        "<vt:variant>\r\n"
        "<vt:lpstr>Worksheets</vt:lpstr>\r\n"
        "</vt:variant>\r\n"
        "<vt:variant>\r\n"
        "<vt:i4>" << sp.tableCount() << "</vt:i4>\r\n"
        "</vt:variant>\r\n"
        "</vt:vector>\r\n"
        "</HeadingPairs>\r\n"
        "<TitlesOfParts>\r\n"
        "<vt:vector size=\"" << sp.tableCount() << "\" baseType=\"lpstr\">\r\n";
    for (int i = 0; i < sp.tableCount(); i++) {
        ToUTF8 utf8(Strings::xmlize(sp.table(i).getName()).c_str());
        ar << "<vt:lpstr>" << utf8.get() << "</vt:lpstr>\r\n";
    }
    ar <<
        "</vt:vector>\r\n"
        "</TitlesOfParts>\r\n"
        "<Company>N/A</Company>\r\n"
        "<LinksUpToDate>false</LinksUpToDate>\r\n"
        "<SharedDoc>false</SharedDoc>\r\n"
        "<HyperlinksChanged>false</HyperlinksChanged>\r\n"
        "<AppVersion>12.0000</AppVersion>\r\n"
        "</Properties>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeCoreDocProps(ZipArchive& ar) {
    ar.openEntry("docProps/core.xml");
    ar <<
        "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>\r\n"
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\r\n"
        "</cp:coreProperties>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeStyles(ZipArchive& ar) {
    ar.openEntry("xl/styles.xml"); 
    ar <<
        "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>\r\n"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/5/main\">\r\n"
        "<fonts count=\"1\">\r\n"
        "<font>\r\n"
        "<sz val=\"11\"/>\r\n"
        "<color theme=\"1\"/>\r\n"
        "<name val=\"Calibri\"/>\r\n"
        "<family val=\"2\"/>\r\n"
        "<charset val=\"204\"/>\r\n"
        "<scheme val=\"minor\"/>\r\n"
        "</font>\r\n"
        "</fonts>\r\n"
        "<fills count=\"2\">\r\n"
        "<fill>\r\n"
        "<patternFill patternType=\"none\"/>\r\n"
        "</fill>\r\n"
        "<fill>\r\n"
        "<patternFill patternType=\"gray125\"/>\r\n"
        "</fill>\r\n"
        "</fills>\r\n"
        "<borders count=\"1\">\r\n"
        "<border>\r\n"
        "<left/>\r\n"
        "<right/>\r\n"
        "<top/>\r\n"
        "<bottom/>\r\n"
        "<diagonal/>\r\n"
        "</border>\r\n"
        "</borders>\r\n"
        "<cellStyleXfs count=\"1\">\r\n"
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\r\n"
        "</cellStyleXfs>\r\n"
        "<cellXfs>\r\n";
    writeXFs(ar);
    ar <<
        "</cellXfs>\r\n"
        "<cellStyles count=\"1\">\r\n"
        "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\r\n"
        "</cellStyles>\r\n"
        "<dxfs count=\"0\"/>\r\n"
        "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/>\r\n"
        "<colors/>\r\n"
        "</styleSheet>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeXFs(ZipArchive& ar) {
    const int NUM_FMT_GENERAL = 0;
    const int NUM_FMT_DATE = 14;
    const int NUM_FMT_TIME = 21;
    int numFmts[] = {NUM_FMT_GENERAL, NUM_FMT_DATE, NUM_FMT_TIME};
    Cell::HAlignment hAligns[] = {Cell::HADEFAULT, Cell::LEFT, Cell::CENTER,
        Cell::RIGHT, Cell::JUSTIFIED, Cell::FILLED};
    Cell::VAlignment vAligns[] = {Cell::BOTTOM, Cell::TOP, Cell::MIDDLE};
    for (int i = 0; i < 3; i++) {
        for (int j = 0; j < 6; j++) {
            for (int k = 0; k < 3; k++) {
                int numFmt = numFmts[i];
                Cell::HAlignment hAlign = hAligns[j];
                Cell::VAlignment vAlign = vAligns[k];
                const char* applyNumFmt = "";
                const char* applyAlignment = "";
                if (numFmt != NUM_FMT_GENERAL) {
                    applyNumFmt = " applyNumberFormat=\"1\"";
                }                
                if (hAlign != Cell::HADEFAULT || vAlign != Cell::BOTTOM) {
                    applyAlignment = " applyAlignment=\"1\"";
                }
                ar << "<xf numFmtId=\"" << numFmt
                   << "\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\""
                   << applyNumFmt << applyAlignment << ">\r\n";
                if (hAlign != Cell::HADEFAULT || vAlign != Cell::BOTTOM) {
                    const char* horz = "";
                    const char* vert = "";
                    switch (hAlign) {
                        case Cell::LEFT:
                            horz = " horizontal=\"left\"";
                            break;
                        case Cell::CENTER:
                            horz = " horizontal=\"center\"";
                            break;
                        case Cell::RIGHT:
                            horz = " horizontal=\"right\"";
                            break;
                        case Cell::JUSTIFIED:
                            horz = " horizontal=\"justify\"";
                            break;
                        case Cell::FILLED:
                            horz = " horizontal=\"fill\"";
                            break;
                    }
                    switch (vAlign) {
                        case Cell::TOP:
                            vert = " vertical=\"top\"";
                            break;
                        case Cell::MIDDLE:
                            vert = " vertical=\"center\"";
                            break;
                    }
                    ar << "<alignment" << horz << vert << "/>\r\n";
                }
                ar << "</xf>\r\n";
            }
        }
    }
}

int XlsxWriterImpl::xfIndex(Cell& cell) {
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
    }
    return i * 3 * 6 + j * 3 + k;
}

void XlsxWriterImpl::writeWorkbookRels(Spreadsheet& sp, ZipArchive& ar) {
    ar.openEntry("xl/_rels/workbook.xml.rels");
    ar <<
        "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>\r\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n";
    int i;
    for (i = 0; i < sp.tableCount(); i++) {
        ar <<
            "<Relationship Id=\"rId" << i + 1 << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
            "Target=\"worksheets/sheet" << i + 1 << ".xml\"/>\r\n";
    }
    ar <<
        "<Relationship Id=\"rId" << i + 1 << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>\r\n"
        "<Relationship Id=\"rId" << i + 2 << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\r\n"
        "</Relationships>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeWorkbook(Spreadsheet& sp, ZipArchive& ar) {
    ar.openEntry("xl/workbook.xml"); 
    ar <<
        "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>\r\n"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/5/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\r\n"
        "<fileVersion lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4017\"/>\r\n"
        "<workbookPr defaultThemeVersion=\"123820\"/>\r\n"
        "<bookViews>\r\n"
        "<workbookView xWindow=\"120\" yWindow=\"45\" windowWidth=\"18975\" windowHeight=\"11955\"/>\r\n"
        "</bookViews>\r\n"
        "<sheets>\r\n";
    for (int i = 0; i < sp.tableCount(); i++) {
        ToUTF8 utf8(Strings::xmlize(sp.table(i).getName()).c_str());
        ar << "<sheet name=\"" << utf8.get() << "\" sheetId=\"" << i + 1
           << "\" r:id=\"rId" << i + 1 << "\"/>\r\n";
    }
    ar <<
        "</sheets>\r\n"
        "<calcPr calcId=\"122211\"/>\r\n"
        "</workbook>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeTheme(ZipArchive& ar) {
    ar.openEntry("xl/theme/theme1.xml");
    ar <<
        "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"\?>\r\n"
        "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/3/main\" name=\"Office Theme\">\r\n"
        "<a:themeElements>\r\n"
        "<a:clrScheme name=\"Office\">\r\n"
        "<a:dk1>\r\n"
        "<a:sysClr val=\"windowText\"/>\r\n"
        "</a:dk1>\r\n"
        "<a:lt1>\r\n"
        "<a:sysClr val=\"window\"/>\r\n"
        "</a:lt1>\r\n"
        "<a:dk2>\r\n"
        "<a:srgbClr val=\"1F497D\"/>\r\n"
        "</a:dk2>\r\n"
        "<a:lt2>\r\n"
        "<a:srgbClr val=\"EEECE1\"/>\r\n"
        "</a:lt2>\r\n"
        "<a:accent1>\r\n"
        "<a:srgbClr val=\"4F81BD\"/>\r\n"
        "</a:accent1>\r\n"
        "<a:accent2>\r\n"
        "<a:srgbClr val=\"C0504D\"/>\r\n"
        "</a:accent2>\r\n"
        "<a:accent3>\r\n"
        "<a:srgbClr val=\"9BBB59\"/>\r\n"
        "</a:accent3>\r\n"
        "<a:accent4>\r\n"
        "<a:srgbClr val=\"8064A2\"/>\r\n"
        "</a:accent4>\r\n"
        "<a:accent5>\r\n"
        "<a:srgbClr val=\"4BACC6\"/>\r\n"
        "</a:accent5>\r\n"
        "<a:accent6>\r\n"
        "<a:srgbClr val=\"F79646\"/>\r\n"
        "</a:accent6>\r\n"
        "<a:hlink>\r\n"
        "<a:srgbClr val=\"0000FF\"/>\r\n"
        "</a:hlink>\r\n"
        "<a:folHlink>\r\n"
        "<a:srgbClr val=\"800080\"/>\r\n"
        "</a:folHlink>\r\n"
        "</a:clrScheme>\r\n"
        "<a:fontScheme name=\"Office\">\r\n"
        "<a:majorFont>\r\n"
        "<a:latin typeface=\"Cambria\"/>\r\n"
        "<a:ea typeface=\"\"/>\r\n"
        "<a:cs typeface=\"\"/>\r\n"
        "<a:font script=\"Jpan\" typeface=\"\xEF\xBC\xAD\xEF\xBC\xB3\x20\xEF\xBC\xB0\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF\"/>\r\n"
        "<a:font script=\"Hang\" typeface=\"\xEB\xA7\x91\xEC\x9D\x80\x20\xEA\xB3\xA0\xEB\x94\x95\"/>\r\n"
        "<a:font script=\"Hans\" typeface=\"\xE5\xAE\x8B\xE4\xBD\x93\"/>\r\n"
        "<a:font script=\"Hant\" typeface=\"\xE6\x96\xB0\xE7\xB4\xB0\xE6\x98\x8E\xE9\xAB\x94\"/>\r\n"
        "<a:font script=\"Arab\" typeface=\"Times New Roman\"/>\r\n"
        "<a:font script=\"Hebr\" typeface=\"Times New Roman\"/>\r\n"
        "<a:font script=\"Thai\" typeface=\"Angsana New\"/>\r\n"
        "<a:font script=\"Ethi\" typeface=\"Nyala\"/>\r\n"
        "<a:font script=\"Beng\" typeface=\"Vrinda\"/>\r\n"
        "<a:font script=\"Gujr\" typeface=\"Shruti\"/>\r\n"
        "<a:font script=\"Khmr\" typeface=\"MoolBoran\"/>\r\n"
        "<a:font script=\"Knda\" typeface=\"Tunga\"/>\r\n"
        "<a:font script=\"Guru\" typeface=\"Raavi\"/>\r\n"
        "<a:font script=\"Cans\" typeface=\"Euphemia\"/>\r\n"
        "<a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>\r\n"
        "<a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>\r\n"
        "<a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>\r\n"
        "<a:font script=\"Thaa\" typeface=\"MV Boli\"/>\r\n"
        "<a:font script=\"Deva\" typeface=\"Mangal\"/>\r\n"
        "<a:font script=\"Telu\" typeface=\"Gautami\"/>\r\n"
        "<a:font script=\"Taml\" typeface=\"Latha\"/>\r\n"
        "<a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>\r\n"
        "<a:font script=\"Orya\" typeface=\"Kalinga\"/>\r\n"
        "<a:font script=\"Mlym\" typeface=\"Kartika\"/>\r\n"
        "<a:font script=\"Laoo\" typeface=\"DokChampa\"/>\r\n"
        "<a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>\r\n"
        "<a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>\r\n"
        "<a:font script=\"Viet\" typeface=\"Times New Roman\"/>\r\n"
        "<a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>\r\n"
        "</a:majorFont>\r\n"
        "<a:minorFont>\r\n"
        "<a:latin typeface=\"Calibri\"/>\r\n"
        "<a:ea typeface=\"\"/>\r\n"
        "<a:cs typeface=\"\"/>\r\n"
        "<a:font script=\"Jpan\" typeface=\"\xEF\xBC\xAD\xEF\xBC\xB3\x20\xEF\xBC\xB0\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF\"/>\r\n"
        "<a:font script=\"Hang\" typeface=\"\xEB\xA7\x91\xEC\x9D\x80\x20\xEA\xB3\xA0\xEB\x94\x95\"/>\r\n"
        "<a:font script=\"Hans\" typeface=\"\xE5\xAE\x8B\xE4\xBD\x93\"/>\r\n"
        "<a:font script=\"Hant\" typeface=\"\xE6\x96\xB0\xE7\xB4\xB0\xE6\x98\x8E\xE9\xAB\x94\"/>\r\n"
        "<a:font script=\"Arab\" typeface=\"Arial\"/>\r\n"
        "<a:font script=\"Hebr\" typeface=\"Arial\"/>\r\n"
        "<a:font script=\"Thai\" typeface=\"Cordia New\"/>\r\n"
        "<a:font script=\"Ethi\" typeface=\"Nyala\"/>\r\n"
        "<a:font script=\"Beng\" typeface=\"Vrinda\"/>\r\n"
        "<a:font script=\"Gujr\" typeface=\"Shruti\"/>\r\n"
        "<a:font script=\"Khmr\" typeface=\"DaunPenh\"/>\r\n"
        "<a:font script=\"Knda\" typeface=\"Tunga\"/>\r\n"
        "<a:font script=\"Guru\" typeface=\"Raavi\"/>\r\n"
        "<a:font script=\"Cans\" typeface=\"Euphemia\"/>\r\n"
        "<a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/>\r\n"
        "<a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/>\r\n"
        "<a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/>\r\n"
        "<a:font script=\"Thaa\" typeface=\"MV Boli\"/>\r\n"
        "<a:font script=\"Deva\" typeface=\"Mangal\"/>\r\n"
        "<a:font script=\"Telu\" typeface=\"Gautami\"/>\r\n"
        "<a:font script=\"Taml\" typeface=\"Latha\"/>\r\n"
        "<a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/>\r\n"
        "<a:font script=\"Orya\" typeface=\"Kalinga\"/>\r\n"
        "<a:font script=\"Mlym\" typeface=\"Kartika\"/>\r\n"
        "<a:font script=\"Laoo\" typeface=\"DokChampa\"/>\r\n"
        "<a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/>\r\n"
        "<a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/>\r\n"
        "<a:font script=\"Viet\" typeface=\"Arial\"/>\r\n"
        "<a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/>\r\n"
        "</a:minorFont>\r\n"
        "</a:fontScheme>\r\n"
        "<a:fmtScheme name=\"Office\">\r\n"
        "<a:fillStyleLst>\r\n"
        "<a:solidFill>\r\n"
        "<a:schemeClr val=\"phClr\"/>\r\n"
        "</a:solidFill>\r\n"
        "<a:gradFill rotWithShape=\"1\">\r\n"
        "<a:gsLst>\r\n"
        "<a:gs pos=\"0\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:tint val=\"50000\"/>\r\n"
        "<a:satMod val=\"300000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "<a:gs pos=\"35000\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:tint val=\"37000\"/>\r\n"
        "<a:satMod val=\"300000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "<a:gs pos=\"100000\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:tint val=\"15000\"/>\r\n"
        "<a:satMod val=\"350000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "</a:gsLst>\r\n"
        "<a:lin ang=\"16200000\" scaled=\"1\"/>\r\n"
        "</a:gradFill>\r\n"
        "<a:gradFill rotWithShape=\"1\">\r\n"
        "<a:gsLst>\r\n"
        "<a:gs pos=\"0\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:shade val=\"40000\"/>\r\n"
        "<a:satMod val=\"155000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "<a:gs pos=\"65000\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:shade val=\"85000\"/>\r\n"
        "<a:satMod val=\"155000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "<a:gs pos=\"100000\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:shade val=\"95000\"/>\r\n"
        "<a:satMod val=\"155000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "</a:gsLst>\r\n"
        "<a:lin ang=\"16200000\" scaled=\"0\"/>\r\n"
        "</a:gradFill>\r\n"
        "</a:fillStyleLst>\r\n"
        "<a:lnStyleLst>\r\n"
        "<a:ln w=\"6350\" cap=\"rnd\" cmpd=\"sng\" algn=\"ctr\">\r\n"
        "<a:solidFill>\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:shade val=\"95000\"/>\r\n"
        "<a:satMod val=\"105000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:solidFill>\r\n"
        "<a:prstDash val=\"solid\"/>\r\n"
        "</a:ln>\r\n"
        "<a:ln w=\"25400\" cap=\"rnd\" cmpd=\"sng\" algn=\"ctr\">\r\n"
        "<a:solidFill>\r\n"
        "<a:schemeClr val=\"phClr\"/>\r\n"
        "</a:solidFill>\r\n"
        "<a:prstDash val=\"solid\"/>\r\n"
        "</a:ln>\r\n"
        "<a:ln w=\"34925\" cap=\"rnd\" cmpd=\"sng\" algn=\"ctr\">\r\n"
        "<a:solidFill>\r\n"
        "<a:schemeClr val=\"phClr\"/>\r\n"
        "</a:solidFill>\r\n"
        "<a:prstDash val=\"solid\"/>\r\n"
        "</a:ln>\r\n"
        "</a:lnStyleLst>\r\n"
        "<a:effectStyleLst>\r\n"
        "<a:effectStyle>\r\n"
        "<a:effectLst>\r\n"
        "<a:outerShdw blurRad=\"50800\" algn=\"tl\" rotWithShape=\"0\">\r\n"
        "<a:srgbClr val=\"000000\">\r\n"
        "<a:alpha val=\"64000\"/>\r\n"
        "</a:srgbClr>\r\n"
        "</a:outerShdw>\r\n"
        "</a:effectLst>\r\n"
        "</a:effectStyle>\r\n"
        "<a:effectStyle>\r\n"
        "<a:effectLst>\r\n"
        "<a:outerShdw blurRad=\"39000\" dist=\"25400\" dir=\"5400000\">\r\n"
        "<a:srgbClr val=\"000000\">\r\n"
        "<a:alpha val=\"35000\"/>\r\n"
        "</a:srgbClr>\r\n"
        "</a:outerShdw>\r\n"
        "</a:effectLst>\r\n"
        "</a:effectStyle>\r\n"
        "<a:effectStyle>\r\n"
        "<a:effectLst>\r\n"
        "<a:outerShdw blurRad=\"39000\" dist=\"25400\" dir=\"5400000\">\r\n"
        "<a:srgbClr val=\"000000\">\r\n"
        "<a:alpha val=\"35000\"/>\r\n"
        "</a:srgbClr>\r\n"
        "</a:outerShdw>\r\n"
        "</a:effectLst>\r\n"
        "<a:scene3d>\r\n"
        "<a:camera prst=\"orthographicFront\" fov=\"0\">\r\n"
        "<a:rot lat=\"0\" lon=\"0\" rev=\"0\"/>\r\n"
        "</a:camera>\r\n"
        "<a:lightRig rig=\"threePt\" dir=\"t\">\r\n"
        "<a:rot lat=\"0\" lon=\"0\" rev=\"0\"/>\r\n"
        "</a:lightRig>\r\n"
        "</a:scene3d>\r\n"
        "<a:sp3d prstMaterial=\"matte\">\r\n"
        "<a:bevelT h=\"22225\"/>\r\n"
        "</a:sp3d>\r\n"
        "</a:effectStyle>\r\n"
        "</a:effectStyleLst>\r\n"
        "<a:bgFillStyleLst>\r\n"
        "<a:solidFill>\r\n"
        "<a:schemeClr val=\"phClr\"/>\r\n"
        "</a:solidFill>\r\n"
        "<a:gradFill rotWithShape=\"1\">\r\n"
        "<a:gsLst>\r\n"
        "<a:gs pos=\"0\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:shade val=\"50000\"/>\r\n"
        "<a:satMod val=\"155000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "<a:gs pos=\"35000\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:shade val=\"75000\"/>\r\n"
        "<a:satMod val=\"155000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "<a:gs pos=\"100000\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:tint val=\"80000\"/>\r\n"
        "<a:satMod val=\"255000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "</a:gsLst>\r\n"
        "<a:lin ang=\"16200000\" scaled=\"0\"/>\r\n"
        "</a:gradFill>\r\n"
        "<a:gradFill rotWithShape=\"1\">\r\n"
        "<a:gsLst>\r\n"
        "<a:gs pos=\"0\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:tint val=\"80000\"/>\r\n"
        "<a:satMod val=\"300000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "<a:gs pos=\"100000\">\r\n"
        "<a:schemeClr val=\"phClr\">\r\n"
        "<a:shade val=\"30000\"/>\r\n"
        "<a:satMod val=\"200000\"/>\r\n"
        "</a:schemeClr>\r\n"
        "</a:gs>\r\n"
        "</a:gsLst>\r\n"
        "<a:path path=\"circle\">\r\n"
        "<a:fillToRect l=\"100000\" t=\"100000\" r=\"100000\" b=\"100000\"/>\r\n"
        "</a:path>\r\n"
        "</a:gradFill>\r\n"
        "</a:bgFillStyleLst>\r\n"
        "</a:fmtScheme>\r\n"
        "</a:themeElements>\r\n"
        "<a:objectDefaults/>\r\n"
        "<a:extraClrSchemeLst/>\r\n"
        "</a:theme>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeSheet(Table& table, int id, ZipArchive& ar) {
    std::basic_stringstream<char> entryName;
    entryName << "xl/worksheets/sheet" << id << ".xml";
    ar.openEntry(entryName.str().c_str());
    ar << "<\?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"\?>\r\n"
          "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/5/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\r\n";        

    // columns
    if (table.columns().size() > 0) {
        ar << "<cols>\r\n";
        Columns::Iterator* colIt = table.columns().iterator();
        while (colIt->hasNext()) {
            Columns::Entry entry = colIt->next();
            int col = entry.index();
            double width = entry.object().getWidth();
            if (width >= 0) {
                double w = ExcelUtil::columnWidthUnits(width);
                ar << "<col min=\"" << col + 1 << "\" max=\"" << col + 1 << "\" width=\"" << w << "\" customWidth=\"1\"/>\r\n";
            }
        }
        delete colIt;
        ar << "</cols>\r\n";
    }

    ar << "<sheetData>\r\n";

    // rows
    Rows::Iterator* rowIt = table.rows().iterator();
    while (rowIt->hasNext()) {
        Rows::Entry entry = rowIt->next();
        int row = entry.index();
        ar << "<row r=\"" << row + 1 << "\"";
        double height = entry.object().getHeight();
        if (height >= 0) {
            ar << " ht=\"" << height << "\" customHeight=\"1\"";
        }            
        ar << ">\r\n";
        Cells::Iterator* cellIt = entry.object().cells().iterator();
        while (cellIt->hasNext()) {
            Cells::Entry entry = cellIt->next();
            int col = entry.index();
            Cell& cell = entry.object();
            if (cell.getType() == Cell::NONE) {
                continue;
            }
            writeCell(cell, col, row, ar);
        }
        delete cellIt;            
        ar << "</row>\r\n";
    }
    delete rowIt;

    ar << "</sheetData>\r\n"
          "</worksheet>\r\n";
    ar.closeEntry();
}

void XlsxWriterImpl::writeCell(Cell& cell, int col, int row, ZipArchive& ar) {
    std::basic_string<char> loc = Util::buildLocation(col, row);
    ar << "<c r=\"" << loc.c_str() << "\"";
    int s = xfIndex(cell);
    if (s != 0) {
        ar << " s=\"" << s << "\"";
    }
    ar << ">\r\n";
    if (cell.getType() == Cell::TEXT) {
        ToUTF8 utf8(Strings::xmlize(cell.getText()).c_str());
        ar << "<is><t>" << utf8.get() << "</t></is>\r\n";
    } else if (cell.getType() == Cell::LONG) {
        ar << "<v>" << cell.getLong() << "</v>\r\n";
    } else if (cell.getType() == Cell::DOUBLE) {
        ar << "<v>" << cell.getDouble() << "</v>\r\n";
    } else if (cell.getType() == Cell::DATE) {
        ar << "<v>" << ExcelUtil::date(cell.getDate()) << "</v>\r\n";
    } else if (cell.getType() == Cell::TIME) {
        ar << "<v>" << ExcelUtil::time(cell.getTime()) << "</v>\r\n";
    } else if (cell.getType() == Cell::FORMULA) {
        ToUTF8 utf8(cell.getFormula());
        ar << "<f>" << utf8.get() << "</f>\r\n";
    }
    ar << "</c>\r\n";
}

}
