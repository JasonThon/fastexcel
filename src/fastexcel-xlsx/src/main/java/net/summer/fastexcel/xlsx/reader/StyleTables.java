package net.summer.fastexcel.xlsx.reader;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellXfs;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTNumFmt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTNumFmts;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTStylesheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.StyleSheetDocument;

import com.google.common.collect.Maps;

import net.summer.fastexcel.stream.utils.CollectionUtils;

import lombok.Getter;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

/**
 * @author songyi
 * @Date 2021/12/2
 **/
class StyleTables {
    @Getter
    private List<StyleFormat> formats;

    StyleTables(final PackagePart packagePart) throws IOException {
        readFrom(packagePart.getInputStream());
    }

    public CellStyle getIndexAt(final int index) {
        return CollectionUtils.safeGet(formats, index);
    }

    private void readFrom(final InputStream is) throws IOException {
        try {
            Map<Short, String> numberFormat = Maps.newHashMap();
            StyleSheetDocument doc = StyleSheetDocument.Factory.parse(is, DEFAULT_XML_OPTIONS);

            CTStylesheet styleSheet = doc.getStyleSheet();

            CTNumFmts ctfmts = styleSheet.getNumFmts();

            if (ctfmts != null) {
                for (CTNumFmt numFmt : ctfmts.getNumFmtArray()) {
                    numberFormat.put((short) numFmt.getNumFmtId(), numFmt.getFormatCode());
                }
            }

            CTCellXfs cellXfs = styleSheet.getCellXfs();
            if (cellXfs != null) {
                this.formats = CollectionUtils.map(
                        Arrays.asList(cellXfs.getXfArray()),
                        xsf -> new StyleFormat(xsf.getNumFmtId(), numberFormat.get((short) xsf.getNumFmtId()))
                );
            }

            numberFormat.clear();

            doc.set(null);
            doc.setStyleSheet(null);

            is.close();
        } catch (XmlException e) {
            throw new IOException(e.getLocalizedMessage());
        }
    }
}
