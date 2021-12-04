package net.summer.fastexcel.stream.utils;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.util.IOUtils;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import net.summer.fastexcel.stream.ExcelType;
import net.summer.fastexcel.stream.exceptions.ExcelAnalysisException;
import net.summer.fastexcel.stream.exceptions.ExcelCommonException;

/**
 * @author songyi
 * @Date 2021/11/24
 **/
public class XmlUtils {

    public static ExcelType getType(final InputStream stream) {
        try {
            if (stream instanceof BufferedInputStream) {
                return recognitionExcelType(stream);
            }

            return recognitionExcelType(new BufferedInputStream(stream));
        } catch (ExcelCommonException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelCommonException("Convert excel format exception.You can try specifying the 'excelType' yourself", e);
        }
    }

    public static void parseXmlSource(final InputStream stream, final ContentHandler handler) {
        InputSource inputSource = new InputSource(stream);
        try {
            SAXParserFactory saxFactory = SAXParserFactory.newInstance();

            saxFactory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            saxFactory.setFeature("http://xml.org/sax/features/external-general-entities", false);
            saxFactory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
            SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader xmlReader = saxParser.getXMLReader();
            xmlReader.setContentHandler(handler);
            xmlReader.parse(inputSource);
            stream.close();
        } catch (IOException | ParserConfigurationException | SAXException e) {
            throw new ExcelAnalysisException(e);
        } finally {
            IOUtils.closeQuietly(stream);
        }
    }

    private static ExcelType recognitionExcelType(final InputStream inputStream) throws Exception {
        FileMagic fileMagic = FileMagic.valueOf(inputStream);
        if (FileMagic.OLE2.equals(fileMagic)) {
            return ExcelType.XLS;
        }
        if (FileMagic.OOXML.equals(fileMagic)) {
            return ExcelType.XLSX;
        }
        throw new ExcelCommonException("Convert excel format exception.You can try specifying the 'excelType' yourself");
    }
}
