package com.alibaba.excel.analysis.v07;

import java.util.List;

import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import com.alibaba.excel.context.AnalysisContext;

/**
 *
 * @author jipengfei
 */
public class XlsxRowHandler extends DefaultHandler {

    private List<XlsxCellHandler> cellHandlers;
    private XlsxRowResultHolder rowResultHolder;

    public XlsxRowHandler(AnalysisContext analysisContext, StylesTable stylesTable) {
        this.cellHandlers = XlsxHandlerFactory.buildCellHandlers(analysisContext, stylesTable);
        for (XlsxCellHandler cellHandler : cellHandlers) {
            if (cellHandler instanceof XlsxRowResultHolder) {
                this.rowResultHolder = (XlsxRowResultHolder)cellHandler;
                break;
            }
        }
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        for (XlsxCellHandler cellHandler : cellHandlers) {
            if (cellHandler.support(name)) {
                cellHandler.startHandle(name, attributes);
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {
        for (XlsxCellHandler cellHandler : cellHandlers) {
            if (cellHandler.support(name)) {
                cellHandler.endHandle(name);
            }
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        if (rowResultHolder != null) {
            rowResultHolder.appendCurrentCellValue(ch, start, length);
        }
    }
}
