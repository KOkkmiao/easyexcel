package com.alibaba.excel.analysis;

import java.io.InputStream;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.excel.analysis.v03.XlsSaxAnalyser;
import com.alibaba.excel.analysis.v07.XlsxSaxAnalyser;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.context.AnalysisContextImpl;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.exception.ExcelAnalysisStopException;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.read.metadata.ReadWorkbook;
import com.alibaba.excel.read.metadata.holder.ReadWorkbookHolder;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.FileUtils;

/**
 * @author jipengfei
 */
public class ExcelAnalyserImpl implements ExcelAnalyser {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelAnalyserImpl.class);

    private AnalysisContext analysisContext;

    private ExcelExecutor excelExecutor;

    public ExcelAnalyserImpl(ReadWorkbook readWorkbook) {
        try {
            analysisContext = new AnalysisContextImpl(readWorkbook);
            choiceExcelExecutor();
        } catch (RuntimeException e) {
            finish();
            throw e;
        } catch (Throwable e) {
            finish();
            throw new ExcelAnalysisException(e);
        }
    }

    private void choiceExcelExecutor() throws Exception {
        ReadWorkbookHolder readWorkbookHolder = analysisContext.readWorkbookHolder();
        ExcelTypeEnum excelType = readWorkbookHolder.getExcelType();
        if (excelType == null) {
            excelExecutor = new XlsxSaxAnalyser(analysisContext, null);
            return;
        }
        switch (excelType) {
            case XLS:
                POIFSFileSystem poifsFileSystem = null;
                if (readWorkbookHolder.getFile() != null) {
                    poifsFileSystem = new POIFSFileSystem(readWorkbookHolder.getFile());
                } else {
                    poifsFileSystem = new POIFSFileSystem(readWorkbookHolder.getInputStream());
                }
                // So in encrypted excel, it looks like XLS but it's actually XLSX
                if (poifsFileSystem.getRoot().hasEntry(Decryptor.DEFAULT_POIFS_ENTRY)) {
                    InputStream decryptedStream = null;
                    try {
                        decryptedStream =
                            DocumentFactoryHelper.getDecryptedStream(poifsFileSystem.getRoot().getFileSystem(), null);
                        excelExecutor = new XlsxSaxAnalyser(analysisContext, decryptedStream);
                        return;
                    } finally {
                        IOUtils.closeQuietly(decryptedStream);
                        // as we processed the full stream already, we can close the filesystem here
                        // otherwise file handles are leaked
                        poifsFileSystem.close();
                    }
                }
                excelExecutor = new XlsSaxAnalyser(analysisContext, poifsFileSystem);
                break;
            case XLSX:
                excelExecutor = new XlsxSaxAnalyser(analysisContext, null);
                break;
            default:
        }
    }

    @Override
    public void analysis(ReadSheet readSheet) {
        try {
            analysisContext.currentSheet(excelExecutor, readSheet);
            try {
                excelExecutor.execute();
            } catch (ExcelAnalysisStopException e) {
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug("Custom stop!");
                }
            }
            analysisContext.readSheetHolder().notifyAfterAllAnalysed(analysisContext);
        } catch (RuntimeException e) {
            finish();
            throw e;
        } catch (Throwable e) {
            finish();
            throw new ExcelAnalysisException(e);
        }
    }

    @Override
    public void finish() {
        if (analysisContext == null || analysisContext.readWorkbookHolder() == null) {
            return;
        }
        ReadWorkbookHolder readWorkbookHolder = analysisContext.readWorkbookHolder();
        try {
            if (readWorkbookHolder.getReadCache() != null) {
                readWorkbookHolder.getReadCache().destroy();
            }
        } catch (Throwable t) {
            throwCanNotCloseIo(t);
        }
        try {
            if (readWorkbookHolder.getOpcPackage() != null) {
                readWorkbookHolder.getOpcPackage().revert();
            }
        } catch (Throwable t) {
            throwCanNotCloseIo(t);
        }
        try {
            if (readWorkbookHolder.getPoifsFileSystem() != null) {
                readWorkbookHolder.getPoifsFileSystem().close();
            }
        } catch (Throwable t) {
            throwCanNotCloseIo(t);
        }
        try {
            if (analysisContext.readWorkbookHolder().getAutoCloseStream()
                && readWorkbookHolder.getInputStream() != null) {
                readWorkbookHolder.getInputStream().close();
            }
        } catch (Throwable t) {
            throwCanNotCloseIo(t);
        }
        try {
            if (readWorkbookHolder.getTempFile() != null) {
                FileUtils.delete(readWorkbookHolder.getTempFile());
            }
        } catch (Throwable t) {
            throwCanNotCloseIo(t);
        }
    }

    private void throwCanNotCloseIo(Throwable t) {
        throw new ExcelAnalysisException("Can not close IO", t);
    }

    @Override
    public com.alibaba.excel.analysis.ExcelExecutor excelExecutor() {
        return excelExecutor;
    }

    @Override
    public AnalysisContext analysisContext() {
        return analysisContext;
    }
}
