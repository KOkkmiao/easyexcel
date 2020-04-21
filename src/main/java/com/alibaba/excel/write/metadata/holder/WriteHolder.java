package com.alibaba.excel.write.metadata.holder;

import java.util.List;
import java.util.Map;

import com.alibaba.excel.metadata.ConfigurationHolder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;

/**
 *
 * Get the corresponding Holder
 *
 * @author Jiaju Zhuang
 **/
public interface WriteHolder extends ConfigurationHolder {
    /**
     * What 'ExcelWriteHeadProperty' does the currently operated cell need to execute
     *
     * @return
     */
    ExcelWriteHeadProperty excelWriteHeadProperty();

    /**
     * What handler does the currently operated cell need to execute
     *
     * @return
     */
    Map<Class<? extends WriteHandler>, List<WriteHandler>> writeHandlerMap();

    /**
     * Whether a header is required for the currently operated cell
     *
     * @return
     */
    boolean needHead();

    /**
     * Writes the head relative to the existing contents of the sheet. Indexes are zero-based.
     *
     * @return
     */
    int relativeHeadRowIndex();
}
