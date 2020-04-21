package com.alibaba.excel.write.metadata;

import java.util.ArrayList;
import java.util.List;

import com.alibaba.excel.metadata.BasicParameter;
import com.alibaba.excel.write.handler.WriteHandler;

/**
 * Write basic parameter
 *
 * @author Jiaju Zhuang
 **/
public class WriteBasicParameter extends BasicParameter {
    /**
     * Writes the head relative to the existing contents of the sheet. Indexes are zero-based.
     */
    private Integer relativeHeadRowIndex;
    /**
     * Need Head
     */
    private Boolean needHead;
    /**
     * Custom type handler override the default
     */
    private List<WriteHandler> customWriteHandlerList = new ArrayList<WriteHandler>();

    /**
     * 定义表格结束后的cell 的操作行为
     * @return
     */
    private WriteHandler endHandler = null;
    public Integer getRelativeHeadRowIndex() {
        return relativeHeadRowIndex;
    }

    public void setRelativeHeadRowIndex(Integer relativeHeadRowIndex) {
        this.relativeHeadRowIndex = relativeHeadRowIndex;
    }

    public Boolean getNeedHead() {
        return needHead;
    }

    public void setNeedHead(Boolean needHead) {
        this.needHead = needHead;
    }

    public List<WriteHandler> getCustomWriteHandlerList() {
        return customWriteHandlerList;
    }

    public void setCustomWriteHandlerList(List<WriteHandler> customWriteHandlerList) {
        this.customWriteHandlerList = customWriteHandlerList;
    }

    public WriteHandler getEndHandler() {
        return endHandler;
    }

    public void setEndHandler(WriteHandler endHandler) {
        this.endHandler = endHandler;
    }
}
