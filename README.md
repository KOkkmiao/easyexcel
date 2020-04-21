easyexcel
======================
#### 功能: 添加了表格绘制之后对内容的二次合并的绘制操作
### 使用方式
    /**
     * @author 35716 苗小鹏 <a href="xiaopeng.miao@1hai.cn">Contact me.</a>
     * @version 1.0
     * @since 2020/04/10 16:46
     * desc : 建议 beforeCellCreate 进行列合并  afterCellCreate 进行 行合并
     */
    public class TenancyManageMentMergeHandler implements CellWriteHandler {
        @Override
        public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, int relativeRowIndex, boolean isHead) {
            Sheet sheet = writeSheetHolder.getSheet();
            Cell cell;
            String value = "";
            int rowNmber = sheet.getLastRowNum();
            for (int i = sheet.getLastRowNum(); i >0; i--) {
                cell = sheet.getRow(i).getCell(0);
                if(null == cell){
                    value = "";
                    rowNmber=i;
                    continue;
                }
                String cellValue = cell.getStringCellValue();
                if(value.equals("")){
                    value = cellValue;
                    rowNmber=i;
                }else if(!value.equals(cellValue)){
                    //不相等 需要进行合并判断
                    if(rowNmber-i!=1){
                        sheet.addMergedRegion(new CellRangeAddress(i+1,rowNmber,0,0));
                    }
                    //合并完之后一定要切换行号和单元格内容
                    rowNmber=i;
                    value=cellValue;
                }
    
            }
        }
    
    使用方式：基本上使用方式很简单 就是创建一个 CellWriteHandler 的书写方式 假如到 registerWriteEndHandler() 方法中
        EasyExcel.write(res.getOutputStream(), TenancyManagementInfo.class)
                .registerWriteHandler(commonEasyExcelStyle).excelType(ExcelTypeEnum.XLS)
                .registerWriteEndHandler(new TenancyManageMentMergeHandler())
                .sheet(false?"":"租期管理").doWrite(tenancyManagementInfos);
#### 如果有新的想法可留言或微信联系搜索json-pro001