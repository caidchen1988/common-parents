<#-- 遍历加载excel内容 -->
<div class="table_container">
    <#if titleCellList??>
        <table id="dataTable">
            <thead>
                <tr>
                    <#list titleCellList as cell>
                    <td class="table_content" id="title-cellid-${cell.colNum}">
                        ${cell.colVal}
                    </td>
                    </#list>
                </tr>
            </thead>
            <tbody>
                <#list rowList as row>
                <tr>
                    <#list row.cellList as cell>
                    <td class="table_content" id="content-cellid-${cell.colNum}">
                        ${cell.colVal}
                    </td>
                    </#list>
                </tr>
                </#list>
            </tbody>
        </table>

        <!-- sheet -->
        <div class="select_container clearfix">
            <#list sheetList as sheet>
            <a href="javascript:;" id="${sheet.sheetNum}">${sheet.name}</a>
            </#list>
        </div>
    <#else>
        数据异常，请联系客服解决。
    </#if>
</div>