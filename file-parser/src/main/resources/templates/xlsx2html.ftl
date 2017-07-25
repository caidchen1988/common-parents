<#-- 遍历加载excel内容 -->
<div class="table_container">
    <table id="dataTable">
        <thead>
            <tr>
                <#list titleCellList as cell>
                <td><input type="text" class="table_content" id="title-cellid-${cell.colNum}" value="${cell.colVal}"/></td>
                </#list>
            </tr>
        </thead>
        <tbody>
            <#list rowList as row>
            <tr>
                <#list row.cellList as cell>
                <td><input type="text" class="table_content" id="content-cellid-${cell.colNum}" value="${cell.colVal}"/></td>
                </#list>
            </tr>
            </#list>
        </tbody>
    </table>

    <!-- sheet -->
    <div class="select_container clearfix">
        <#list sheetList as sheet>
        <a href="javascript:;" class="active" id="sheetid-${sheet.sheetNum}">${sheet.name}</a>
        </#list>
    </div>
</div>