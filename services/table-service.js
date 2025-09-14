// services/table-service.js
export class TableService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 在光标位置插入表格
    async insertTable(opts) {
        const options = opts || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            rows: options.rows || 3,
            columns: options.columns || 3,
            widthType: options.widthType || 'percent',
            width: options.width || 100,
            headers: options.headers || [],
            data: options.data || [],
            style: options.style || 'default'
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            var rows = (Asc.scope && Asc.scope.rows) || 3;
            var columns = (Asc.scope && Asc.scope.columns) || 3;
            var widthType = (Asc.scope && Asc.scope.widthType) || 'percent';
            var width = (Asc.scope && Asc.scope.width) || 100;
            var headers = (Asc.scope && Asc.scope.headers) || [];
            var data = (Asc.scope && Asc.scope.data) || [];
            var style = (Asc.scope && Asc.scope.style) || 'default';

            console.log('=== 表格插入 ===');
            console.log('参数:', { rows, columns, widthType, width, headers, data, style });

            try {
                // 方法1: 尝试使用InsertContent插入表格
                console.log('尝试方法1: InsertContent插入表格...');
                var table = Api.CreateTable(columns, rows);

                if (!table) {
                    console.log('创建表格失败');
                    return { success: false, error: 'Failed to create table' };
                }

                console.log('表格创建成功:', table);

                // 设置表格宽度
                if (typeof table.SetWidth === 'function') {
                    table.SetWidth(widthType, width);
                    console.log('设置表格宽度:', widthType, width);
                }

                // 填充表头
                if (headers.length > 0) {
                    console.log('填充表头数据...');
                    for (var j = 0; j < Math.min(headers.length, columns); j++) {
                        if (typeof table.GetRow === 'function') {
                            var row = table.GetRow(0);
                            if (row && typeof row.GetCell === 'function') {
                                var cell = row.GetCell(j);
                                if (cell && typeof cell.GetContent === 'function') {
                                    var cellContent = cell.GetContent();
                                    if (cellContent && typeof cellContent.GetElement === 'function') {
                                        var para = cellContent.GetElement(0);
                                        if (para && typeof para.AddText === 'function') {
                                            para.AddText(headers[j]);
                                            console.log(`表头[${j}]: ${headers[j]}`);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 填充数据
                if (data.length > 0) {
                    console.log('填充表格数据...');
                    var startRow = headers.length > 0 ? 1 : 0;
                    for (var i = 0; i < Math.min(data.length, rows - startRow); i++) {
                        var rowData = data[i];
                        if (Array.isArray(rowData)) {
                            for (var j = 0; j < Math.min(rowData.length, columns); j++) {
                                if (typeof table.GetRow === 'function') {
                                    var row = table.GetRow(startRow + i);
                                    if (row && typeof row.GetCell === 'function') {
                                        var cell = row.GetCell(j);
                                        if (cell && typeof cell.GetContent === 'function') {
                                            var cellContent = cell.GetContent();
                                            if (cellContent && typeof cellContent.GetElement === 'function') {
                                                var para = cellContent.GetElement(0);
                                                if (para && typeof para.AddText === 'function') {
                                                    para.AddText(String(rowData[j]));
                                                    console.log(`数据[${i}][${j}]: ${rowData[j]}`);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 尝试使用InsertContent插入表格
                if (typeof doc.InsertContent === 'function') {
                    console.log('尝试使用InsertContent插入表格...');
                    var insertResult = doc.InsertContent([table], false);
                    console.log('InsertContent结果:', insertResult);

                    if (insertResult) {
                        console.log('✅ 表格通过InsertContent插入成功');
                        return {
                            success: true,
                            message: 'Table inserted using InsertContent',
                            method: 'InsertContent',
                            parameters: { rows, columns, width: `${width}${widthType}` }
                        };
                    }
                }

                // 方法2: 获取当前段落并插入表格
                console.log('方法2: 尝试段落级别插入...');
                var range = doc.GetRangeBySelect();
                var targetParagraph = null;

                if (range && typeof range.GetParagraph === 'function') {
                    targetParagraph = range.GetParagraph();
                }

                if (!targetParagraph && typeof doc.GetElement === 'function') {
                    targetParagraph = doc.GetElement(0);
                    console.log('使用第一个段落作为插入点');
                }

                if (targetParagraph) {
                    // 尝试在段落后添加表格
                    console.log('尝试在段落后添加表格...');

                    // 获取段落的父容器
                    if (typeof targetParagraph.GetParent === 'function') {
                        var parentContent = targetParagraph.GetParent();
                        if (parentContent && typeof parentContent.AddElement === 'function') {
                            // 找到当前段落的位置
                            var elementCount = 0;
                            var targetIndex = -1;
                            for (var i = 0; i < 50; i++) {
                                var element = doc.GetElement(i);
                                if (element) {
                                    if (element === targetParagraph) {
                                        targetIndex = i + 1;
                                        break;
                                    }
                                    elementCount++;
                                } else {
                                    break;
                                }
                            }

                            if (targetIndex > 0) {
                                console.log(`在位置 ${targetIndex} 插入表格`);
                                // 使用Document的Push或类似方法
                                if (typeof doc.Push === 'function') {
                                    doc.Push(table);
                                    console.log('✅ 表格通过Document.Push插入成功');
                                    return {
                                        success: true,
                                        message: 'Table inserted using Document.Push',
                                        method: 'Document-Push',
                                        parameters: { rows, columns, width: `${width}${widthType}` }
                                    };
                                }
                            }
                        }
                    }
                }

                // 方法3: 作为文档元素直接添加
                console.log('方法3: 尝试作为文档元素添加...');
                if (typeof doc.Push === 'function') {
                    doc.Push(table);
                    console.log('✅ 表格通过Document.Push添加成功');
                    return {
                        success: true,
                        message: 'Table added to document end',
                        method: 'Document-Push-End',
                        parameters: { rows, columns, width: `${width}${widthType}` }
                    };
                }

                return { success: false, error: 'All table insertion methods failed' };

            } catch (e) {
                console.log('表格插入失败:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }

    // 插入动态数据表格（从宿主系统绑定）
    async insertDynamicTable(dynamicData) {
        if (!dynamicData) {
            return { success: false, error: 'No dynamic data provided' };
        }

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            title: dynamicData.title || '数据表格',
            headers: dynamicData.headers || [],
            data: dynamicData.data || [],
            metadata: dynamicData.metadata || {},
            styling: dynamicData.styling || {},
            rows: (dynamicData.data || []).length + (dynamicData.headers ? 1 : 0),
            columns: Math.max(
                (dynamicData.headers || []).length,
                Math.max(...(dynamicData.data || []).map(row => Array.isArray(row) ? row.length : 0), 0)
            )
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();
            var scope = Asc.scope;

            console.log('=== 动态数据表格插入 ===');
            console.log('标题:', scope.title);
            console.log('元数据:', scope.metadata);
            console.log('样式:', scope.styling);
            console.log('数据结构:', { rows: scope.rows, columns: scope.columns });

            try {
                // 插入标题（如果有）
                var titlePara = null;
                if (scope.title) {
                    titlePara = Api.CreateParagraph();
                    titlePara.SetSpacingAfter(200);
                    titlePara.SetJc('center');

                    var titleRun = Api.CreateRun();
                    titleRun.SetBold(true);
                    titleRun.SetFontSize(14);
                    titleRun.AddText(scope.title);
                    titlePara.AddElement(titleRun);

                    // 添加元数据信息
                    if (scope.metadata.generatedAt) {
                        titleRun.AddLineBreak();
                        var metaRun = Api.CreateRun();
                        metaRun.SetFontSize(10);
                        metaRun.SetColor(128, 128, 128);
                        metaRun.AddText('生成时间: ' + scope.metadata.generatedAt);
                        titlePara.AddElement(metaRun);
                    }

                    if (scope.metadata.dataSource) {
                        var sourceRun = Api.CreateRun();
                        sourceRun.SetFontSize(10);
                        sourceRun.SetColor(128, 128, 128);
                        sourceRun.AddText(' | 数据源: ' + scope.metadata.dataSource);
                        titlePara.AddElement(sourceRun);
                    }
                }

                // 创建表格
                var table = Api.CreateTable(scope.columns, scope.rows);
                if (!table) {
                    console.log('创建动态表格失败');
                    return { success: false, error: 'Failed to create dynamic table' };
                }

                console.log('动态表格创建成功');

                // 设置表格样式
                if (typeof table.SetWidth === 'function') {
                    table.SetWidth('percent', 100);
                }

                // 设置表格边框
                if (typeof table.SetTableBorderTop === 'function') {
                    table.SetTableBorderTop('single', 4, 0, 0, 0, 0);
                    table.SetTableBorderBottom('single', 4, 0, 0, 0, 0);
                    table.SetTableBorderLeft('single', 4, 0, 0, 0, 0);
                    table.SetTableBorderRight('single', 4, 0, 0, 0, 0);
                    table.SetTableBorderInsideH('single', 2, 0, 0, 0, 0);
                    table.SetTableBorderInsideV('single', 2, 0, 0, 0, 0);
                }

                var currentRow = 0;

                // 填充表头
                if (scope.headers.length > 0) {
                    console.log('填充动态表头...');
                    var headerRow = table.GetRow(0);
                    if (headerRow) {
                        for (var j = 0; j < Math.min(scope.headers.length, scope.columns); j++) {
                            var cell = headerRow.GetCell(j);
                            if (cell) {
                                var cellContent = cell.GetContent();
                                if (cellContent) {
                                    var para = cellContent.GetElement(0);
                                    if (para) {
                                        para.AddText(String(scope.headers[j]));

                                        // 表头样式
                                        if (scope.styling.headerStyle === 'bold') {
                                            var textPr = Api.CreateTextPr();
                                            textPr.SetBold(true);
                                            para.SetTextPr(textPr);
                                        }

                                        // 表头居中
                                        para.SetJc('center');

                                        // 表头背景色
                                        if (typeof cell.SetShd === 'function') {
                                            cell.SetShd('clear', 220, 220, 220);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    currentRow = 1;
                }

                // 填充数据
                if (scope.data.length > 0) {
                    console.log('填充动态数据...');
                    for (var i = 0; i < Math.min(scope.data.length, scope.rows - currentRow); i++) {
                        var rowData = scope.data[i];
                        if (Array.isArray(rowData)) {
                            var row = table.GetRow(currentRow + i);
                            if (row) {
                                for (var j = 0; j < Math.min(rowData.length, scope.columns); j++) {
                                    var cell = row.GetCell(j);
                                    if (cell) {
                                        var cellContent = cell.GetContent();
                                        if (cellContent) {
                                            var para = cellContent.GetElement(0);
                                            if (para) {
                                                var cellText = String(rowData[j] || '');
                                                para.AddText(cellText);

                                                // 数字居右对齐
                                                if (!isNaN(parseFloat(cellText.replace(/[^\d.-]/g, '')))) {
                                                    para.SetJc('right');
                                                }

                                                // 合计行样式
                                                if (i === scope.data.length - 1 && scope.styling.totalRowStyle === 'bold') {
                                                    var textPr = Api.CreateTextPr();
                                                    textPr.SetBold(true);
                                                    para.SetTextPr(textPr);

                                                    if (typeof cell.SetShd === 'function') {
                                                        cell.SetShd('clear', 240, 240, 240);
                                                    }
                                                }

                                                // 交替行颜色
                                                else if (scope.styling.alternateRowColors && i % 2 === 1) {
                                                    if (typeof cell.SetShd === 'function') {
                                                        cell.SetShd('clear', 250, 250, 250);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 插入表格到文档
                var insertSuccess = false;

                // 先插入标题
                if (titlePara) {
                    if (typeof doc.Push === 'function') {
                        doc.Push(titlePara);
                    }
                }

                if (typeof doc.InsertContent === 'function') {
                    console.log('尝试使用InsertContent插入动态表格...');
                    var insertResult = doc.InsertContent([table], false);
                    if (insertResult) {
                        insertSuccess = true;
                        console.log('✅ 动态表格通过InsertContent插入成功');
                    }
                }

                if (!insertSuccess && typeof doc.Push === 'function') {
                    doc.Push(table);
                    insertSuccess = true;
                    console.log('✅ 动态表格通过Document.Push插入成功');
                }

                if (insertSuccess) {
                    return {
                        success: true,
                        message: 'Dynamic table inserted successfully',
                        method: 'Dynamic-Data-Binding',
                        parameters: {
                            title: scope.title,
                            rows: scope.rows,
                            columns: scope.columns,
                            dataSource: scope.metadata.dataSource,
                            recordCount: scope.data.length
                        }
                    };
                } else {
                    return { success: false, error: 'Dynamic table insertion failed' };
                }

            } catch (e) {
                console.log('动态表格插入失败:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }

    // 插入预设样式的表格
    async insertPresetTable(preset, customData) {
        const presets = {
            'simple': {
                rows: 3,
                columns: 3,
                headers: ['列1', '列2', '列3'],
                data: [
                    ['数据1', '数据2', '数据3'],
                    ['数据4', '数据5', '数据6']
                ]
            },
            'schedule': {
                rows: 6,
                columns: 7,
                headers: ['时间', '周一', '周二', '周三', '周四', '周五', '周六'],
                data: [
                    ['9:00', '', '', '', '', '', ''],
                    ['10:00', '', '', '', '', '', ''],
                    ['11:00', '', '', '', '', '', ''],
                    ['14:00', '', '', '', '', '', ''],
                    ['15:00', '', '', '', '', '', '']
                ]
            },
            'comparison': {
                rows: 5,
                columns: 3,
                headers: ['功能', '方案A', '方案B'],
                data: [
                    ['性能', '优秀', '良好'],
                    ['成本', '高', '中'],
                    ['维护', '复杂', '简单'],
                    ['扩展性', '强', '一般']
                ]
            }
        };

        const options = { ...presets[preset] || presets['simple'], ...customData };
        return this.insertTable(options);
    }

    // 处理表格点击事件
    async handleTableClick() {
        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            console.log('=== 表格点击检测 ===');

            try {
                // 获取当前选区
                var range = doc.GetRangeBySelect();
                if (!range) {
                    console.log('未检测到选区');
                    return { success: false, error: 'No selection found' };
                }

                console.log('选区信息:', range);

                // 检测是否在表格中
                var currentElement = null;
                var tableInfo = {
                    inTable: false,
                    tableIndex: -1,
                    rowIndex: -1,
                    cellIndex: -1,
                    cellText: '',
                    tableData: null
                };

                // 尝试获取当前段落
                if (typeof range.GetParagraph === 'function') {
                    var para = range.GetParagraph();
                    if (para) {
                        console.log('找到段落:', para);

                        // 检查段落是否在表格单元格中
                        var parent = para;
                        var cellFound = false;

                        // 向上查找父元素直到找到表格单元格
                        for (var level = 0; level < 10; level++) {
                            if (parent && typeof parent.GetParent === 'function') {
                                parent = parent.GetParent();
                                console.log(`父级 ${level}:`, parent, typeof parent);

                                // 检查是否是表格单元格
                                if (parent && typeof parent.GetClassType === 'function') {
                                    var classType = parent.GetClassType();
                                    console.log(`ClassType ${level}:`, classType);

                                    if (classType === 'CDocumentContent' || classType === 'CTableCell') {
                                        cellFound = true;
                                        console.log('发现表格单元格!');

                                        // 获取单元格文本
                                        if (typeof parent.GetElementsCount === 'function') {
                                            var elementCount = parent.GetElementsCount();
                                            console.log('单元格元素数量:', elementCount);

                                            for (var i = 0; i < elementCount; i++) {
                                                var element = parent.GetElement(i);
                                                if (element && typeof element.GetText === 'function') {
                                                    tableInfo.cellText += element.GetText();
                                                }
                                            }
                                        }
                                        break;
                                    }
                                }
                            } else {
                                break;
                            }
                        }

                        if (cellFound) {
                            tableInfo.inTable = true;
                            console.log('✅ 检测到表格点击');
                        }
                    }
                }

                // 方法2：尝试通过文档元素查找表格
                if (!tableInfo.inTable) {
                    console.log('尝试方法2：遍历文档元素查找表格...');

                    for (var i = 0; i < 20; i++) {
                        var element = doc.GetElement(i);
                        if (!element) break;

                        if (typeof element.GetClassType === 'function') {
                            var classType = element.GetClassType();
                            console.log(`文档元素 ${i}:`, classType);

                            if (classType === 'CTable') {
                                console.log(`发现表格 ${i}!`);
                                tableInfo.tableIndex = i;
                                tableInfo.inTable = true;

                                // 获取表格信息
                                if (typeof element.GetRowsCount === 'function') {
                                    var rowCount = element.GetRowsCount();
                                    console.log('表格行数:', rowCount);

                                    tableInfo.tableData = {
                                        rows: rowCount,
                                        columns: 0,
                                        content: []
                                    };

                                    // 获取表格内容
                                    for (var r = 0; r < Math.min(rowCount, 5); r++) {
                                        if (typeof element.GetRow === 'function') {
                                            var row = element.GetRow(r);
                                            if (row && typeof row.GetCellsCount === 'function') {
                                                var cellCount = row.GetCellsCount();
                                                tableInfo.tableData.columns = Math.max(tableInfo.tableData.columns, cellCount);

                                                var rowData = [];
                                                for (var c = 0; c < cellCount; c++) {
                                                    if (typeof row.GetCell === 'function') {
                                                        var cell = row.GetCell(c);
                                                        if (cell && typeof cell.GetContent === 'function') {
                                                            var cellContent = cell.GetContent();
                                                            var cellText = '';
                                                            if (cellContent && typeof cellContent.GetElementsCount === 'function') {
                                                                var elemCount = cellContent.GetElementsCount();
                                                                for (var e = 0; e < elemCount; e++) {
                                                                    var cellElem = cellContent.GetElement(e);
                                                                    if (cellElem && typeof cellElem.GetText === 'function') {
                                                                        cellText += cellElem.GetText();
                                                                    }
                                                                }
                                                            }
                                                            rowData.push(cellText.trim());
                                                        }
                                                    }
                                                }
                                                tableInfo.tableData.content.push(rowData);
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }

                if (tableInfo.inTable) {
                    console.log('表格点击信息:', tableInfo);
                    return {
                        success: true,
                        message: 'Table click detected',
                        data: {
                            clickType: 'table',
                            tableIndex: tableInfo.tableIndex,
                            rowIndex: tableInfo.rowIndex,
                            cellIndex: tableInfo.cellIndex,
                            cellText: tableInfo.cellText,
                            tableData: tableInfo.tableData,
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    };
                } else {
                    console.log('未检测到表格点击');
                    return {
                        success: false,
                        error: 'Click not in table',
                        data: {
                            clickType: 'document',
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    };
                }

            } catch (e) {
                console.log('表格点击检测失败:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }
}