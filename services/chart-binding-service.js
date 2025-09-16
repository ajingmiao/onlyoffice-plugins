// services/chart-binding-service.js
export class ChartBindingService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 为文档中的图表绑定隐藏数据
    async bindDataToChart(chartData) {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                var doc = Api.GetDocument();

                console.log('=== 图表数据绑定开始 ===');
                console.log('绑定数据:', chartData);

                // 将chartData复制到Asc.scope中，避免作用域问题
                window.Asc = window.Asc || {};
                window.Asc.scope = {
                    chartData: chartData
                };

                try {
                    var boundCharts = [];
                    var bindingResults = [];

                    // 扫描文档中的所有图表
                    for (var i = 0; i < 100; i++) {
                        try {
                            var element = doc.GetElement(i);
                            if (!element) {
                                console.log(`文档元素 ${i}: null，结束遍历`);
                                break;
                            }

                            var elementType = 'unknown';
                            if (typeof element.GetClassType === 'function') {
                                elementType = element.GetClassType();
                            }

                            console.log(`文档元素 ${i}: ${elementType}`);

                            // 检测图表、图形、绘图等元素
                            if (elementType.includes('Drawing') ||
                                elementType.includes('Chart') ||
                                elementType === 'CDrawing' ||
                                elementType === 'GraphicFrame') {

                                console.log(`✅ 发现图表/图形在位置 ${i}: ${elementType}`);

                                // 使用Asc.scope中的数据
                                var chartDataToUse = window.Asc.scope.chartData;

                                // 创建绑定信息
                                var bindingInfo = {
                                    chartIndex: i,
                                    chartType: elementType,
                                    boundData: chartDataToUse.data || {},
                                    bindingId: 'chart_' + i + '_' + Date.now(),
                                    boundAt: new Date().toISOString(),
                                    metadata: chartDataToUse.metadata || {}
                                };

                                // 方法1: 尝试在图表周围创建隐藏的内容控件来存储数据
                                try {
                                    var hiddenMarker = doc.AddTextContentControl();
                                    if (hiddenMarker) {
                                        var tagData = 'chart-data:' + JSON.stringify(bindingInfo);

                                        if (typeof hiddenMarker.SetTag === 'function') {
                                            hiddenMarker.SetTag(tagData);
                                            console.log('✅ 图表数据已绑定到Tag:', tagData.substring(0, 100) + '...');
                                        }

                                        if (typeof hiddenMarker.SetAlias === 'function') {
                                            hiddenMarker.SetAlias('图表数据: ' + bindingInfo.bindingId);
                                        }

                                        // 设置隐藏标记（不显示内容）
                                        if (typeof hiddenMarker.SetPlaceholderText === 'function') {
                                            hiddenMarker.SetPlaceholderText('');
                                        }

                                        // 添加隐藏文本
                                        if (typeof hiddenMarker.AddText === 'function') {
                                            hiddenMarker.AddText(''); // 空文本，不显示
                                        }

                                        bindingInfo.markerCreated = true;
                                        bindingInfo.markerMethod = 'content-control';
                                    }
                                } catch (markerError) {
                                    console.log('创建内容控件标记失败:', markerError);
                                    bindingInfo.markerCreated = false;
                                    bindingInfo.markerError = markerError.message;
                                }

                                // 方法2: 尝试直接在元素上设置自定义属性（如果支持）
                                try {
                                    if (typeof element.SetCustomProperty === 'function') {
                                        element.SetCustomProperty('chartData', JSON.stringify(bindingInfo));
                                        bindingInfo.directBinding = true;
                                        console.log('✅ 直接在图表元素上绑定数据成功');
                                    } else {
                                        bindingInfo.directBinding = false;
                                        console.log('图表元素不支持SetCustomProperty');
                                    }
                                } catch (directError) {
                                    console.log('直接绑定失败:', directError);
                                    bindingInfo.directBinding = false;
                                    bindingInfo.directError = directError.message;
                                }

                                boundCharts.push({
                                    element: element,
                                    elementType: elementType,
                                    index: i
                                });

                                bindingResults.push(bindingInfo);
                            }

                        } catch (elementError) {
                            console.log(`读取文档元素 ${i} 出错:`, elementError);
                        }
                    }

                    var result = {
                        success: true,
                        message: `成功绑定 ${boundCharts.length} 个图表`,
                        data: {
                            chartsFound: boundCharts.length,
                            bindingResults: bindingResults,
                            boundData: window.Asc.scope.chartData,
                            bindingMethod: 'content-control-marker',
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    };

                    console.log('✅ 图表数据绑定完成:', result);
                    resolve(result);

                } catch (error) {
                    console.log('❌ 图表数据绑定失败:', error);
                    resolve({
                        success: false,
                        error: error.message,
                        data: {
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    });
                }

            }, { async: false, cb: (res) => resolve(res) });
        });
    }

    // 检测图表点击并获取绑定的数据
    async detectChartClick() {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                var doc = Api.GetDocument();
                var range = doc.GetRangeBySelect();

                console.log('=== 图表点击检测 ===');
                console.log('当前选区:', range);

                try {
                    var chartDetectionResults = [];

                    // 扫描文档中的所有图表和相关的数据绑定
                    for (var i = 0; i < 100; i++) {
                        try {
                            var element = doc.GetElement(i);
                            if (!element) break;

                            var elementType = 'unknown';
                            if (typeof element.GetClassType === 'function') {
                                elementType = element.GetClassType();
                            }

                            // 检查是否是图表相关元素
                            if (elementType.includes('Drawing') ||
                                elementType.includes('Chart') ||
                                elementType === 'CDrawing' ||
                                elementType === 'GraphicFrame') {

                                console.log(`检查图表 ${i}: ${elementType}`);

                                var chartInfo = {
                                    chartIndex: i,
                                    chartType: elementType,
                                    boundData: null,
                                    hasBindingData: false
                                };

                                // 方法1: 检查自定义属性
                                try {
                                    if (typeof element.GetCustomProperty === 'function') {
                                        var customData = element.GetCustomProperty('chartData');
                                        if (customData) {
                                            chartInfo.boundData = JSON.parse(customData);
                                            chartInfo.hasBindingData = true;
                                            chartInfo.bindingMethod = 'custom-property';
                                            console.log('✅ 从自定义属性获取到图表数据:', chartInfo.boundData);
                                        }
                                    }
                                } catch (customError) {
                                    console.log('读取自定义属性失败:', customError);
                                }

                                chartDetectionResults.push(chartInfo);
                            }

                            // 检查内容控件标记（包含图表数据）
                            else if (elementType.includes('ContentControl') || elementType.includes('Sdt')) {
                                try {
                                    var tag = '';
                                    if (typeof element.GetTag === 'function') {
                                        tag = element.GetTag();
                                    }

                                    if (tag && tag.startsWith('chart-data:')) {
                                        console.log('✅ 发现图表数据标记:', tag.substring(0, 100) + '...');

                                        var bindingData = JSON.parse(tag.substring('chart-data:'.length));

                                        // 查找对应的图表信息
                                        var targetChart = chartDetectionResults.find(chart =>
                                            chart.chartIndex === bindingData.chartIndex
                                        );

                                        if (targetChart) {
                                            targetChart.boundData = bindingData;
                                            targetChart.hasBindingData = true;
                                            targetChart.bindingMethod = 'content-control-marker';
                                            console.log('✅ 图表数据绑定匹配成功');
                                        } else {
                                            // 创建新的图表记录
                                            chartDetectionResults.push({
                                                chartIndex: bindingData.chartIndex,
                                                chartType: bindingData.chartType,
                                                boundData: bindingData,
                                                hasBindingData: true,
                                                bindingMethod: 'content-control-marker'
                                            });
                                        }
                                    }
                                } catch (tagError) {
                                    console.log('解析图表数据标记失败:', tagError);
                                }
                            }

                        } catch (elementError) {
                            console.log(`检查元素 ${i} 出错:`, elementError);
                        }
                    }

                    // 查找被点击的图表
                    var clickedChart = null;
                    var hasSelection = !!range;

                    if (chartDetectionResults.length > 0) {
                        // 简单策略：如果有选区，取最后一个有数据的图表
                        // 更精确的方法需要位置比较，但比较复杂
                        for (var i = chartDetectionResults.length - 1; i >= 0; i--) {
                            if (chartDetectionResults[i].hasBindingData) {
                                clickedChart = chartDetectionResults[i];
                                break;
                            }
                        }

                        // 如果没有找到有数据的图表，取最后一个
                        if (!clickedChart && chartDetectionResults.length > 0) {
                            clickedChart = chartDetectionResults[chartDetectionResults.length - 1];
                        }
                    }

                    if (clickedChart && clickedChart.hasBindingData) {
                        var result = {
                            success: true,
                            message: 'Chart with bound data detected',
                            data: {
                                clickType: 'chart',
                                chartInfo: clickedChart,
                                boundData: clickedChart.boundData.boundData || clickedChart.boundData,
                                bindingMetadata: {
                                    bindingId: clickedChart.boundData.bindingId,
                                    boundAt: clickedChart.boundData.boundAt,
                                    bindingMethod: clickedChart.bindingMethod
                                },
                                detectionSummary: {
                                    totalChartsFound: chartDetectionResults.length,
                                    chartsWithData: chartDetectionResults.filter(c => c.hasBindingData).length,
                                    hasSelection: hasSelection
                                },
                                timestamp: new Date().toLocaleString('zh-CN')
                            }
                        };

                        console.log('✅ 检测到图表点击，包含绑定数据:', result);
                        resolve(result);

                    } else if (chartDetectionResults.length > 0) {
                        // 有图表但没有绑定数据
                        resolve({
                            success: true,
                            message: 'Chart detected but no bound data',
                            data: {
                                clickType: 'chart',
                                chartInfo: clickedChart || chartDetectionResults[0],
                                boundData: null,
                                detectionSummary: {
                                    totalChartsFound: chartDetectionResults.length,
                                    chartsWithData: 0,
                                    hasSelection: hasSelection
                                },
                                timestamp: new Date().toLocaleString('zh-CN')
                            }
                        });

                    } else {
                        // 没有检测到图表
                        resolve({
                            success: false,
                            error: 'No charts found in document',
                            data: {
                                clickType: 'other',
                                hasSelection: hasSelection,
                                timestamp: new Date().toLocaleString('zh-CN')
                            }
                        });
                    }

                } catch (error) {
                    console.log('❌ 图表点击检测失败:', error);
                    resolve({
                        success: false,
                        error: error.message,
                        data: {
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    });
                }

            }, { async: false, cb: (res) => resolve(res) });
        });
    }

    // 获取文档中所有图表的绑定数据摘要
    async getChartBindingSummary() {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                var doc = Api.GetDocument();

                console.log('=== 获取图表绑定摘要 ===');

                try {
                    var summary = {
                        totalCharts: 0,
                        chartsWithData: 0,
                        bindingSummary: []
                    };

                    // 扫描所有图表和绑定数据
                    for (var i = 0; i < 100; i++) {
                        try {
                            var element = doc.GetElement(i);
                            if (!element) break;

                            var elementType = 'unknown';
                            if (typeof element.GetClassType === 'function') {
                                elementType = element.GetClassType();
                            }

                            if (elementType.includes('Drawing') ||
                                elementType.includes('Chart') ||
                                elementType === 'CDrawing') {

                                summary.totalCharts++;

                                var hasData = false;
                                var bindingInfo = null;

                                // 检查是否有绑定数据
                                try {
                                    if (typeof element.GetCustomProperty === 'function') {
                                        var customData = element.GetCustomProperty('chartData');
                                        if (customData) {
                                            hasData = true;
                                            bindingInfo = JSON.parse(customData);
                                        }
                                    }
                                } catch (e) {
                                    // 忽略解析错误
                                }

                                if (hasData) {
                                    summary.chartsWithData++;
                                }

                                summary.bindingSummary.push({
                                    chartIndex: i,
                                    chartType: elementType,
                                    hasBindingData: hasData,
                                    bindingPreview: bindingInfo ?
                                        (bindingInfo.bindingId || 'unknown') : null
                                });
                            }

                        } catch (elementError) {
                            console.log(`检查元素 ${i} 出错:`, elementError);
                        }
                    }

                    resolve({
                        success: true,
                        data: summary,
                        timestamp: new Date().toLocaleString('zh-CN')
                    });

                } catch (error) {
                    resolve({
                        success: false,
                        error: error.message
                    });
                }

            }, { async: false, cb: (res) => resolve(res) });
        });
    }
}