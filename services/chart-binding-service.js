// services/chart-binding-service.js
import { ChartTypeDetector } from './chart-type-detector.js';
import { ChartDataBinder } from './chart-data-binder.js';
import { ChartDetector } from './chart-detector.js';

export class ChartBindingService {
    constructor(editorService) {
        this.editor = editorService;

        // 注入专门的服务类
        this.typeDetector = new ChartTypeDetector();
        this.dataBinder = new ChartDataBinder();
        this.chartDetector = new ChartDetector();
    }

    /**
     * 为文档中的图表绑定隐藏数据
     * @param {Object} chartData - 图表数据
     */
    async bindDataToChart(chartData) {
        console.log('开始图表数据绑定流程');
        console.log('从宿主页面接收到的图表数据:', chartData);

        // 检查是否从宿主页面传递了有效数据（在runInDoc外部检查）
        if (!chartData || !chartData.data || !chartData.metadata) {
            console.log('❌ 未收到有效的图表数据，无法进行绑定');
            return {
                success: false,
                error: '未从宿主页面接收到有效的图表数据，请确保调用时传递了完整的图表数据结构',
                data: {
                    timestamp: new Date().toLocaleString()
                }
            };
        }

        console.log('✅ 使用宿主页面传递的图表数据');

        // 由于runInDoc是隔离环境，将数据直接嵌入到函数体中
        const chartDataJson = JSON.stringify(chartData);
        console.log('🔍 准备嵌入的数据JSON长度:', chartDataJson.length);

        // 创建包含数据的函数字符串，然后用new Function执行
        const funcStr = `
            const doc = Api.GetDocument();
            console.log('=== 图表数据绑定开始 ===');

            try {
                // 直接使用嵌入的数据
                const bindingData = ${chartDataJson};
                console.log('📊 使用嵌入的绑定数据:', bindingData);

                if (!bindingData || !bindingData.data || !bindingData.metadata) {
                    console.log('❌ 嵌入的数据无效');
                    return {
                        success: false,
                        error: '嵌入的图表数据无效',
                        data: {
                            timestamp: new Date().toLocaleString()
                        }
                    };
                }

                console.log('📊 最终使用的绑定数据:', bindingData);

                // 直接在内部实现图表扫描逻辑
                console.log('开始扫描文档中的图表...');
                const documentLevelCharts = [];

                try {
                    if (typeof doc.GetAllDrawingObjects === 'function') {
                        const docDrawingObjects = doc.GetAllDrawingObjects();
                        console.log('📄 文档级绘图对象:', docDrawingObjects);

                        if (docDrawingObjects && docDrawingObjects.length > 0) {
                            console.log('🎯 找到 ' + docDrawingObjects.length + ' 个文档级绘图对象！');

                            for (let j = 0; j < docDrawingObjects.length; j++) {
                                const drawingObj = docDrawingObjects[j];
                                console.log('📊 绘图对象 ' + j + ':', drawingObj);

                                let drawingType = 'unknown';
                                if (typeof drawingObj.GetClassType === 'function') {
                                    drawingType = drawingObj.GetClassType();
                                }
                                console.log('📊 绘图对象类型:', drawingType);

                                // 检查是否是图表类型
                                if (drawingType === 'chart' ||
                                    drawingType.includes('Chart') ||
                                    drawingType.includes('Drawing') ||
                                    drawingType.includes('Shape') ||
                                    drawingType.includes('Image')) {

                                    console.log('✅ 发现文档级图表/图形:', drawingType);

                                    const chartInfo = {
                                        element: drawingObj,
                                        elementType: drawingType,
                                        index: 'doc_drawing_' + j,
                                        drawingIndex: j,
                                        isDocumentLevel: true,
                                        source: 'document-level'
                                    };

                                    documentLevelCharts.push(chartInfo);
                                }
                            }
                        }
                    } else {
                        console.log('文档不支持GetAllDrawingObjects方法');
                    }
                } catch (docError) {
                    console.log('检查文档级绘图对象失败:', docError);
                }

                console.log('扫描结果 - 文档级图表数:', documentLevelCharts.length);

                const boundCharts = [];
                const bindingResults = [];

                // 处理找到的图表
                documentLevelCharts.forEach(function(chartInfo) {
                    console.log('处理图表:', chartInfo.index, chartInfo.elementType);

                    // 简化的图表类型识别（直接内联实现）
                    let detailedChartType = {
                        category: 'chart',
                        specificType: 'unknown',
                        description: '图表',
                        confidence: 0.8
                    };

                    // 尝试获取准确的图表类型
                    try {
                        // 方法1: 尝试GetPrevChart
                        if (typeof chartInfo.element.GetPrevChart === 'function') {
                            console.log('🔍 优先尝试GetPrevChart方法');

                            let prevChart;
                            try {
                                prevChart = chartInfo.element.GetPrevChart();
                                console.log('📊 GetPrevChart返回:', prevChart);
                            } catch (sdkError) {
                                console.log('🚨 GetPrevChart触发SDK错误:', sdkError.message);
                                prevChart = null;
                            }

                            if (prevChart && typeof prevChart.GetChartType === 'function') {
                                try {
                                    const chartType = prevChart.GetChartType();
                                    console.log('📊 GetChartType返回:', chartType);
                                    if (chartType) {
                                        console.log('✅ 通过GetPrevChart获得准确图表类型:', chartType);
                                        detailedChartType.specificType = chartType;
                                        detailedChartType.description = '图表 (' + chartType + ')';
                                        detailedChartType.confidence = 1.0;
                                    } else {
                                        console.log('⚠️ GetChartType返回空值');
                                    }
                                } catch (chartTypeError) {
                                    console.log('🚨 GetChartType调用失败:', chartTypeError.message);
                                }
                            } else {
                                console.log('⚠️ prevChart为空，尝试其他方法');
                            }
                        }

                        // 方法2: 如果GetPrevChart失败，尝试直接调用GetChart
                        if (detailedChartType.specificType === 'unknown' && typeof chartInfo.element.GetChart === 'function') {
                            console.log('🔍 尝试GetChart方法');
                            try {
                                const chart = chartInfo.element.GetChart();
                                console.log('📊 GetChart返回:', chart);
                                if (chart && typeof chart.GetChartType === 'function') {
                                    const chartType = chart.GetChartType();
                                    console.log('📊 Chart.GetChartType返回:', chartType);
                                    if (chartType) {
                                        console.log('✅ 通过GetChart获得准确图表类型:', chartType);
                                        detailedChartType.specificType = chartType;
                                        detailedChartType.description = '图表 (' + chartType + ')';
                                        detailedChartType.confidence = 1.0;
                                    }
                                }
                            } catch (chartError) {
                                console.log('🚨 GetChart调用失败:', chartError.message);
                            }
                        }

                        // 方法2.5: 尝试直接调用图表元素的GetChartType方法
                        if (detailedChartType.specificType === 'unknown' && typeof chartInfo.element.GetChartType === 'function') {
                            console.log('🔍 尝试直接调用图表元素的GetChartType方法');
                            try {
                                const chartType = chartInfo.element.GetChartType();
                                console.log('📊 直接GetChartType返回:', chartType);
                                if (chartType && chartType !== 'chart') {
                                    console.log('✅ 通过直接GetChartType获得准确图表类型:', chartType);
                                    detailedChartType.specificType = chartType;
                                    detailedChartType.description = '图表 (' + chartType + ')';
                                    detailedChartType.confidence = 1.0;
                                }
                            } catch (chartTypeError) {
                                console.log('🚨 直接GetChartType调用失败:', chartTypeError.message);
                            }
                        }

                    } catch (chartTypeError) {
                        console.log('🚨 图表类型识别失败:', chartTypeError.message);
                    }

                    console.log('📊 图表类型识别结果:', detailedChartType);

                    // 生成唯一标识符
                    const uniqueId = 'chart_' + chartInfo.drawingIndex + '_' + Date.now();

                    // 创建绑定信息
                    const bindingInfo = {
                        chartIndex: chartInfo.index,
                        chartType: chartInfo.elementType,
                        detailedChartType: detailedChartType,
                        uniqueId: uniqueId,
                        boundData: bindingData.data || {},
                        bindingId: 'doc_chart_' + chartInfo.drawingIndex + '_' + Date.now(),
                        boundAt: new Date().toISOString(),
                        metadata: bindingData.metadata || {},
                        isDocumentLevel: true,
                        drawingIndex: chartInfo.drawingIndex
                    };

                    // 简化的数据绑定（使用内存存储）
                    let bindingResult = {
                        directBinding: false,
                        bindingMethod: 'memory-storage',
                        storageKey: null
                    };

                    try {
                        // 尝试直接绑定
                        if (typeof chartInfo.element.SetCustomProperty === 'function') {
                            chartInfo.element.SetCustomProperty('chartData', JSON.stringify(bindingInfo));
                            bindingResult.directBinding = true;
                            bindingResult.bindingMethod = 'custom-property';
                            console.log('✅ 直接在图表元素上绑定数据成功');
                        } else {
                            // 使用内存存储
                            if (!window.chartDataStorage) {
                                window.chartDataStorage = {};
                            }
                            const storageKey = 'doc_chart_' + chartInfo.drawingIndex + '_' + Date.now();
                            window.chartDataStorage[storageKey] = bindingInfo;
                            bindingResult.storageKey = storageKey;
                            console.log('✅ 使用内存存储绑定图表数据，存储键:', storageKey);
                        }
                    } catch (bindingError) {
                        console.log('数据绑定失败:', bindingError.message);
                    }

                    // 合并绑定结果
                    Object.assign(bindingInfo, bindingResult);

                    boundCharts.push(chartInfo);
                    bindingResults.push(bindingInfo);

                    console.log('🎯 图表绑定完成:', {
                        uniqueId: uniqueId,
                        chartType: detailedChartType.description,
                        specificType: detailedChartType.specificType,
                        bindingMethod: bindingResult.bindingMethod
                    });

                    // 验证绑定：立即获取刚绑定的数据并打印
                    console.log('🔍 验证数据绑定 - 开始获取绑定数据...');
                    try {
                        let retrievedData = null;
                        let retrievalMethod = null;

                        // 方法1：从自定义属性获取
                        if (bindingResult.bindingMethod === 'custom-property') {
                            if (typeof chartInfo.element.GetCustomProperty === 'function') {
                                const customData = chartInfo.element.GetCustomProperty('chartData');
                                if (customData) {
                                    retrievedData = JSON.parse(customData);
                                    retrievalMethod = 'custom-property';
                                    console.log('✅ 从自定义属性成功获取绑定数据');
                                }
                            }
                        }

                        // 方法2：从内存存储获取
                        if (!retrievedData && bindingResult.bindingMethod === 'memory-storage' && bindingResult.storageKey) {
                            if (window.chartDataStorage && window.chartDataStorage[bindingResult.storageKey]) {
                                retrievedData = window.chartDataStorage[bindingResult.storageKey];
                                retrievalMethod = 'memory-storage';
                                console.log('✅ 从内存存储成功获取绑定数据，存储键:', bindingResult.storageKey);
                            }
                        }

                        if (retrievedData) {
                            console.log('📋 验证成功！获取到的绑定数据:');
                            console.log('   - 获取方法:', retrievalMethod);
                            console.log('   - 图表标题:', retrievedData.boundData?.title || '未知');
                            console.log('   - 图表类型:', retrievedData.boundData?.type || '未知');
                            console.log('   - 数据源:', retrievedData.boundData?.dataSource || '未知');
                            console.log('   - 绑定ID:', retrievedData.bindingId);
                            console.log('   - 绑定时间:', retrievedData.boundAt);

                            // 打印关键指标
                            if (retrievedData.boundData?.metrics) {
                                console.log('   - 关键指标:');
                                const metrics = retrievedData.boundData.metrics;
                                if (metrics.totalSales) console.log('     * 总销售额: ¥' + metrics.totalSales.toLocaleString());
                                if (metrics.growthRate) console.log('     * 增长率: ' + metrics.growthRate + '%');
                                if (metrics.topProduct) console.log('     * 热门产品: ' + metrics.topProduct);
                                if (metrics.targetAchievement) console.log('     * 目标达成: ' + metrics.targetAchievement + '%');
                            }

                            // 打印数据系列预览
                            if (retrievedData.boundData?.series && retrievedData.boundData.series.length > 0) {
                                console.log('   - 数据系列:');
                                retrievedData.boundData.series.forEach(function(series) {
                                    console.log('     * ' + series.name + ': [' + series.data.slice(0, 3).join(', ') +
                                              (series.data.length > 3 ? '...' : '') + ']');
                                });
                            }

                            console.log('   - 完整数据对象:', retrievedData);
                        } else {
                            console.log('⚠️ 数据绑定验证失败：无法获取刚绑定的数据');
                            console.log('   - 尝试的方法:', bindingResult.bindingMethod);
                            console.log('   - 存储键:', bindingResult.storageKey || '无');
                        }
                    } catch (verifyError) {
                        console.log('❌ 数据绑定验证出错:', verifyError.message);
                    }
                });

                // 如果没有找到图表，提供指导
                if (boundCharts.length === 0) {
                    console.log('💡 提示: 当前文档中没有检测到图表元素');
                }

                return {
                    success: true,
                    message: boundCharts.length > 0 ?
                        ('成功绑定 ' + boundCharts.length + ' 个图表') :
                        ('文档扫描完成，但未发现图表元素。'),
                    data: {
                        chartsFound: boundCharts.length,
                        bindingResults: bindingResults,
                        boundData: bindingData,
                        bindingMethod: 'inline-processing',
                        timestamp: new Date().toLocaleString()
                    }
                };

            } catch (error) {
                console.log('❌ 图表数据绑定失败:', error);
                return {
                    success: false,
                    error: error.message,
                    data: {
                        timestamp: new Date().toLocaleString()
                    }
                };
            }
        `;

        // 使用new Function执行包含数据的代码
        const dynamicFunction = new Function(funcStr);
        return this.editor.runInDoc(dynamicFunction);
    }

    /**
     * 检测图表点击并获取绑定的数据
     */
    async detectChartClick() {
        // 在runInDoc外部获取服务引用，避免作用域问题
        const chartDetector = this.chartDetector;
        const dataBinder = this.dataBinder;
        const typeDetector = this.typeDetector;

        return new Promise((resolve) => {
            this.editor.runInDoc(() => {
                const doc = Api.GetDocument();
                console.log('=== 图表点击检测 ===');

                try {
                    // 获取当前选区信息
                    const selectionInfo = chartDetector.getCurrentSelection(doc);

                    // 扫描所有图表，但使用安全模式
                    console.log('🔒 使用安全模式扫描图表...');
                    const scanResults = chartDetector.scanDocument(doc);
                    const chartDetectionResults = [];

                    // 检查文档级图表的绑定数据 - 使用安全模式
                    scanResults.documentLevelCharts.forEach((chartInfo) => {
                        console.log('🔒 安全模式检查图表:', chartInfo.index);

                        try {
                            // 使用安全模式识别图表类型，避免触发SDK错误
                            let detailedChartType;
                            try {
                                detailedChartType = typeDetector.identifyChartType(chartInfo.element, chartInfo.elementType);
                            } catch (identifyError) {
                                console.log('🚨 标准识别失败，使用安全模式:', identifyError.message);
                                detailedChartType = {
                                    category: 'chart',
                                    specificType: 'unknown',
                                    description: '图表（安全模式）',
                                    properties: {
                                        safeMode: true,
                                        error: identifyError.message
                                    },
                                    confidence: 0.5
                                };
                            }
                            console.log('📊 安全模式图表类型识别:', detailedChartType);

                            // 生成图表唯一标识符时也要小心
                            let uniqueId;
                            try {
                                uniqueId = typeDetector.generateChartUniqueId(
                                    chartInfo.element,
                                    chartInfo.elementType,
                                    chartInfo.drawingIndex
                                );
                            } catch (idError) {
                                console.log('🚨 生成唯一ID失败，使用基础ID:', idError.message);
                                uniqueId = `chart_safe_${chartInfo.drawingIndex}_${Date.now()}`;
                            }

                            const chartResult = {
                                chartIndex: chartInfo.index,
                                chartType: chartInfo.elementType,
                                detailedChartType: detailedChartType,
                                uniqueId: uniqueId,
                                boundData: null,
                                hasBindingData: false,
                                isDocumentLevel: true,
                                drawingIndex: chartInfo.drawingIndex
                            };

                            // 检查绑定数据
                            const dataResult = dataBinder.getBoundData(chartInfo.element, chartInfo.drawingIndex);
                            Object.assign(chartResult, dataResult);

                            // 如果找到绑定数据，记录匹配信息
                            if (chartResult.hasBindingData && chartResult.boundData) {
                                console.log('🔍 安全模式检查图表数据匹配:', {
                                    currentId: uniqueId,
                                    boundId: chartResult.boundData.uniqueId,
                                    chartType: detailedChartType.description
                                });
                            }

                            chartDetectionResults.push(chartResult);

                        } catch (chartError) {
                            console.log('🚨 安全检查图表失败，跳过:', chartError.message);

                            // 即使失败也添加一个基础记录
                            chartDetectionResults.push({
                                chartIndex: chartInfo.index,
                                chartType: chartInfo.elementType,
                                detailedChartType: {
                                    category: 'chart',
                                    specificType: 'unknown',
                                    description: '图表（安全模式）',
                                    confidence: 0.5
                                },
                                uniqueId: `chart_safe_${chartInfo.drawingIndex}_${Date.now()}`,
                                boundData: null,
                                hasBindingData: false,
                                isDocumentLevel: true,
                                drawingIndex: chartInfo.drawingIndex,
                                safeMode: true
                            });
                        }
                    });

                    // 查找被点击的图表 - 内联处理避免this作用域问题
                    let clickedChart = null;
                    if (chartDetectionResults.length > 0) {
                        // 简单策略：如果有选区，取最后一个有数据的图表
                        for (let i = chartDetectionResults.length - 1; i >= 0; i--) {
                            if (chartDetectionResults[i].hasBindingData) {
                                clickedChart = chartDetectionResults[i];
                                break;
                            }
                        }
                        // 如果没有找到有数据的图表，取最后一个
                        if (!clickedChart) {
                            clickedChart = chartDetectionResults[chartDetectionResults.length - 1];
                        }
                    }

                    if (clickedChart && clickedChart.hasBindingData) {
                        const result = {
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
                                    hasSelection: selectionInfo.hasSelection
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
                                    hasSelection: selectionInfo.hasSelection
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
                                hasSelection: selectionInfo.hasSelection,
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

    /**
     * 获取文档中所有图表的绑定数据摘要
     */
    async getChartBindingSummary() {
        // 在runInDoc外部获取服务引用，避免作用域问题
        const chartDetector = this.chartDetector;
        const dataBinder = this.dataBinder;

        return new Promise((resolve) => {
            this.editor.runInDoc(() => {
                const doc = Api.GetDocument();
                console.log('=== 获取图表绑定摘要 ===');

                try {
                    const scanResults = chartDetector.scanDocument(doc);
                    const summary = {
                        totalCharts: 0,
                        chartsWithData: 0,
                        bindingSummary: []
                    };

                    // 统计文档级图表
                    scanResults.documentLevelCharts.forEach((chartInfo, index) => {
                        summary.totalCharts++;
                        const dataResult = dataBinder.getBoundData(chartInfo.element, chartInfo.drawingIndex);

                        if (dataResult.hasBindingData) {
                            summary.chartsWithData++;
                        }

                        summary.bindingSummary.push({
                            chartIndex: chartInfo.index,
                            chartType: chartInfo.elementType,
                            hasBindingData: dataResult.hasBindingData,
                            bindingPreview: dataResult.boundData?.bindingId || 'unknown',
                            source: 'document-level'
                        });
                    });

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

    /**
     * 清理临时数据
     */
    cleanupTempData() {
        this.dataBinder.cleanupTempData();
    }
}