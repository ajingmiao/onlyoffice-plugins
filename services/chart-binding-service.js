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
     * @param {Object} chartData - 图表数据（可选，会使用测试数据）
     */
    async bindDataToChart(chartData) {
        console.log('开始图表数据绑定流程');

        return this.editor.runInDoc(() => {
            const doc = Api.GetDocument();
            console.log('=== 图表数据绑定开始 ===');

            try {
                // 直接在runInDoc内部创建所有需要的实例和数据
                console.log('在文档环境中创建服务实例...');

                // 创建测试绑定数据
                const bindingData = {
                    data: {
                        title: '测试图表数据',
                        type: 'line-chart',
                        dataSource: '测试系统',
                        category: '测试分析',
                        period: '2024年测试期间',
                        metrics: {
                            totalSales: 1000000,
                            growthRate: 12.5,
                            topProduct: '测试产品A',
                            targetAchievement: 110
                        },
                        series: [
                            { name: '实际销售', data: [100, 120, 110, 140] },
                            { name: '目标销售', data: [110, 115, 120, 130] }
                        ],
                        categories: ['Q1', 'Q2', 'Q3', 'Q4']
                    },
                    metadata: {
                        bindingId: 'test_chart_' + Date.now(),
                        boundAt: new Date().toISOString(),
                        dataVersion: '1.0',
                        refreshInterval: 3600,
                        lastUpdated: new Date().toLocaleString(),
                        permissions: {
                            canEdit: true,
                            canRefresh: true,
                            canExport: true
                        }
                    }
                };

                console.log('使用测试绑定数据:', bindingData);

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
                documentLevelCharts.forEach((chartInfo) => {
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

                        // 方法3: 尝试OOXML分析来获取图表类型
                        if (detailedChartType.specificType === 'unknown') {
                            console.log('🔍 尝试OOXML分析...');
                            const xmlMethods = ['GetOOXML', 'GetXML', 'ToXML', 'GetDocumentXML'];

                            for (const xmlMethod of xmlMethods) {
                                if (typeof chartInfo.element[xmlMethod] === 'function') {
                                    try {
                                        console.log(`🔍 尝试调用方法: ${xmlMethod}`);
                                        const xmlResult = chartInfo.element[xmlMethod]();
                                        if (xmlResult && typeof xmlResult === 'string' && xmlResult.length > 0) {
                                            console.log(`✅ 成功获取XML内容，方法: ${xmlMethod}`);
                                            console.log('📄 XML内容长度:', xmlResult.length);
                                            console.log('📄 XML内容片段:', xmlResult.substring(0, 300) + '...');

                                            // 分析XML内容中的图表类型模式
                                            const chartTypePatterns = {
                                                pie: /pieChart|pie|doughnut/i,
                                                bar: /barChart|bar|column/i,
                                                line: /lineChart|line/i,
                                                area: /areaChart|area/i,
                                                scatter: /scatterChart|scatter|xy/i,
                                                bubble: /bubbleChart|bubble/i
                                            };

                                            for (const [chartType, pattern] of Object.entries(chartTypePatterns)) {
                                                if (pattern.test(xmlResult)) {
                                                    console.log('✅ 通过XML分析识别图表类型:', chartType);
                                                    detailedChartType.specificType = chartType;
                                                    detailedChartType.description = '图表 (' + chartType + ')';
                                                    detailedChartType.confidence = 0.95;
                                                    detailedChartType.detectionMethod = 'xml-pattern-analysis';
                                                    break;
                                                }
                                            }

                                            if (detailedChartType.specificType !== 'unknown') {
                                                break; // 找到了图表类型，退出循环
                                            }
                                        }
                                    } catch (xmlError) {
                                        console.log(`🚨 ${xmlMethod} 调用失败:`, xmlError.message);
                                    }
                                }
                            }
                        }

                        // 方法4: 尝试查看图表元素的所有可用方法
                        if (detailedChartType.specificType === 'unknown') {
                            console.log('🔍 分析图表元素的可用方法...');
                            const allMethods = [];
                            for (const prop in chartInfo.element) {
                                if (typeof chartInfo.element[prop] === 'function') {
                                    allMethods.push(prop);
                                }
                            }
                            console.log('📊 图表元素可用方法 (前20个):', allMethods.slice(0, 20));

                            // 尝试一些可能的图表类型相关方法
                            const typeMethods = allMethods.filter(method =>
                                method.toLowerCase().includes('type') ||
                                method.toLowerCase().includes('chart')
                            );
                            console.log('📊 类型相关方法:', typeMethods);

                            // 尝试调用这些方法
                            for (const typeMethod of typeMethods.slice(0, 3)) {
                                try {
                                    console.log(`🔍 尝试调用方法: ${typeMethod}`);
                                    const result = chartInfo.element[typeMethod]();
                                    console.log(`📊 ${typeMethod} 返回:`, result);
                                    if (result && typeof result === 'string') {
                                        detailedChartType.specificType = result;
                                        detailedChartType.description = '图表 (' + result + ')';
                                        detailedChartType.confidence = 0.9;
                                        console.log('✅ 通过方法' + typeMethod + '获得图表类型:', result);
                                        break;
                                    }
                                } catch (methodError) {
                                    console.log(`🚨 ${typeMethod} 调用失败:`, methodError.message);
                                }
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
        });
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

                    // 检查元素级图表的绑定数据
                    scanResults.elementLevelCharts.forEach((chartInfo) => {
                        // 简化的图表类型识别
                        const detailedChartType = {
                            category: chartInfo.elementType.includes('Chart') ? 'chart' :
                                     chartInfo.elementType.includes('Drawing') ? 'drawing' :
                                     chartInfo.elementType.includes('Shape') ? 'shape' : 'unknown',
                            specificType: 'unknown',
                            description: chartInfo.elementType.includes('Chart') ? '图表' :
                                        chartInfo.elementType.includes('Drawing') ? '绘图对象' :
                                        chartInfo.elementType.includes('Shape') ? '形状' : '未知类型',
                            properties: {},
                            confidence: 0.7
                        };

                        console.log('📊 图表类型识别:', detailedChartType);

                        const chartResult = {
                            chartIndex: chartInfo.index,
                            chartType: chartInfo.elementType,
                            detailedChartType: detailedChartType,
                            boundData: null,
                            hasBindingData: false
                        };

                        // 检查绑定数据
                        const dataResult = dataBinder.getBoundData(chartInfo.element, chartInfo.index);
                        Object.assign(chartResult, dataResult);

                        chartDetectionResults.push(chartResult);
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
     * 安全检查文档级图表数据（避免SDK内部错误）
     */
    _safeCheckDocumentLevelChartData(documentLevelCharts, chartDetectionResults) {
        documentLevelCharts.forEach((chartInfo) => {
            console.log('🔒 安全模式检查图表:', chartInfo.index);

            try {
                // 使用安全模式识别图表类型，避免触发SDK错误
                const detailedChartType = this._safeIdentifyChartType(chartInfo.element, chartInfo.elementType);
                console.log('📊 安全模式图表类型识别:', detailedChartType);

                // 生成图表唯一标识符时也要小心
                let uniqueId;
                try {
                    uniqueId = this.typeDetector.generateChartUniqueId(
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
                const dataResult = this.dataBinder.getBoundData(chartInfo.element, chartInfo.drawingIndex);
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
    }

    /**
     * 安全图表类型识别（避免触发SDK错误）
     */
    _safeIdentifyChartType(element, elementType) {
        try {
            // 首先尝试标准识别
            return this.typeDetector.identifyChartType(element, elementType);
        } catch (identifyError) {
            console.log('🚨 标准识别失败，使用安全模式:', identifyError.message);

            // 返回安全的基础识别结果
            return {
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

                    // 统计元素级图表
                    scanResults.elementLevelCharts.forEach((chartInfo) => {
                        summary.totalCharts++;
                        const dataResult = dataBinder.getBoundData(chartInfo.element, chartInfo.index);

                        if (dataResult.hasBindingData) {
                            summary.chartsWithData++;
                        }

                        summary.bindingSummary.push({
                            chartIndex: chartInfo.index,
                            chartType: chartInfo.elementType,
                            hasBindingData: dataResult.hasBindingData,
                            bindingPreview: dataResult.boundData?.bindingId || 'unknown',
                            source: 'element-level'
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

    /**
     * 处理文档级图表
     */
    _processDocumentLevelCharts(documentLevelCharts, bindingData, boundCharts, bindingResults) {
        documentLevelCharts.forEach((chartInfo) => {
            // 识别图表类型
            const detailedChartType = this.typeDetector.identifyChartType(chartInfo.element, chartInfo.elementType);
            console.log('📊 文档级图表类型识别结果:', detailedChartType);

            // 生成图表唯一标识符
            const uniqueId = this.typeDetector.generateChartUniqueId(
                chartInfo.element,
                chartInfo.elementType,
                chartInfo.drawingIndex
            );

            // 创建绑定信息
            const bindingInfo = {
                chartIndex: chartInfo.index,
                chartType: chartInfo.elementType,
                detailedChartType: detailedChartType,
                uniqueId: uniqueId, // 添加唯一标识符
                boundData: bindingData.data || {},
                bindingId: 'doc_chart_' + chartInfo.drawingIndex + '_' + Date.now(),
                boundAt: new Date().toISOString(),
                metadata: bindingData.metadata || {},
                isDocumentLevel: true,
                drawingIndex: chartInfo.drawingIndex
            };

            // 绑定数据
            const bindingResult = this.dataBinder.bindDataToChart(
                chartInfo.element,
                bindingInfo,
                chartInfo.drawingIndex
            );

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
        });
    }

    /**
     * 处理元素级图表
     */
    _processElementLevelCharts(elementLevelCharts, bindingData, boundCharts, bindingResults) {
        elementLevelCharts.forEach((chartInfo) => {
            // 识别图表类型
            const detailedChartType = this.typeDetector.identifyChartType(chartInfo.element, chartInfo.elementType);
            console.log('📊 图表类型识别结果:', detailedChartType);

            // 创建绑定信息
            const bindingInfo = {
                chartIndex: chartInfo.index,
                chartType: chartInfo.elementType,
                detailedChartType: detailedChartType,
                boundData: bindingData.data || {},
                bindingId: 'chart_' + chartInfo.index + '_' + Date.now(),
                boundAt: new Date().toISOString(),
                metadata: bindingData.metadata || {},
                debugInfo: chartInfo.debugInfo
            };

            // 绑定数据
            const bindingResult = this.dataBinder.bindDataToChart(
                chartInfo.element,
                bindingInfo,
                chartInfo.index
            );

            // 合并绑定结果
            Object.assign(bindingInfo, bindingResult);

            boundCharts.push(chartInfo);
            bindingResults.push(bindingInfo);
        });
    }

    /**
     * 检查文档级图表数据
     */
    _checkDocumentLevelChartData(documentLevelCharts, chartDetectionResults) {
        documentLevelCharts.forEach((chartInfo) => {
            // 识别图表类型
            const detailedChartType = this.typeDetector.identifyChartType(chartInfo.element, chartInfo.elementType);
            console.log('📊 文档级图表类型识别:', detailedChartType);

            // 生成图表唯一标识符（点击时重新生成，用于匹配）
            const uniqueId = this.typeDetector.generateChartUniqueId(
                chartInfo.element,
                chartInfo.elementType,
                chartInfo.drawingIndex
            );

            const chartResult = {
                chartIndex: chartInfo.index,
                chartType: chartInfo.elementType,
                detailedChartType: detailedChartType,
                uniqueId: uniqueId, // 添加唯一标识符
                boundData: null,
                hasBindingData: false,
                isDocumentLevel: true,
                drawingIndex: chartInfo.drawingIndex
            };

            // 检查绑定数据
            const dataResult = this.dataBinder.getBoundData(chartInfo.element, chartInfo.drawingIndex);
            Object.assign(chartResult, dataResult);

            // 如果找到绑定数据，尝试匹配唯一标识符
            if (chartResult.hasBindingData && chartResult.boundData) {
                console.log('🔍 检查图表数据匹配:', {
                    currentId: uniqueId,
                    boundId: chartResult.boundData.uniqueId,
                    chartType: detailedChartType.description
                });
            }

            chartDetectionResults.push(chartResult);
        });
    }

    /**
     * 检查元素级图表数据
     */
    _checkElementLevelChartData(elementLevelCharts, chartDetectionResults) {
        elementLevelCharts.forEach((chartInfo) => {
            // 简化的图表类型识别
            const detailedChartType = {
                category: chartInfo.elementType.includes('Chart') ? 'chart' :
                         chartInfo.elementType.includes('Drawing') ? 'drawing' :
                         chartInfo.elementType.includes('Shape') ? 'shape' : 'unknown',
                specificType: 'unknown',
                description: chartInfo.elementType.includes('Chart') ? '图表' :
                            chartInfo.elementType.includes('Drawing') ? '绘图对象' :
                            chartInfo.elementType.includes('Shape') ? '形状' : '未知类型',
                properties: {},
                confidence: 0.7
            };

            console.log('📊 图表类型识别:', detailedChartType);

            const chartResult = {
                chartIndex: chartInfo.index,
                chartType: chartInfo.elementType,
                detailedChartType: detailedChartType,
                boundData: null,
                hasBindingData: false
            };

            // 检查绑定数据
            const dataResult = this.dataBinder.getBoundData(chartInfo.element, chartInfo.index);
            Object.assign(chartResult, dataResult);

            chartDetectionResults.push(chartResult);
        });
    }

    /**
     * 查找被点击的图表
     */
    _findClickedChart(chartDetectionResults, hasSelection) {
        if (chartDetectionResults.length === 0) return null;

        // 简单策略：如果有选区，取最后一个有数据的图表
        for (let i = chartDetectionResults.length - 1; i >= 0; i--) {
            if (chartDetectionResults[i].hasBindingData) {
                return chartDetectionResults[i];
            }
        }

        // 如果没有找到有数据的图表，取最后一个
        return chartDetectionResults[chartDetectionResults.length - 1];
    }
}