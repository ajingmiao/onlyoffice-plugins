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
        return this.editor.runInDoc(() => {
            const doc = Api.GetDocument();
            console.log('=== 图表数据绑定开始 ===');

            try {
                // 创建测试绑定数据
                const bindingData = this.dataBinder.createTestBindingData();
                console.log('使用测试绑定数据:', bindingData);

                // 扫描文档中的图表
                const scanResults = this.chartDetector.scanDocument(doc);
                const boundCharts = [];
                const bindingResults = [];

                // 处理文档级图表
                this._processDocumentLevelCharts(
                    scanResults.documentLevelCharts,
                    bindingData,
                    boundCharts,
                    bindingResults
                );

                // 处理元素级图表
                this._processElementLevelCharts(
                    scanResults.elementLevelCharts,
                    bindingData,
                    boundCharts,
                    bindingResults
                );

                // 如果没有找到图表，提供指导
                if (boundCharts.length === 0) {
                    this.chartDetector.provideDetectionGuidance();
                }

                return {
                    success: true,
                    message: boundCharts.length > 0 ?
                        ('成功绑定 ' + boundCharts.length + ' 个图表') :
                        ('文档扫描完成，但未发现图表元素。扫描了 ' + scanResults.scanStatistics.totalElementsScanned + ' 个元素。'),
                    data: {
                        chartsFound: boundCharts.length,
                        bindingResults: bindingResults,
                        boundData: bindingData,
                        bindingMethod: 'content-control-marker',
                        scanStatistics: scanResults.scanStatistics,
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
        return new Promise((resolve) => {
            this.editor.runInDoc(() => {
                const doc = Api.GetDocument();
                console.log('=== 图表点击检测 ===');

                try {
                    // 获取当前选区信息
                    const selectionInfo = this.chartDetector.getCurrentSelection(doc);

                    // 扫描所有图表
                    const scanResults = this.chartDetector.scanDocument(doc);
                    const chartDetectionResults = [];

                    // 检查文档级图表的绑定数据
                    this._checkDocumentLevelChartData(scanResults.documentLevelCharts, chartDetectionResults);

                    // 检查元素级图表的绑定数据
                    this._checkElementLevelChartData(scanResults.elementLevelCharts, chartDetectionResults);

                    // 查找被点击的图表
                    const clickedChart = this._findClickedChart(chartDetectionResults, selectionInfo.hasSelection);

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
        return new Promise((resolve) => {
            this.editor.runInDoc(() => {
                const doc = Api.GetDocument();
                console.log('=== 获取图表绑定摘要 ===');

                try {
                    const scanResults = this.chartDetector.scanDocument(doc);
                    const summary = {
                        totalCharts: 0,
                        chartsWithData: 0,
                        bindingSummary: []
                    };

                    // 统计文档级图表
                    scanResults.documentLevelCharts.forEach((chartInfo, index) => {
                        summary.totalCharts++;
                        const dataResult = this.dataBinder.getBoundData(chartInfo.element, chartInfo.drawingIndex);

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
                        const dataResult = this.dataBinder.getBoundData(chartInfo.element, chartInfo.index);

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

            // 创建绑定信息
            const bindingInfo = {
                chartIndex: chartInfo.index,
                chartType: chartInfo.elementType,
                detailedChartType: detailedChartType,
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

            const chartResult = {
                chartIndex: chartInfo.index,
                chartType: chartInfo.elementType,
                detailedChartType: detailedChartType,
                boundData: null,
                hasBindingData: false,
                isDocumentLevel: true,
                drawingIndex: chartInfo.drawingIndex
            };

            // 检查绑定数据
            const dataResult = this.dataBinder.getBoundData(chartInfo.element, chartInfo.drawingIndex);
            Object.assign(chartResult, dataResult);

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