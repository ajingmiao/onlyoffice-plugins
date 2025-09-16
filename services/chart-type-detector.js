// services/chart-type-detector.js
export class ChartTypeDetector {
    constructor() {
        // 图表类型模式匹配
        this.chartTypePatterns = {
            pie: /pie|doughnut/i,
            bar: /bar|column/i,
            line: /line/i,
            area: /area/i,
            scatter: /scatter|xy/i,
            bubble: /bubble/i
        };

        // 常见的OOXML获取方法
        this.xmlMethods = ['GetOOXML', 'GetXML', 'ToXML', 'GetDocumentXML', 'GetChart', 'GetChartSpace'];
    }

    /**
     * 识别图表类型
     * @param {Object} element - 图表元素对象
     * @param {string} elementType - 元素类型字符串
     * @returns {Object} 图表类型识别结果
     */
    identifyChartType(element, elementType) {
        const result = {
            category: 'unknown',
            specificType: 'unknown',
            description: '未知类型',
            properties: {},
            confidence: 0
        };

        try {
            // 基于元素类型的初步分类
            if (elementType.includes('Chart') || elementType === 'chart') {
                result.category = 'chart';
                result.description = '图表';
                result.confidence = 0.8;

                // 尝试获取图表的具体类型
                this._detectFromChartObject(element, result);

                // 如果无法直接获取Chart对象，使用深度分析
                if (result.specificType === 'unknown') {
                    this._detectFromDeepAnalysis(element, result);
                }
            } else {
                this._classifyNonChartElements(element, elementType, result);
            }

            // 获取通用属性
            this._extractCommonProperties(element, result);

        } catch (identifyError) {
            console.log('图表类型识别出错:', identifyError);
            result.error = identifyError.message;
        }

        return result;
    }

    /**
     * 从Chart对象检测类型（最直接的方法）
     */
    _detectFromChartObject(element, result) {
        if (typeof element.GetChart === 'function') {
            try {
                const chart = element.GetChart();
                if (chart) {
                    result.confidence = 0.9;

                    // 尝试获取图表类型
                    if (typeof chart.GetChartType === 'function') {
                        const chartType = chart.GetChartType();
                        if (chartType) {
                            result.specificType = chartType;
                            result.confidence = 1.0;
                            result.description = this._getChartTypeDescription(chartType);
                            result.properties.detectionMethod = 'direct-chart-api';
                        }
                    }

                    // 尝试获取系列数量
                    if (typeof chart.GetSeriesCount === 'function') {
                        result.properties.seriesCount = chart.GetSeriesCount();
                    }
                }
            } catch (error) {
                console.log('获取Chart对象失败:', error.message);
            }
        }
    }

    /**
     * 深度分析方法（OOXML、方法调用等）
     */
    _detectFromDeepAnalysis(element, result) {
        result.specificType = 'chart_object';
        result.description = '图表对象';
        result.confidence = 0.9;

        // 获取基本属性
        this._extractBasicProperties(element, result);

        // 1. 分析可用方法
        const methodAnalysis = this._analyzeAvailableMethods(element);
        result.properties = { ...result.properties, ...methodAnalysis };

        // 2. 尝试OOXML分析
        const xmlResult = this._analyzeXMLContent(element);
        if (xmlResult.success) {
            result.specificType = xmlResult.chartType;
            result.description = this._getChartTypeDescription(xmlResult.chartType);
            result.confidence = 0.95;
            result.properties.detectionMethod = 'xml-pattern-analysis';
            result.properties.xmlMethod = xmlResult.method;
            result.properties.xmlSample = xmlResult.sample;
            return; // 找到了，直接返回
        }

        // 3. 尝试类型方法调用
        const typeResult = this._analyzeTypeMethods(element, methodAnalysis.typeMethods || []);
        if (typeResult.success) {
            result.specificType = typeResult.chartType;
            result.description = this._getChartTypeDescription(typeResult.chartType);
            result.confidence = 0.9;
            result.properties.detectionMethod = typeResult.method;
            return;
        }

        // 4. 尝试图表方法调用
        const chartResult = this._analyzeChartMethods(element, methodAnalysis.chartMethods || []);
        if (chartResult.success) {
            result.specificType = chartResult.chartType;
            result.description = this._getChartTypeDescription(chartResult.chartType);
            result.confidence = 1.0;
            result.properties.detectionMethod = chartResult.method;
        }

        // 5. 记录尺寸信息（仅作参考）
        if (result.properties.width && result.properties.height) {
            const aspectRatio = result.properties.width / result.properties.height;
            result.properties.aspectRatio = Math.round(aspectRatio * 100) / 100;
            result.properties.dimensionNote = '尺寸仅供参考，类型识别基于内容分析';
        }
    }

    /**
     * 分析可用方法
     */
    _analyzeAvailableMethods(element) {
        const allProps = [];
        const typeMethods = [];
        const chartMethods = [];

        for (const prop in element) {
            if (typeof element[prop] === 'function') {
                allProps.push(prop);
                if (prop.toLowerCase().includes('type')) {
                    typeMethods.push(prop);
                }
                if (prop.toLowerCase().includes('chart') ||
                    prop.toLowerCase().includes('series') ||
                    prop.toLowerCase().includes('data') ||
                    prop.toLowerCase().includes('ooxml') ||
                    prop.toLowerCase().includes('xml')) {
                    chartMethods.push(prop);
                }
            }
        }

        console.log('🔍 图表对象所有方法 (前20个):', allProps.slice(0, 20));
        console.log('🔍 类型相关方法:', typeMethods);
        console.log('🔍 图表相关方法:', chartMethods);

        return {
            availableMethods: allProps.slice(0, 20),
            typeMethods,
            chartMethods
        };
    }

    /**
     * 分析XML内容
     */
    _analyzeXMLContent(element) {
        for (const methodName of this.xmlMethods) {
            if (typeof element[methodName] === 'function') {
                try {
                    console.log('🔍 尝试调用方法:', methodName);
                    const xmlResult = element[methodName]();
                    if (xmlResult && typeof xmlResult === 'string' && xmlResult.length > 0) {
                        console.log('✅ 成功获取XML内容，方法:', methodName);
                        console.log('📄 XML内容长度:', xmlResult.length);

                        // 查找图表类型模式
                        for (const [chartType, pattern] of Object.entries(this.chartTypePatterns)) {
                            if (pattern.test(xmlResult)) {
                                console.log('✅ 通过XML分析识别图表类型:', chartType);
                                return {
                                    success: true,
                                    chartType,
                                    method: methodName,
                                    sample: xmlResult.substring(0, 200) + '...'
                                };
                            }
                        }
                    }
                } catch (xmlError) {
                    console.log('调用' + methodName + '失败:', xmlError.message);
                }
            }
        }
        return { success: false };
    }

    /**
     * 分析类型方法
     */
    _analyzeTypeMethods(element, typeMethods) {
        for (const typeMethod of typeMethods) {
            try {
                console.log('🔍 尝试调用类型方法:', typeMethod);
                const typeResult = element[typeMethod]();
                if (typeResult !== undefined && typeResult !== null) {
                    console.log('📊 ' + typeMethod + ' 返回:', typeResult);

                    if (typeof typeResult === 'string' || typeof typeResult === 'number') {
                        const typeStr = String(typeResult).toLowerCase();
                        const chartType = this._parseTypeString(typeStr);
                        if (chartType) {
                            return {
                                success: true,
                                chartType,
                                method: 'method-' + typeMethod
                            };
                        }
                    }
                }
            } catch (methodError) {
                console.log('调用' + typeMethod + '失败:', methodError.message);
            }
        }
        return { success: false };
    }

    /**
     * 分析图表方法
     */
    _analyzeChartMethods(element, chartMethods) {
        for (const chartMethod of chartMethods) {
            try {
                console.log('🔍 尝试调用图表方法:', chartMethod);
                const chartResult = element[chartMethod]();
                if (chartResult && typeof chartResult === 'object') {
                    console.log('📊 ' + chartMethod + ' 返回对象:', chartResult);

                    // 如果返回的是图表对象，尝试进一步分析
                    if (chartResult && typeof chartResult.GetChartType === 'function') {
                        try {
                            const actualChartType = chartResult.GetChartType();
                            if (actualChartType) {
                                console.log('✅ 通过' + chartMethod + '获得准确图表类型:', actualChartType);
                                return {
                                    success: true,
                                    chartType: actualChartType,
                                    method: chartMethod + '.GetChartType'
                                };
                            }
                        } catch (innerError) {
                            console.log('调用GetChartType失败:', innerError.message);
                        }
                    }
                }
            } catch (methodError) {
                console.log('调用' + chartMethod + '失败:', methodError.message);
            }
        }
        return { success: false };
    }

    /**
     * 获取基本属性
     */
    _extractBasicProperties(element, result) {
        // 尝试获取尺寸
        if (typeof element.GetWidth === 'function' && typeof element.GetHeight === 'function') {
            result.properties.width = element.GetWidth();
            result.properties.height = element.GetHeight();
        }

        // 检查是否有标题
        if (typeof element.GetTitle === 'function') {
            const title = element.GetTitle();
            if (title) {
                result.properties.title = title;
            }
        }
    }

    /**
     * 分类非图表元素
     */
    _classifyNonChartElements(element, elementType, result) {
        if (elementType.includes('Drawing')) {
            result.category = 'drawing';
            result.description = '绘图对象';
            result.confidence = 0.7;
        } else if (elementType.includes('Shape')) {
            result.category = 'shape';
            result.description = '形状';
            result.confidence = 0.8;
            this._detectShapeType(element, result);
        } else if (elementType.includes('Image')) {
            result.category = 'image';
            result.description = '图像';
            result.confidence = 0.9;
        } else if (elementType.includes('Ole')) {
            result.category = 'ole';
            result.description = 'OLE对象';
            result.confidence = 0.8;
        } else if (elementType === 'GraphicFrame') {
            result.category = 'graphic_frame';
            result.description = '图形框架';
            result.confidence = 0.7;
        }
    }

    /**
     * 检测形状类型
     */
    _detectShapeType(element, result) {
        if (typeof element.GetShapeType === 'function') {
            try {
                const shapeType = element.GetShapeType();
                if (shapeType) {
                    result.specificType = shapeType;
                    result.description = '形状 (' + shapeType + ')';
                    result.confidence = 0.9;
                }
            } catch (error) {
                console.log('获取形状类型失败:', error.message);
            }
        }
    }

    /**
     * 提取通用属性
     */
    _extractCommonProperties(element, result) {
        // 获取通用尺寸属性
        if (typeof element.GetWidth === 'function' && typeof element.GetHeight === 'function') {
            result.properties.width = element.GetWidth();
            result.properties.height = element.GetHeight();
        }

        // 检查是否有文本内容
        if (typeof element.GetText === 'function') {
            try {
                const text = element.GetText();
                if (text && text.trim()) {
                    result.properties.hasText = true;
                    result.properties.textLength = text.length;
                }
            } catch (error) {
                // 忽略文本获取错误
            }
        }
    }

    /**
     * 从类型字符串解析图表类型
     */
    _parseTypeString(typeStr) {
        if (typeStr.includes('pie') || typeStr.includes('doughnut')) {
            return 'pie';
        } else if (typeStr.includes('bar') || typeStr.includes('column')) {
            return 'bar';
        } else if (typeStr.includes('line')) {
            return 'line';
        } else if (typeStr.includes('area')) {
            return 'area';
        } else if (typeStr.includes('scatter')) {
            return 'scatter';
        } else if (typeStr.includes('bubble')) {
            return 'bubble';
        }
        return null;
    }

    /**
     * 获取图表类型的中文描述
     */
    _getChartTypeDescription(chartType) {
        const descriptions = {
            pie: '饼图',
            bar: '柱状图',
            column: '柱状图',
            line: '折线图',
            area: '面积图',
            scatter: '散点图',
            bubble: '气泡图',
            doughnut: '环形图'
        };

        return descriptions[chartType] || `图表 (${chartType})`;
    }
}