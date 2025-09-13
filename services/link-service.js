// services/link-service.js
export class LinkService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 处理链接点击事件 - 使用sdt-service.js的exact模式
    async handleLinkClick() {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                var doc = Api.GetDocument();
                var range = doc.GetRangeBySelect();

                console.log('=== Handling link click (sdt-service pattern) ===');
                console.log('Current range:', range);

                if (!range) {
                    console.log('No range selected');
                    return null;
                }

                // 完全按照sdt-service.js的方式检测控件
                var ctrls = doc.GetAllContentControls();
                console.log('Found content controls:', ctrls.length);

                for (var i = 0; i < ctrls.length; i++) {
                    var ctrl = ctrls[i];
                    console.log('Checking control', i, 'with tag:', ctrl.GetTag ? ctrl.GetTag() : 'no tag');

                    // 使用与sdt-service.js完全相同的hit检测逻辑
                    var hit =
                        (typeof range.IsInContentControl === 'function' && range.IsInContentControl(ctrl)) ||
                        (typeof ctrl.IsRangeIn === 'function' && ctrl.IsRangeIn(range)) ||
                        (typeof ctrl.IsSelected === 'function' && ctrl.IsSelected());

                    console.log('Control hit status:', hit);

                    if (hit) {
                        // 检查是否是我们的链接控件
                        var tag = ctrl.GetTag ? ctrl.GetTag() : '';
                        if (tag && tag.startsWith('link-data:')) {
                            console.log('Found link control!');
                            try {
                                var jsonStr = tag.substring('link-data:'.length);
                                var jsonData = jsonStr ? JSON.parse(jsonStr) : {};

                                // 返回控件信息和数据
                                var result = {
                                    controlId: ctrl.GetId ? ctrl.GetId() : 'unknown',
                                    controlTitle: ctrl.GetAlias ? ctrl.GetAlias() : '',
                                    tag: tag,
                                    data: jsonData
                                };

                                console.log('Link click detected, returning data:', result);
                                return result;
                            } catch (e) {
                                console.log('Error parsing link data:', e);
                                return null;
                            }
                        }
                    }
                }

                console.log('No link control found at current position');
                return null;

            }, { async: false, cb: (res) => resolve(res) });
        });
    }

    insertLinkInline(opts) {
        const options = opts || {};

        // 设置 Asc.scope 让文档内脚本能访问
        window.Asc = window.Asc || {};
        window.Asc.scope = {
            text: options.text || '点击这里',
            url: options.url || '',
            style: options.style || {},
            json: options.json || ''
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            // —— 只读 Asc.scope —— //
            var text = (Asc.scope && Asc.scope.text) || '点击这里';
            var url = (Asc.scope && Asc.scope.url) || '';
            var style = (Asc.scope && Asc.scope.style) || {};
            var bold = (style.bold !== undefined) ? !!style.bold : true;
            var underline = (style.underline !== undefined) ? !!style.underline : true;
            var json = (Asc.scope && Asc.scope.json) || '';

            function toRgb(c) {
                if (!c) return [25, 118, 210];                       // 默认 #1976D2
                if (Array.isArray(c) && c.length === 3) return c;  // [r,g,b]
                if (typeof c === 'string' && /^#?[0-9a-fA-F]{6}$/.test(c)) {
                    var hex = c[0] === '#' ? c.slice(1) : c;
                    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
                }
                return [25, 118, 210];
            }
            var rgb = toRgb(style.color);

            console.log('=== Link insertion (using sdt-service pattern) ===');
            console.log('Parameters:', { text: text, url: url, json: json, style: style });
            console.log('RGB color:', rgb);
            console.log('Bold:', bold, 'Underline:', underline);

            var jsonData = json ? JSON.stringify(json) : '';

            try {
                if (!url || !String(url).trim()) {
                    // 空URL - 完全参考sdt-service.js的成功模式
                    var sdt = doc.AddComboBoxContentControl();

                    // 设置Tag - 用于数据绑定和识别
                    sdt.SetTag('link-data:' + jsonData);
                    if (sdt.SetAlias) {
                        sdt.SetAlias('可点击链接');
                    }

                    // 清空默认内容
                    var content = sdt.GetContent();
                    content.RemoveAllElements();
                    var p = content.GetElement(0);

                    // 添加带链接样式的文本
                    if (p) {
                        var run = Api.CreateRun();
                        run.AddText(text);

                        // 确保样式设置成功
                        console.log('Setting text styles...');
                        if (run.SetBold) {
                            var boldResult = run.SetBold(bold);
                            console.log('Set bold result:', boldResult, 'value:', bold);
                        }
                        if (run.SetUnderline) {
                            var underlineResult = run.SetUnderline(underline);
                            console.log('Set underline result:', underlineResult, 'value:', underline);
                        }
                        if (run.SetColor) {
                            var colorResult = run.SetColor(rgb[0], rgb[1], rgb[2]);
                            console.log('Set color result:', colorResult, 'RGB:', rgb[0], rgb[1], rgb[2]);
                        }

                        console.log('Text run created with text:', text);
                        p.AddElement(run);
                    }

                    console.log('Added content control with JSON data binding at cursor position');
                } else {
                    // 有URL - 使用真实超链接
                    var sdt = doc.AddComboBoxContentControl();
                    sdt.SetTag('link-data:' + jsonData);
                    if (sdt.SetAlias) {
                        sdt.SetAlias('超链接: ' + url);
                    }

                    var content = sdt.GetContent();
                    content.RemoveAllElements();
                    var p = content.GetElement(0);

                    if (p) {
                        // 创建超链接元素
                        var hyperlink = Api.CreateHyperlink(url, text);
                        if (hyperlink.SetBold) hyperlink.SetBold(bold);
                        if (hyperlink.SetUnderline) hyperlink.SetUnderline(underline);
                        if (hyperlink.SetColor) hyperlink.SetColor(rgb[0], rgb[1], rgb[2]);
                        p.AddElement(hyperlink);
                    }

                    console.log('Added hyperlink content control at cursor position');
                }

                console.log('Link insertion completed successfully');

            } catch (e) {
                console.log('Link insertion failed:', e);
                console.log('Error stack:', e.stack);
            }
        });
    }
}
