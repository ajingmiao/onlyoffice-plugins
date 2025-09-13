// services/link-service.js
export class LinkService {
    constructor(editorService) { this.editor = editorService; }

    insertLinkInline(opts) {
        const options = opts || {};
        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            // —— 只读 Asc.scope —— //
            var text = (Asc.scope && Asc.scope.text) || '点击这里';
            var url = (Asc.scope && Asc.scope.url) || '';
            var style = (Asc.scope && Asc.scope.style) || {};
            var bold = (style.bold !== undefined) ? !!style.bold : true;
            var underline = (style.underline !== undefined) ? !!style.underline : true;

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

            // 原始是否为空（用于占位标记）
            var rawEmpty = !(url && String(url).trim());
            var finalUrl = rawEmpty ? 'about:blank' : url;

            // Run + 样式
            var run = Api.CreateRun();
            run.AddText(text);
            run.SetBold && run.SetBold(bold);
            run.SetUnderline && run.SetUnderline(underline);
            run.SetColor && run.SetColor(rgb[0], rgb[1], rgb[2]);

            // 超链接（有 API 用超链，缺失则退化为“像链接”的文本）
            var hyperlink = null;
            if (typeof Api.CreateHyperlink === 'function') {
                try {
                    hyperlink = Api.CreateHyperlink(finalUrl, text);
                    hyperlink.SetBold && hyperlink.SetBold(bold);
                    hyperlink.SetUnderline && hyperlink.SetUnderline(underline);
                    hyperlink.SetColor && hyperlink.SetColor(rgb[0], rgb[1], rgb[2]);
                } catch (e) {
                    try { hyperlink = Api.CreateHyperlink(finalUrl, run); } catch (e2) { hyperlink = null; }
                }
            }

            // 插入位置
            var range = doc.GetRangeBySelect();

            // 1) 优先：让文档自己在光标处插入
            if (typeof doc.AddHyperlink === 'function') {
                // 直接在光标处插入链接（返回的对象通常就是超链）
                var hl = doc.AddHyperlink(finalUrl, text);

                // 如果返回对象支持样式，则按需设样式
                if (hl) {
                    if (hl.SetBold) hl.SetBold(bold);
                    if (hl.SetUnderline) hl.SetUnderline(underline);
                    if (hl.SetColor) hl.SetColor(rgb[0], rgb[1], rgb[2]);
                }
            } else if (hyperlink) {
                // 2) 次选：我们已创建好 hyperlink 对象 → 退回段末插入
                var p = (range && range.GetParagraph) ? range.GetParagraph() : doc.GetElement(0);
                if (!p) { p = Api.CreateParagraph(); doc.Push(p); }
                p.AddElement(hyperlink);
            } else {
                // 3) 最差：没有超链 API，就插一个“像链接”的 run（蓝+粗+下划线）
                var p2 = (range && range.GetParagraph) ? range.GetParagraph() : doc.GetElement(0);
                if (!p2) { p2 = Api.CreateParagraph(); doc.Push(p2); }
                p2.AddElement(run);
            }

            // === 空链接占位打标（可选） ===
            if (!(url && String(url).trim())) {
                var sdt = Api.CreateInlineLvlSdt();
                var mark = Api.CreateRun(); mark.AddText('');
                sdt.AddElement(mark, 0);
                sdt.SetTag('link:pending');
                sdt.SetAlias('空链接占位');
                // 占位控件也尽量插到“当前段落”
                var p3 = (range && range.GetParagraph) ? range.GetParagraph() : doc.GetElement(0);
                if (!p3) { p3 = Api.CreateParagraph(); doc.Push(p3); }
                p3.AddInlineLvlSdt(sdt);
            }
        }, {
            async: false,
            // ★ 在调用前把需要的参数放进 Asc.scope
            scope: {
                text: options.text,
                url: options.url,
                style: { bold: options.bold, underline: options.underline, color: options.color }
            }
        });
    }
}
