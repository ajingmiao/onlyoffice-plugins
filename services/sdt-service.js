import { TAG_PREFIX } from '../core/constants.js';

export class SdtService {
  constructor(editorService) {
    this.editor = editorService;
  }

  /**
   * 插入一个行内内容控件（示例用组合框；你也可切换 DropDownList）
   * tagKey: 业务标识（会被加上 TAG_PREFIX）
   * title:  标题
   * placeholder: 占位显示文本
   */
  insertInlineSdt({ tagKey, title, placeholder }) {
    return this.editor.runInDoc(function () {
      var doc = Api.GetDocument();

      // 组合框控件（也可用 doc.AddDropDownListContentControl()）
      var sdt = doc.AddComboBoxContentControl();

      // 显示内容（清空默认内容后，加入占位文字）
      var content = sdt.GetContent();
      content.RemoveAllElements();
      var p = content.GetElement(0);
      p.AddText(placeholder || '{{placeholder}}');

      // Tag / Title（Tag 用于数据绑定）
      sdt.SetTag((TAG_PREFIX + tagKey) || 'bind:unknown');
      if (title) sdt.SetTitle(title);
    }, { async: false });
  }

  /** 在当前 Range/选择内探测是否命中某个 SDT，并返回其 tag/alias */
  detectActiveSdt() {
    return new Promise((resolve) => {
      this.editor.runInDoc(function () {
        var doc = Api.GetDocument();
        var range = doc.GetRangeBySelect();
        if (!range) return null;

        var ctrls = doc.GetAllContentControls();
        for (var i = 0; i < ctrls.length; i++) {
          var c = ctrls[i];
          var hit =
            (typeof range.IsInContentControl === 'function' && range.IsInContentControl(c)) ||
            (typeof c.IsRangeIn === 'function' && c.IsRangeIn(range)) ||
            (typeof c.IsSelected === 'function' && c.IsSelected());
          if (hit) {
            return {
              tag: c.GetTag ? c.GetTag() : "",
              alias: c.GetAlias ? c.GetAlias() : ""
            };
          }
        }
        return null;
      }, { async: false, cb: (res) => resolve(res) });
    });
  }
}
