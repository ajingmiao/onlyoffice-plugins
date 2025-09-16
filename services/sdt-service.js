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
    console.log('🔍 detectActiveSdt 开始执行...');
    return new Promise((resolve) => {
      this.editor.runInDoc(function () {
        console.log('📄 进入 runInDoc 上下文...');
        var doc = Api.GetDocument();
        var range = doc.GetRangeBySelect();
        console.log('📋 获取选区:', range);

        if (!range) {
          console.log('❌ 没有选区，返回 null');
          return null;
        }

        console.log('🔍 开始扫描内容控件...');
        var ctrls = doc.GetAllContentControls();
        console.log('📊 找到', ctrls.length, '个内容控件');

        for (var i = 0; i < ctrls.length; i++) {
          var c = ctrls[i];
          console.log('🔍 检查内容控件', i);

          var hit = false;

          // 方法1: range.IsInContentControl
          try {
            if (typeof range.IsInContentControl === 'function') {
              hit = range.IsInContentControl(c);
              console.log('方法1结果:', hit);
            } else {
              console.log('方法1不存在');
            }
          } catch (e) {
            console.log('方法1异常:', e.message);
          }

          // 方法2: c.IsRangeIn
          if (!hit) {
            try {
              if (typeof c.IsRangeIn === 'function') {
                hit = c.IsRangeIn(range);
                console.log('方法2结果:', hit);
              } else {
                console.log('方法2不存在');
              }
            } catch (e) {
              console.log('方法2异常:', e.message);
            }
          }

          // 方法3: c.IsSelected
          if (!hit) {
            try {
              if (typeof c.IsSelected === 'function') {
                hit = c.IsSelected();
                console.log('方法3结果:', hit);
              } else {
                console.log('方法3不存在');
              }
            } catch (e) {
              console.log('方法3异常:', e.message);
            }
          }

          console.log('内容控件', i, '最终结果:', hit);

          if (hit) {
            var result = {
              tag: c.GetTag ? c.GetTag() : "",
              alias: c.GetAlias ? c.GetAlias() : ""
            };
            console.log('✅ 匹配成功:', result);
            return result;
          }
        }
        console.log('❌ 没有找到匹配的内容控件');
        return null;
      }, { async: false, cb: (res) => {
        console.log('🔄 runInDoc 回调被调用，结果:', res);
        resolve(res);
      } });
    });
  }
}
