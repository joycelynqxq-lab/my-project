# 星露谷帮助栏目

一个纯静态的星露谷物语 Mod / 玩法帮助手册，用浏览器直接打开 `index.html` 即可查看，无需安装任何依赖。

## 目录结构

```
index.html          首页（分类入口）
mod-basic.html       Mod 基础操作（MOD菜单、拖拉机、CJB作弊菜单、自动化等）
furniture.html       家具与装饰（快乐家具设计师工具、家具获取方式）
appearance.html      外观与服饰（FS服饰美化、桃子体设置）
market-town.html     集镇经营攻略（Market Town Mod 完整攻略）
assets/
  css/style.css      全局样式
  js/main.js         导航折叠 / 图片灯箱交互
  images/            各页面用到的截图
```

## 本地预览

直接双击 `index.html` 用浏览器打开，或在项目目录下起一个静态服务器，例如：

```bash
python3 -m http.server 8000
```

然后访问 `http://localhost:8000`。

## 更新内容

内容会随着资料补充持续更新，新增分类时在 `index.html` 的分类卡片区添加入口，并新建对应页面即可。
