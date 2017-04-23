# wiz 快速搜索插件

wiz（为知笔记）快速搜索插件，仅支持 windows。

![](示例图.jpg)

## 使用说明

### 安装

以插件形式安装：在 [Releases](https://github.com/ywzhaiqi/wizQuickSearch/releases) 处下载并安装 `wizQuickSearch.wizplugin`，然后在 wiz 菜单 `工具->wiz快速搜索` 点击启动。

也可直接下载[源文件](https://github.com/ywzhaiqi/wizQuickSearch.git)，放在独立目录，点击 `wizQuickSearch.exe` 启动。

注意：wiz 关闭后会自动退出，重新打开 wiz 也要重新启动本工具。

### 使用

启动后，使用按键 `#q`（win键 + Q键） 打开搜索框搜索（支持拼音首字母）。
 - 直接搜索文档标题+标签
 - `@` 搜索文件夹
 - `#` 搜索标签
 - `*` 使用 acc 搜索（将被移除）

搜索框输入完成后：
 - 直接按 `Up`、`Down` 切换选中条目
 - 直接按 `Left`、`Right` 在输入框移动
 - 直接按 `Enter` 打开选中的文档
 - 直接按 `Ctrl + Enter` 发送搜索到的列表到wiz

最后按 `esc` 键隐藏搜索界面

### 缘由

之前使用 wiz 自带搜索框，经常需要在全文搜索、标题搜索之间切换，且搜索后，需要再次点击才能打开，不符合直达的要求，故写了此插件。