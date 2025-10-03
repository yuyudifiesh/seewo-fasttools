# FastToolsForOldSeewo
为希沃 windows7 系统和低性能的一体机提供的便携授课文件查看和处理工具

## 目录
1. [启发](#启发)
2. [PPT 压缩工具](#ppt-压缩工具)
3. [PDF 阅读工具](#pdf-阅读工具)

## 启发
初中的电脑用的是鸿合的 win10 系统，多装几个软件（微信/B站）同时运行在后台，再打开较大的 PPT 文件时候总是异常卡顿。

高中用的是希沃的 win10 系统，~~可能是学校的采购问题，也可能……（我没说过）~~，反正，在未激活的希沃一体机上打开超过 60 MB 的 PPT 并使用批量激活的盗版 Microsoft Powerpoint 读取时经常出现窗口未响应的卡死情况。总之，我做了一些小工具，希望能改善这些问题qwq。

<br>

## PPT 压缩工具
您的 PPT 做得太好啦，就是~~稍微~~有点大。

![image](./view/pptxy.gif)

此工具无需联网，将在不破坏任何动画、切换、母版、字体的前提下，对图片进行二次压缩，并剔除未使用的母版/版式/字体关系，使文件体积减小 **≥30%**。[【Releases】](https://github.com/yuyudifiesh/seewo-fasttools/releases)

| 适用系统 | 文件名 | 版本 | 更新时间 | 下载 |
| :---: | :---: | :---: | :---: | :---: |
| Windows10/11 | pptxya_20251003_1.exe | 1.0.0 | 2025 年 10 月 03 日 | [【下载】](./soft/pptxya_20251003_1.exe) |
| Windows7 | pptxya_win7.exe | 1.0.0-AlphaTest | 2025 年 10 月 03 日 | [【下载】](./soft/pptxya_win7.exe)  |

> 关于 Windows7 预览版本的说明
>
> 对于 Windows7 的版本，在代码上进行了深度精简优化，并移除了部分现代化的 UI 设计，在处理时仅使用单线程压缩，在确保运行的时候可能忽略了您的使用体验，如果您对上述版本不满意，我们提供了 `.spec` 文件和 `.py` 文件，如果您有兴趣可以自行优化代码打包安装。
>
> [在 Github 上打开 FastToolsForOldSeewo 开源仓库中的 Packages 文件夹](https://github.com/yuyudifiesh/seewo-fasttools/tree/main/Packages/pptxya/win7)

<br>

## PDF 阅读工具
使用此工具以缓解在一体机上使用默认浏览器打开 PDF 无响应问题。

若您所在学校未启用希沃冰点还原或者其他还原点工具，那么默认浏览器可能被篡改为 360 浏览器或者 Edge 浏览器，低性能一体机在打开过大 PDF 可能会造成小段时间的未响应。

![image](./view/pdfview.gif)

此工具无需联网，只包含最简单的 PDF 阅读功能，您可以通过右上角的设置按钮来调整顶部工具栏的大小以适配大屏操作。[【Releases】](https://github.com/yuyudifiesh/seewo-fasttools/releases)

| 适用系统 | 文件名 | 版本 | 更新时间 | 下载 |
| :---: | :---: | :---: | :---: | :---: |
| Windows7/10/11 | pdfview.exe | 1.0.0 | 2025 年 10 月 03 日 | [【下载】](./soft/pdfview.exe) |
