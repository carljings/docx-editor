公文字体部署说明（纯内网离线）：

1. 将客户已取得网页部署授权的原版公文字体放入当前目录（public/fonts/），例如：
   - 方正小标宋_GBK.ttf
   - 方正仿宋_GBK.ttf
   - 方正楷体_GBK.ttf
   - 方正黑体_GBK.ttf
   - 仿宋_GB2312.ttf / 仿宋.ttf
   - 楷体_GB2312.ttf / 楷体.ttf
   - 宋体.ttf / 黑体.ttf
   - calibri.ttf / times.ttf

2. 未放置原版授权字体时，@weboffice/vue 默认自动降级加载本目录 fandol/ 下的开源 Fandol 0.3 字体（GPL/开源许可，无商业版权风险）与 @docx-editor.dev/fonts 内置西文字体。

3. 严禁将本机 WPS 桌面安装包中的商业字形（如 FangS-SC / ShuS-SC / HYKaiTiJ）提交至仓库或打包至生产环境。
