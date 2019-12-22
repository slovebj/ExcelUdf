# 开源说明
此项目为Excel催化剂插件里的一部分，主要是自定义函数篇，使用ExcelDna的框架开发，可使用.Net语言，开发自定义函数供Excel使用，使用体验也很不错，具体优点如下：
1. 可以有充足的注释说明供Excel用户调用时查看，且无论是在函数体书写还是在函数向导上都可很清晰地看到注释信息，详细到每个函数的参数都可设置注释信息。
![函数向导注释效果](https://upload-images.jianshu.io/upload_images/9936495-dbc5085ef3e6489d.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
![书写函数体时的注释效果](https://upload-images.jianshu.io/upload_images/9936495-bd1f58a87f2e9c85.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

2. 可方便部署
只需生成打包成一个xll文件，即可发布给用户使用，用户可按xlam的加载项方式使用一次安装，日后长期使用或单次使用只需双击此xll加载自定义函数即可使用。

**亦可自行用程序来打包封装，实现用户一键安装，因xll文件区分32位和64位Excel运行，故需考虑用户Excel的位数将对应位数的xll安装到用户电脑内，
后期可开源Console控制台程序的方式安装xll**

xcopy "$(SolutionDir)\packages\ExcelDna.AddIn.1.0.0\tools\ExcelDna.xll" "$(TargetDir)ExcelUDF-AddIn.xll*" /C /Y
xcopy "$(TargetDir)ExcelUDF-AddIn.dna*" "$(TargetDir)ExcelUDF-AddIn64.dna*" /C /Y
xcopy "$(SolutionDir)\packages\ExcelDna.AddIn.1.0.0\tools\ExcelDna64.xll" "$(TargetDir)ExcelUDF-AddIn64.xll*" /C /Y
"$(SolutionDir)\packages\ExcelDna.AddIn.1.0.0\tools\ExcelDnaPack.exe" "$(TargetDir)ExcelUDF-AddIn.dna" /Y
"$(SolutionDir)\packages\ExcelDna.AddIn.1.0.0\tools\ExcelDnaPack.exe" "$(TargetDir)ExcelUDF-AddIn64.dna" /Y