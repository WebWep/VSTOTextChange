# VSTOTextChange
MS Office VSTO Word 插件, 监听文本输入.

1. 开发信息

   Windows 7 x64<br/>
   Visual Studio 2017<br/>
   
   Word 2013 和 2016 VSTO 外接程序<br/>
   .NET Web API<br/>
   .NET Framework 4.5

2. 实现功能

    (1)点击按钮，发送整个文章内容到指定接口。<br/>
    (2)打开开关，实时监听文章更新和修改，并且将更改后的句子发送到接口。

3. 运行

    项目"WordAddInVSTO"为Word插件项目<br/>
    项目"WordAPI"为WebAPI项目，作为Word插件的接口

    "WordAddInVSTO"中有 APP.config 配置文件，用来配置接口地址<br/>
    默认为："WordAPI" 项目运行后的 localhost 地址。


4. 运行结果

   "WordADDInVSTO"项目为启动项，运行后程序自动打开Word，<br/>
    并且自动加载加载项，在Word标签页中会看到"TextChangeDemo"。<br/>
    点击"TextChangeDemo"，看见"远程保存"按钮和"更新开关"开关。

    ![VSTO add-in](https://github.com/WebWep/VSTOTextChange/blob/master/MDImage/p1.png)
    
    "远程保存"按钮，点击后发送整个文档内容到接口，接口接受到内容，<br/>
    将内容保存到"WordAPI"项目根目录下的 Upload/Document.txt 文件中。

    ![VSTO add-in](https://github.com/WebWep/VSTOTextChange/blob/master/MDImage/p2.png)

    "更新开关"打开后，Word插件的任务窗格同时打开，程序自动监听文档更新，<br/>
    并且将更新后的句子，实时保存到"WordAPI"项目根目录下的 Upload/Sentence.txt 文件中。<br/>
    同时任务窗格会实时显示当前句子。

    ![VSTO add-in](https://github.com/WebWep/VSTOTextChange/blob/master/MDImage/p3.png)

