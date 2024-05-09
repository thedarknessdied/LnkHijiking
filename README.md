# LnkHijiking
windows Lnk(.lnk) hook

## 希望
  希望各位朋友们多多测试，多多提出BUG

## 1.attack.vbs
### 使用
  修改脚本中的 implant 参数为后门文件路径，修改脚本中的 curUserDesktop 参数为需要替换lnk指向的文件夹
### 原理
  windows中快捷方式可以通过 DEO 对象运行程序，这里我们通过劫持运行逻辑的方式，释放一个中间执行文件，在不影响原有程序执行的情况下加载后门程序的运行。期间也可以通过修改HookLnk函数中的newTarget参数的值来设定中间执行文件的路径，可以通过修改Set File = FSO.CreateTextFile(newTarget,True)中的File对象的内容修改中间转发文件的运行逻辑

## 提示
本工具集合只为学习使用，不要用其进行非法操作
