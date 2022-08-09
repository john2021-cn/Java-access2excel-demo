# Java-access2excel-demo
利用Apache POI组件将access数据库数据填入excel表格  
**注意：已对代码进行简化处理，但不影响阅读思路**

该项目由三个类构成，分别是OneClickReport、Report、ButtonListener。  
## OneClickReport
这是启动类，负责将Report类进行实例化，并调用Report类中的initUI方法显示界面。
## Report
Report类继承JFrame类，由四个方法构成分别是initUI、showFileOpenDialog、showDBOpenDialog、showExcel。
### initUI
initUI方法负责绘制程序界面，绘制界面结束后需要实例化ButtonListener类来实现按键事件监听。
### showFileOpenDialog
showFileOpenDialog方法负责打开Excel文件对话框，通过实例化JFileChooser类来实现。
### showDBOpenDialog
showDBOpenDialog方法负责打开数据库文件对话框，通过实例化JFileChooser类来实现。
### showExcel
showExcel方法负责根据用户选择的excel模板来创建数据库连接读取数据并把数据写入到excel文件中。
## ButtonListener
ButtonListener类实现ActionListener事件监听器，并重写actionPerformed方法，分别监听用户点击"获取Excel报告位置"、"获取数据库Data位置"、"填写报告"按钮时需要触发的方法。
