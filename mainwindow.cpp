#include "mainwindow.h"

#include <QTableWidgetItem>
#include <QHeaderView>
#include <QDebug>
#include <QtXlsx>
#include <QVBoxLayout>
#include <QWidget>
#include "xlsxabstractsheet.h"

QTXLSX_USE_NAMESPACE        //该命名空间不可少

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
{
    init();
}

MainWindow::~MainWindow()
{

}

void MainWindow::init()
{
    //初始化控件和布局
    QWidget *widget = new QWidget(this);
    QHBoxLayout *hLayout = new QHBoxLayout;
    QVBoxLayout *vLayout = new QVBoxLayout;

    tableWidget = new QTableWidget(8, 10, widget);
    btn = new QPushButton(tr("生成Excel"), widget);
    connect(btn, SIGNAL(clicked()), this, SLOT(slot_writeToExcel()));

//    hLayout->addSpacing(500);
    hLayout->addStretch(10);
    hLayout->addWidget(btn);

    vLayout->addWidget(tableWidget);
    vLayout->addLayout(hLayout);

    widget->setLayout(vLayout);

    setCentralWidget(widget);
    resize(500, 300);

    int rows = tableWidget->rowCount();         //行数
    int cols = tableWidget->columnCount();      //列数
    for(int i = 0; i < rows; i++){
        for(int j = 0; j < cols; j++){
            QTableWidgetItem *item = new QTableWidgetItem(tr("%1").arg(i+j));
            item->setTextAlignment(Qt::AlignCenter);
            tableWidget->setItem(i, j, item);
        }
    }

    tableWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    tableWidget->setEditTriggers(QHeaderView::NoEditTriggers);      //不可编辑
    tableWidget->setSelectionBehavior(QHeaderView::SelectRows);     //整行选中
    tableWidget->horizontalHeader()->setVisible(false);             //隐藏列表头
    tableWidget->verticalHeader()->setVisible(false);               //隐藏行表头
}

void MainWindow::slot_writeToExcel()
{
    QXlsx::Document xlsx;

    int rows = tableWidget->rowCount();         //行数
    int cols = tableWidget->columnCount();      //列数

    QString text;
    for(int i = 0; i < rows; i++){
        for(int j = 0; j < cols; j++){
            text = tableWidget->item(i, j)->text();
            xlsx.write(i + 1, j + 1, text);
        }
    }

    QXlsx::Format format;                       //格式
    format.setFont(QFont(tr("华文新魏")));       //字体
    format.setFontBold(true);                   //加粗
    format.setFontItalic(true);                 //倾斜
    format.setFontUnderline(Format::FontUnderlineSingle);   //下划线
    format.setFontColor(Qt::red);               //字体颜色
    format.setHorizontalAlignment(Format::AlignRight);  //对齐方式

    xlsx.setRowFormat(3, 5, format);
    xlsx.setRowHidden(2, true);             //隐藏第二行

    xlsx.renameSheet(tr("Sheet1"), tr("工作计划"));     //重命名当前Sheet

    xlsx.copySheet("工作计划", "CopyOfTheFirst");       //无格式拷贝

    xlsx.selectSheet("CopyOfTheFirst");                 //设为当前显示Sheet
    xlsx.write(25, 2, "On the Copy Sheet");             //写
    qDebug() << "111:" << xlsx.read(25, 2).toString();  //读
    qDebug() << "222:" << xlsx.read("B25").toString();  //读

    xlsx.copySheet("CopyOfTheFirst", "work1");          //无格式拷贝
    xlsx.moveSheet("work1", 0);                         //移动Sheet在Excel中的位置

    xlsx.sheet("work1")->setVisible(true);              //显示指定Sheet
//    xlsx.sheet("CopyOfTheFirst")->setVisible(false);    //隐藏指定Sheet

//    xlsx.deleteSheet("CopyOfTheFirst");                //删除指定Sheet

    xlsx.addSheet(tr("work2"));                         //添加一个Sheet
//    xlsx.sheet("work2")->setSheetState(AbstractSheet::SS_VeryHidden);   //设置指定Sheet状态为隐藏

    QString curSheetName = xlsx.currentSheet()->sheetName();
    qDebug() << "curSheetName" << curSheetName;
    xlsx.currentWorksheet()->setGridLinesVisible(false);    //不显示网格

    xlsx.mergeCells("B1:B3");                           //合并单元格B1:B3

    //设置属性
    xlsx.write("A1", "View the properties through:");
    xlsx.write("A2", "Office Button -> Prepare -> Properties option in Excel");

    xlsx.setDocumentProperty("title", "This is an example spreadsheet");
    xlsx.setDocumentProperty("subject", "With document properties");
    xlsx.setDocumentProperty("creator", "Debao Zhang");
    xlsx.setDocumentProperty("company", "HMICN");
    xlsx.setDocumentProperty("category", "Example spreadsheets");
    xlsx.setDocumentProperty("keywords", "Sample, Example, Properties");
    xlsx.setDocumentProperty("description", "Created with Qt Xlsx");

    xlsx.selectSheet("工作计划");                 //设为当前显示Sheet

    //获取表格的行数、列数
    QXlsx::CellRange range;
    range = xlsx.dimension();
    int rowCount = range.rowCount();
    int colCount = range.columnCount();
    qDebug() << "rowCount:" << rowCount << "colCount:" << colCount << xlsx.currentSheet()->sheetName();

    //输出表格内容
    for (int i = 1; i <= rowCount; i++){
        for (int j = 1; j <= colCount; j++){
            qDebug() << i << j << xlsx.cellAt(i, j)->value().toString();
        }
    }

    xlsx.saveAs("Test.xlsx");                           //另存为
}
