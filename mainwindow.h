#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDebug>
#include <QDir>
#include "xlsxdocument.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void createAndWriteExcelFile();  // 创建并写入Excel文件
    void readExcelFile();            // 读取Excel文件
    void updateExcelFile();          // 更新Excel文件

    void createTransientTestExcel();      //创建瞬态测试表
    void updateTransientTestExcel();    // 更i新瞬态测试模拟数据到Excel表格的槽函数

    void createStaticTestExcel();  // 创建稳态测试Excel表格的槽函数
    void updateStaticTestExcel();  // 更新稳态测试Excel表格模拟数据的槽函数

    void createWorkConditionExcel(); //创建工况试验报告

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
