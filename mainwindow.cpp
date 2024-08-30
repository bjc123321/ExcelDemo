#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    // 创建菜单或按钮来触发这些槽函数
    connect(ui->pushButton, &QPushButton::clicked, this, &MainWindow::createAndWriteExcelFile);
    connect(ui->pushButton_2, &QPushButton::clicked, this, &MainWindow::readExcelFile);
    connect(ui->pushButton_3, &QPushButton::clicked, this, &MainWindow::updateExcelFile);
    //绑定创建复杂表的的槽
    connect(ui->pushButton_4, &QPushButton::clicked, this, &MainWindow::createTransientTestExcel);
    connect(ui->pushButton_5, &QPushButton::clicked, this, &MainWindow::updateTransientTestExcel);

    // 连接槽函数到按钮,稳态测试数据表
        connect(ui->pushButton_6, &QPushButton::clicked, this, &MainWindow::createStaticTestExcel);
        connect(ui->pushButton_7, &QPushButton::clicked, this, &MainWindow::updateStaticTestExcel);


}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::createAndWriteExcelFile()
{
    QString fileName = "test.xlsx";
    QXlsx::Document xlsx;

    // 写入数据到单元格
    xlsx.write("A1", "Header 1");
    xlsx.write("B1", "Header 2");
    xlsx.write("C1", "Header 3");

    xlsx.write("A2", 123);
    xlsx.write("B2", 456);
    xlsx.write("C2", 789);

    xlsx.write("A3", QString::fromUtf8("中文"));
    xlsx.write("B3", 654);
    xlsx.write("C3", 321);

    // 保存文件到桌面
    QString filePath = QDir::homePath() + "/Desktop/" + fileName;
    if (xlsx.saveAs(filePath)) {
        qDebug() << "Workbook saved successfully at:" << filePath;
    } else {
        qDebug() << "Failed to save the workbook.";
    }
}

void MainWindow::readExcelFile()
{
    QString fileName = "test.xlsx";
    QString filePath = QDir::homePath() + "/Desktop/" + fileName;
    QXlsx::Document xlsx(filePath);

        QVariant header1 = xlsx.read("A1");
        QVariant header2 = xlsx.read("B1");
        QVariant header3 = xlsx.read("C1");

        qDebug() << "Header 1:" << header1.toString();
        qDebug() << "Header 2:" << header2.toString();
        qDebug() << "Header 3:" << header3.toString();

        QVariant data1 = xlsx.read("A2");
        QVariant data2 = xlsx.read("B2");
        QVariant data3 = xlsx.read("C2");

        qDebug() << "Data 1:" << data1.toInt();
        qDebug() << "Data 2:" << data2.toInt();
        qDebug() << "Data 3:" << data3.toInt();

}

void MainWindow::updateExcelFile()
{
    QString fileName = "test.xlsx";
    QString filePath = QDir::homePath() + "/Desktop/" + fileName;
    QXlsx::Document xlsx(filePath);


        // 更新某个单元格的数据
        xlsx.write("A2", 321);
        xlsx.write("B2", 654);
        xlsx.write("C2", 987);

        // 保存更新
        if (xlsx.saveAs(filePath)) {
            qDebug() << "Workbook updated and saved successfully at:" << filePath;
        } else {
            qDebug() << "Failed to save the workbook.";
        }

}

void MainWindow::createTransientTestExcel()
{
    QXlsx::Document xlsx;

        // 设置标题和合并单元格
        xlsx.mergeCells("A1:N1");
        QXlsx::Format titleFormat;
        titleFormat.setFontSize(16);
        titleFormat.setFontBold(true);
        titleFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
        xlsx.write("A1", "瞬态试验报告（工频）", titleFormat);

        // 设置列宽
        xlsx.setColumnWidth(1, 12);
        xlsx.setColumnWidth(2, 12);
        xlsx.setColumnWidth(3, 12);
        xlsx.setColumnWidth(4, 12);
        xlsx.setColumnWidth(5, 12);
        xlsx.setColumnWidth(6, 12);
        xlsx.setColumnWidth(7, 12);
        xlsx.setColumnWidth(8, 12);
        xlsx.setColumnWidth(9, 12);
        xlsx.setColumnWidth(10, 12);
        xlsx.setColumnWidth(11, 12);
        xlsx.setColumnWidth(12, 12);
        xlsx.setColumnWidth(13, 12);
        xlsx.setColumnWidth(14, 12);

        // 设置表头部分
        QXlsx::Format headerFormat;
        headerFormat.setFontBold(true);
        headerFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
        headerFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
        headerFormat.setTextWarp(true);
        headerFormat.setBorderStyle(QXlsx::Format::BorderThin);

        // 第一行表头
        xlsx.mergeCells("A2:C2", headerFormat);
        xlsx.write("A2", "负载状况 (%)", headerFormat);

        xlsx.mergeCells("D2:H2", headerFormat);
        xlsx.write("D2", "初始电压 (V)", headerFormat);

        xlsx.mergeCells("I2:K2", headerFormat);
        xlsx.write("I2", "瞬态电压 (V)", headerFormat);

        xlsx.mergeCells("L2:L2", headerFormat);
        xlsx.write("L2", "电压稳定时间 (s)", headerFormat);

        xlsx.mergeCells("M2:M2", headerFormat);
        xlsx.write("M2", "初始频率 (Hz)", headerFormat);

        xlsx.mergeCells("N2:N2", headerFormat);
        xlsx.write("N2", "瞬态频率 (Hz)", headerFormat);

        // 第二行表头
        xlsx.write("A3", "负载状况 (%)", headerFormat);
        xlsx.write("B3", "有功功率 (KW)", headerFormat);
        xlsx.write("C3", "功率因数", headerFormat);
        xlsx.write("D3", "Uuno", headerFormat);
        xlsx.write("E3", "Uvns", headerFormat);
        xlsx.write("F3", "Uvo", headerFormat);
        xlsx.write("G3", "Uv", headerFormat);
        xlsx.write("H3", "Us", headerFormat);
        xlsx.write("I3", "Uuno", headerFormat);
        xlsx.write("J3", "Uvns", headerFormat);
        xlsx.write("K3", "Us", headerFormat);
        xlsx.write("L3", "电压稳定时间 (s)", headerFormat);
        xlsx.write("M3", "farb", headerFormat);
        xlsx.write("N3", "fmax", headerFormat);

        // 设置下方数据行
        QXlsx::Format dataFormat;
        dataFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
        dataFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
        dataFormat.setBorderStyle(QXlsx::Format::BorderThin);

        for (int row = 4; row <= 7; ++row) {
            xlsx.write(row, 1, "0->100", dataFormat);
            xlsx.write(row, 2, "", dataFormat);
            xlsx.write(row, 3, "", dataFormat);
            xlsx.write(row, 4, "", dataFormat);
            xlsx.write(row, 5, "", dataFormat);
            xlsx.write(row, 6, "", dataFormat);
            xlsx.write(row, 7, "", dataFormat);
            xlsx.write(row, 8, "", dataFormat);
            xlsx.write(row, 9, "", dataFormat);
            xlsx.write(row, 10, "", dataFormat);
            xlsx.write(row, 11, "", dataFormat);
            xlsx.write(row, 12, "", dataFormat);
            xlsx.write(row, 13, "", dataFormat);
            xlsx.write(row, 14, "", dataFormat);
        }

        // 设置最后部分负载状态
        xlsx.write("A8", "负载状况 (%)", headerFormat);
        xlsx.mergeCells("B8:D8", headerFormat);
        xlsx.write("B8", "0->100", headerFormat);
        xlsx.mergeCells("E8:G8", headerFormat);
        xlsx.write("E8", "100->0", headerFormat);
        xlsx.mergeCells("H8:J8", headerFormat);
        xlsx.write("H8", "0->100", headerFormat);
        xlsx.mergeCells("K8:M8", headerFormat);
        xlsx.write("K8", "100->0", headerFormat);

        xlsx.write("A9", "瞬态频率调整系数δfss (%)", headerFormat);
        xlsx.write("A10", "瞬态电压调整系数δUss (%)", headerFormat);

        for (int col = 2; col <= 14; col += 3) {
            xlsx.write(9, col, "", dataFormat);
            xlsx.write(9, col+1, "", dataFormat);
            xlsx.write(9, col+2, "", dataFormat);

            xlsx.write(10, col, "", dataFormat);
            xlsx.write(10, col+1, "", dataFormat);
            xlsx.write(10, col+2, "", dataFormat);
        }

        // 保存文件到桌面
        QString filePath = QDir::homePath() + "/Desktop/TestReport.xlsx";
        if (xlsx.saveAs(filePath)) {
            qDebug() << "Excel table created and saved at:" << filePath;
        } else {
            qDebug() << "Failed to save the Excel table.";
        }
}

void MainWindow::updateTransientTestExcel()
{
    QString filePath = QDir::homePath() + "/Desktop/TestReport.xlsx";
    QXlsx::Document xlsx(filePath);


    QXlsx::Format dataFormat;
    dataFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    dataFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    dataFormat.setBorderStyle(QXlsx::Format::BorderThin);

    // 添加模拟数据
    int startRow = 4;
    for (int row = startRow; row < startRow + 4; ++row) {
        xlsx.write(row, 2, 50 + row, dataFormat); // 有功功率
        xlsx.write(row, 3, 0.85, dataFormat); // 功率因数
        xlsx.write(row, 4, 220 + row, dataFormat); // Uuno
        xlsx.write(row, 5, 225 + row, dataFormat); // Uvns
        xlsx.write(row, 6, 230 + row, dataFormat); // Uvo
        xlsx.write(row, 7, 235 + row, dataFormat); // Uv
        xlsx.write(row, 8, 240 + row, dataFormat); // Us
        xlsx.write(row, 9, 220 + row, dataFormat); // 瞬态 Uuno
        xlsx.write(row, 10, 225 + row, dataFormat); // 瞬态 Uvns
        xlsx.write(row, 11, 230 + row, dataFormat); // 瞬态 Us
        xlsx.write(row, 12, 0.5 + (row - startRow) * 0.1, dataFormat); // 电压稳定时间
        xlsx.write(row, 13, 50 + row, dataFormat); // farb
        xlsx.write(row, 14, 60 - (row - startRow), dataFormat); // fmax
    }

    // 添加数据到最后部分
    xlsx.write(9, 2, 5.0, dataFormat);
    xlsx.write(9, 5, 10.0, dataFormat);
    xlsx.write(9, 8, 7.5, dataFormat);
    xlsx.write(9, 11, 12.5, dataFormat);

    xlsx.write(10, 2, 2.0, dataFormat);
    xlsx.write(10, 5, 4.0, dataFormat);
    xlsx.write(10, 8, 3.0, dataFormat);
    xlsx.write(10, 11, 6.0, dataFormat);

    // 保存更新后的文件
    if (xlsx.saveAs(filePath)) {
        qDebug() << "Mock data added and saved successfully at:" << filePath;
    } else {
        qDebug() << "Failed to save the Excel table with mock data.";
    }

}


void MainWindow::createStaticTestExcel()
{
    QXlsx::Document xlsx;

    // 第一部分：生成第一个表格
    xlsx.mergeCells("A1:N1");
    QXlsx::Format titleFormat;
    titleFormat.setFontSize(16);
    titleFormat.setFontBold(true);
    titleFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    xlsx.write("A1", "电流、电压和频率数据", titleFormat);

    // 设置列宽
    for (int col = 1; col <= 14; ++col) {
        xlsx.setColumnWidth(col, 12);
    }

    // 第一行表头
    QXlsx::Format headerFormat;
    headerFormat.setFontBold(true);
    headerFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    headerFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    headerFormat.setTextWarp(true);
    headerFormat.setBorderStyle(QXlsx::Format::BorderThin);

    xlsx.mergeCells("A2:C2", headerFormat);
    xlsx.write("A2", "负载状况 (%)", headerFormat);

    xlsx.mergeCells("D2:F2", headerFormat);
    xlsx.write("D2", "电流 (A)", headerFormat);

    xlsx.mergeCells("G2:L2", headerFormat);
    xlsx.write("G2", "电压 (V)", headerFormat);

    xlsx.mergeCells("M2:N2", headerFormat);
    xlsx.write("M2", "频率 (Hz)", headerFormat);

    // 第二行表头
    xlsx.write("A3", "负载状况 (%)", headerFormat);
    xlsx.write("B3", "有功功率 (KW)", headerFormat);
    xlsx.write("C3", "功率因数 PF", headerFormat);
    xlsx.write("D3", "Ia", headerFormat);
    xlsx.write("E3", "Ib", headerFormat);
    xlsx.write("F3", "Ic", headerFormat);
    xlsx.write("G3", "Uan", headerFormat);
    xlsx.write("H3", "Ubn", headerFormat);
    xlsx.write("I3", "Ucn", headerFormat);
    xlsx.write("J3", "Uab", headerFormat);
    xlsx.write("K3", "Ubc", headerFormat);
    xlsx.write("L3", "Uca", headerFormat);
    xlsx.write("M3", "farb", headerFormat);
    xlsx.write("N3", "fmax", headerFormat);

    // 设置数据行
    QXlsx::Format dataFormat;
    dataFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    dataFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    dataFormat.setBorderStyle(QXlsx::Format::BorderThin);

    QStringList loadConditions = {"0", "25", "50", "75", "100", "75", "50", "25", "0"};
    for (int row = 4; row <= 12; ++row) {
        //第4行至第12行，的第一列填上loadConditions的每一个元素
        xlsx.write(row, 1, loadConditions[row - 4], dataFormat);
        for (int col = 2; col <= 14; ++col) {
            //第4行至第12行，的其他列填上空字符串
            xlsx.write(row, col, "", dataFormat);
        }
    }

    // 第二部分：生成第二个表格
    int startRow = 14; // 第二个表格从第14行开始

    xlsx.mergeCells(QString("A%1:J%1").arg(startRow)); // 标题合并单元格
    xlsx.write(QString("A%1").arg(startRow), "稳态电压和频率调整数据", titleFormat);

    startRow += 2;

    // 第一行表头
    xlsx.mergeCells(QString("A%1:A%2").arg(startRow).arg(startRow + 1), headerFormat);
    xlsx.write(QString("A%1").arg(startRow), "负载状况 (%)", headerFormat);

    QStringList conditions = {"0", "25", "50", "75", "100", "75", "50", "25", "0"};
    for (int col = 2; col <= 10; ++col) {
        xlsx.write(startRow, col, conditions[col - 2], headerFormat);
    }

    // 第二行数据标签
    QStringList labels = {"稳态电压调整系数δU (%)", "电压波动率δUb (%)", "稳态频率调整系数δf (%)", "频率波动率δf (%)"};
    for (int row = startRow + 1; row <= startRow + 4; ++row) {
        xlsx.write(row, 1, labels[row - startRow - 1], headerFormat);
        for (int col = 2; col <= 10; ++col) {
            xlsx.write(row, col, "", dataFormat);
        }
    }

    // 保存文件到桌面
    QString filePath = QDir::homePath() + "/Desktop/StaticTestReport.xlsx";
    if (xlsx.saveAs(filePath)) {
        qDebug() << "Excel report with two tables created and saved at:" << filePath;
    } else {
        qDebug() << "Failed to save the Excel report.";
    }
}

void MainWindow::updateStaticTestExcel()
{
    QString filePath = QDir::homePath() + "/Desktop/StaticTestReport.xlsx";
    QXlsx::Document xlsx(filePath);


    QXlsx::Format dataFormat;
    dataFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    dataFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    dataFormat.setBorderStyle(QXlsx::Format::BorderThin);

    // 添加模拟数据到第一个表格
    int startRow = 4;
    for (int row = startRow; row <= 12; ++row) {
        xlsx.write(row, 2, 100 + (row - startRow) * 20, dataFormat); // 有功功率
        xlsx.write(row, 3, 0.95, dataFormat); // 功率因数
        xlsx.write(row, 4, 200 + (row - startRow) * 5, dataFormat); // Ia
        xlsx.write(row, 5, 210 + (row - startRow) * 5, dataFormat); // Ib
        xlsx.write(row, 6, 220 + (row - startRow) * 5, dataFormat); // Ic
        xlsx.write(row, 7, 230 + (row - startRow) * 5, dataFormat); // Uan
        xlsx.write(row, 8, 240 + (row - startRow) * 5, dataFormat); // Ubn
        xlsx.write(row, 9, 250 + (row - startRow) * 5, dataFormat); // Ucn
        xlsx.write(row, 10, 260 + (row - startRow) * 5, dataFormat); // Uab
        xlsx.write(row, 11, 270 + (row - startRow) * 5, dataFormat); // Ubc
        xlsx.write(row, 12, 280 + (row - startRow) * 5, dataFormat); // Uca
        xlsx.write(row, 13, 50 + (row - startRow), dataFormat); // farb
        xlsx.write(row, 14, 60 - (row - startRow), dataFormat); // fmax
    }

    // 添加模拟数据到第二个表格
    startRow = 17; // 第二个表格从第17行开始数据填充
    xlsx.write(startRow, 2, 0.5, dataFormat); // 稳态电压调整系数δU
    xlsx.write(startRow + 1, 2, 0.4, dataFormat); // 电压波动率δUb
    xlsx.write(startRow + 2, 2, 0.3, dataFormat); // 稳态频率调整系数δf
    xlsx.write(startRow + 3, 2, 0.2, dataFormat); // 频率波动率δf

    for (int col = 3; col <= 10; ++col) {
        xlsx.write(startRow, col, 0.5 + col * 0.1, dataFormat); // 模拟数据
        xlsx.write(startRow + 1, col, 0.4 + col * 0.1, dataFormat);
        xlsx.write(startRow + 2, col, 0.3 + col * 0.1, dataFormat);
        xlsx.write(startRow + 3, col, 0.2 + col * 0.1, dataFormat);
    }

    // 保存更新后的文件
    if (xlsx.saveAs(filePath)) {
        qDebug() << "Mock data added and saved successfully at:" << filePath;
    } else {
        qDebug() << "Failed to save the Excel table with mock data.";
    }


}
