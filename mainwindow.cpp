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


    connect(ui->pushButton_8, &QPushButton::clicked, this, &MainWindow::createWorkConditionExcel);


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
            qDebug() << "Failed to save the workbook. Perhaps Excel is opened! ";
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

void MainWindow::createWorkConditionExcel()
{

      QXlsx::Document xlsx;
    // 设置标题和合并单元格
    xlsx.mergeCells("A1:X1");
    QXlsx::Format titleFormat;
    titleFormat.setFontSize(16);
    titleFormat.setFontBold(true);
    titleFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    xlsx.write("A1", "连续工况试验报告（工频）", titleFormat);

    // 创建单独用于第2-6行的格式
       QXlsx::Format headerFormat;
       headerFormat.setFontName("DengXian");  // 字体等线
       headerFormat.setFontSize(10.5);  // 字号10.5

       // 填写第 2-6 行数据并应用格式
       xlsx.write("A2", "测试单位：", headerFormat);
       xlsx.write("A3", "额定功率（KW）：", headerFormat);
       xlsx.write("A4", "产品名称：", headerFormat);
       xlsx.write("A5", "油机型号：", headerFormat);
       xlsx.write("A6", "环境温度（℃）：", headerFormat);

       xlsx.write("F2", "执行标准：", headerFormat);
       xlsx.write("F3", "额定电压 (V)：", headerFormat);
       xlsx.write("F4", "产品编号：", headerFormat);
       xlsx.write("F5", "油机编号：", headerFormat);
       xlsx.write("F6", "相对湿度：", headerFormat);

       xlsx.write("L3", "额定频率（Hz）：", headerFormat);
       xlsx.write("L4", "产品编号：", headerFormat);
       xlsx.write("L5", "电机型号：", headerFormat);
       xlsx.write("L6", "大气压力（KPa）：", headerFormat);

       xlsx.write("S2", "试验时间：", headerFormat);
       xlsx.write("S3", "产品状态：", headerFormat);
       xlsx.write("S4", "相/线：", headerFormat);
       xlsx.write("S5", "电机编号：", headerFormat);
       xlsx.write("S6", "测试负责人：", headerFormat);

       // 合并顶部信息的单元格
       xlsx.mergeCells("A2:E2");
       xlsx.mergeCells("A3:E3");
       xlsx.mergeCells("A4:E4");
       xlsx.mergeCells("A5:E5");
       xlsx.mergeCells("A6:E6");

       xlsx.mergeCells("F2:K2");
       xlsx.mergeCells("F3:K3");
       xlsx.mergeCells("F4:K4");
       xlsx.mergeCells("F5:K5");
       xlsx.mergeCells("F6:K6");

       xlsx.mergeCells("L3:R3");
       xlsx.mergeCells("L4:R4");
       xlsx.mergeCells("L5:R5");
       xlsx.mergeCells("L6:R6");

       xlsx.mergeCells("S2:X2");
       xlsx.mergeCells("S3:X3");
       xlsx.mergeCells("S4:X4");
       xlsx.mergeCells("S5:X5");
       xlsx.mergeCells("S6:X6");

        // 创建格式：粗框线、居中对齐、字体等线，字号10.5
            QXlsx::Format format;
            format.setBorderStyle(QXlsx::Format::BorderThin);  // 细框线
            format.setHorizontalAlignment(QXlsx::Format::AlignHCenter);  // 水平居中
            format.setVerticalAlignment(QXlsx::Format::AlignVCenter);  // 垂直居中
            format.setFontName("DengXian");  // 设置字体为等线
            format.setFontSize(10.5);  // 设置字号为10.5
            format.setTextWarp(true);  // 自动换行

            // 设置列标题（从第7行开始）
            xlsx.write("A7", "序号", format);
            xlsx.write("B7", "记录时间", format);
            xlsx.write("C7", "电路 (A)", format);
            xlsx.write("C8", "Iu", format);
            xlsx.write("D8", "Iv", format);
            xlsx.write("E8", "Iw", format);
            xlsx.write("F7", "功率 (KW)", format);
            xlsx.write("F8", "Pu", format);
            xlsx.write("G8", "Pv", format);
            xlsx.write("H8", "Pw", format);
            xlsx.write("I7", "电压 (V)", format);
            xlsx.write("I8", "Uun", format);
            xlsx.write("J8", "Uvn", format);
            xlsx.write("K8", "Uwn", format);
            xlsx.write("L7", "频率 (Hz)", format);
            xlsx.write("M7", "功率因数", format);  // 修正：添加功率因数
            xlsx.write("M8", "因数", format);
            xlsx.write("N7", "冷却介质温度 (℃)", format);
            xlsx.write("N8", "1", format);
            xlsx.write("O8", "2", format);
            xlsx.write("P7", "油温 (℃)", format);
            xlsx.write("Q7", "油压 (KPa)", format);
            xlsx.write("R7", "环境温度 (℃)", format);
            xlsx.write("R8", "1", format);
            xlsx.write("S8", "2", format);
            xlsx.write("T7", "水温 (℃)", format);
            xlsx.write("U7", "相对湿度 (%)", format);
            xlsx.write("V7", "大气压力 (KPa)", format);
            xlsx.write("W7", "添加燃油时间", format);
            xlsx.write("X7", "电压变化 (%)", format);

            // 合并表格中的单元格
            xlsx.mergeCells("A7:A8", format);
            xlsx.mergeCells("B7:B8", format);
            xlsx.mergeCells("C7:E7", format);
            xlsx.mergeCells("F7:H7", format);
            xlsx.mergeCells("I7:K7", format);
            xlsx.mergeCells("L7:L8", format);
            xlsx.mergeCells("M7:M8", format);  // 修正：合并功率因数的单元格
            xlsx.mergeCells("N7:O7", format);
            xlsx.mergeCells("P7:P8", format);
            xlsx.mergeCells("Q7:Q8", format);
            xlsx.mergeCells("R7:S7", format);
            xlsx.mergeCells("T7:T8", format);
            xlsx.mergeCells("U7:U8", format);
            xlsx.mergeCells("V7:V8", format);
            xlsx.mergeCells("W7:W8", format);
            xlsx.mergeCells("X7:X8", format);  // 合并电压变化的单元格

            // 填充序号行（9-24行，只生成空表格，无模拟数据）
            for (int row = 9; row <= 24; ++row) {
                xlsx.write(QString("A%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("B%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("C%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("D%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("E%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("F%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("G%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("H%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("I%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("J%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("K%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("L%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("M%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("N%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("O%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("P%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("Q%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("R%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("S%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("T%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("U%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("V%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("W%1").arg(row), "", format);  // 空单元格
                xlsx.write(QString("X%1").arg(row), "", format);  // 空单元格
            }



        // 保存Excel文件到桌面
        QString filePath = QDir::homePath() + "/Desktop/WorkingCondition.xlsx";
        if (xlsx.saveAs(filePath)) {
            qDebug() << "Excel report with two tables created and saved at:" << filePath;
        } else {
            qDebug() << "Failed to save the Excel report.";
        }

}

