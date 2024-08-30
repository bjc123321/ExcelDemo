#-------------------------------------------------
#
# Project created by QtCreator 2024-08-17T09:10:08
#
#-------------------------------------------------

QT       += core gui axcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

# 启用预编译头文件
CONFIG += precompile_header

# 指定预编译头文件
PRECOMPILED_HEADER = precompiled.h


TARGET = ExcelDemo
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp

HEADERS  += mainwindow.h \
    precompiled.h

FORMS    += mainwindow.ui

include(src/xlsx/qtxlsx.pri)
