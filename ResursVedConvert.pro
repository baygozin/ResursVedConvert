#-------------------------------------------------
#
# Project created by QtCreator 2015-01-29T10:44:55
#
#-------------------------------------------------

#QMAKE_CXXFLAGS += -std=c++14
#QMAKE_CXXFLAGS_DEBUG += -pg
#QMAKE_LFLAGS_DEBUG += -pg
#QMAKE_LFLAGS += -static -static-libgcc
QMAKE_CXXFLAGS_RELEASE += -MP9
QMAKE_CFLAGS_DEBUG += -MP9

include(3rdparty/qtxlsx/src/xlsx/qtxlsx.pri)

QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = ResursVedConvert
TEMPLATE = app

SOURCES += main.cpp\
        mainwindow.cpp \
    resurssection.cpp \
    laborman.cpp \
    machine.cpp \
    materials.cpp \
    bovutils.cpp \
    equipment.cpp

HEADERS  += mainwindow.h \
    resurssection.h \
    laborman.h \
    machine.h \
    materials.h \
    bovutils.h \
    equipment.h

FORMS    += mainwindow.ui

RESOURCES += \
    resources.qrc

RC_ICONS = panda.ico
