#-------------------------------------------------
#
# Project created by QtCreator 2015-06-07T11:16:43
#
#-------------------------------------------------

QT       -= gui
QT       += core
QT       += xlsx
TARGET = Xlsxcontrol
TEMPLATE = lib

DEFINES += XLSXCONTROL_LIBRARY

SOURCES += xlsxcontrol.cpp

HEADERS += xlsxcontrol.h

unix {
    target.path = /usr/lib
    INSTALLS += target
}
