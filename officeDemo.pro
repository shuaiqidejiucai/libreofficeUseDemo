QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17
DEFINES += LIBOLECF_HAVE_WIDE_CHARACTER_TYPE
# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
    main.cpp \
    mainwindow.cpp

HEADERS += \
    mainwindow.h

FORMS += \
    mainwindow.ui

INCLUDEPATH += "E:/libreoffice25.2.5x64/sdk/inc"
INCLUDEPATH += "E:/libreoffice25.2.5x64/sdk/include"
INCLUDEPATH += "D:/libolecfNew/libolecf/include"
LIBS += -L"E:/libreoffice25.2.5x64/sdk/lib" \
        -licppu \
        -licppuhelper \
        -lipurpenvhelper \
        -lisal \
        -lisalhelper
LIBS += -L"D:/libolecfNew/libolecf/msvscpp/x64/VSDebug"
        -llibolecf


# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target
