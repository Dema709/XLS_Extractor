QT -= gui

CONFIG += c++11 console
CONFIG -= app_bundle

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
        BasicExcel.cpp \
        main.cpp

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

HEADERS += \
    BasicExcel.hpp

TARGET = XLS_Extractor
VERSION = 1.2
RC_ICONS = icon.ico
QMAKE_TARGET_COMPANY = ChakaPon
QMAKE_TARGET_PRODUCT = Xls extractor
QMAKE_TARGET_DESCRIPTION = Extract xls file into new in the same folder
QMAKE_TARGET_COPYRIGHT = (c) ChakaPon
