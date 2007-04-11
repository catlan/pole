TEMPLATE = app
TARGET = poleview

HEADERS += ../pole.h
HEADERS += poleview.h

SOURCES += ../pole.cpp
SOURCES += poleview.cpp

INCLUDEPATH += ..

CONFIG += qt
CONFIG += thread 
CONFIG += warn_on
CONFIG += exceptions
CONFIG += release
#The following line was inserted by qt3to4
QT +=  qt3support 
