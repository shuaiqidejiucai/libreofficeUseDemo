#include "mainwindow.h"

#include <QApplication>

//#include "com/sun/star/

//#define OUSTR(x) OUString(RTL_CONSTASCII_USTRINGPARAM(x))
int main(int argc, char *argv[])
{
    QApplication a(argc, argv);








    //关闭文档
    // Reference<util::XCloseable> xClosable(doc, uno::UNO_QUERY);
    // if (xClosable.is()) xClosable->close(true);
    MainWindow w;
    w.show();
    return a.exec();
}
