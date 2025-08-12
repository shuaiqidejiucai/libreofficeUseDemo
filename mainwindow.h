#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "com/sun/star/uno/Reference.hxx"
#include "com/sun/star/uno/XComponentContext.hpp"
#include "rtl/bootstrap.hxx"
#include <cppuhelper/bootstrap.hxx>
#include "com/sun/star/lang/XMultiComponentFactory.hpp"
#include "com/sun/star/lang/XMultiServiceFactory.hpp"
#include "com/sun/star/uno/XInterface.hpp"
#include "rtl/ustring.hxx"
#include "com/sun/star/frame/XDesktop.hpp"
#include "com/sun/star/frame/XComponentLoader.hpp"
#include "com/sun/star/lang/XComponent.hpp"
#include "com/sun/star/uno/Sequence.h"
#include "com/sun/star/beans/PropertyValue.hpp"
#include "com/sun/star/text/XTextDocument.hpp"
#include "cppuhelper/compbase1.hxx"
#include "com/sun/star/container/XEnumerationAccess.hpp"
#include "com/sun/star/frame/XStorable.hpp"
#include <com/sun/star/uno/Type.hxx>
#include <com/sun/star/drawing/XDrawPageSupplier.hpp>
#include <com/sun/star/graphic/XGraphic.hpp>
#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/document/XGraphicStorageHandler.hpp>
#include <com/sun/star/document/GraphicStorageHandler.hpp>
#include <com/sun/star/awt/XBitmap.hpp>
#include <com/sun/star/util/XCloseable.hpp>
#include <com/sun/star/util/XSearchDescriptor.hpp>
#include <com/sun/star/util/XReplaceDescriptor.hpp>
#include <com/sun/star/util/XReplaceable.hpp>
QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE
using namespace com::sun::star;
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void reLoader();
private slots:
    void on_pushButton_clicked();

    void on_pushButton_2_clicked();

private:
    Ui::MainWindow *ui;
    com::sun::star::uno::Reference<com::sun::star::lang::XComponent> xComponent;
    com::sun::star::uno::Reference<com::sun::star::frame::XComponentLoader> xLoader;
    com::sun::star::uno::Reference<com::sun::star::text::XTextDocument> xTextDoc;
    com::sun::star::uno::Reference<com::sun::star::frame::XDesktop>     xDesktop;
};
#endif // MAINWINDOW_H
