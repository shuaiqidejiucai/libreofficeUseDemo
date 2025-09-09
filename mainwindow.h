#ifndef MAINWINDOW_H
#define MAINWINDOW_H
#include <libolecf.h>
#include <libbfio_handle.h>
#include <libbfio_memory_range.h>

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
#include <com/sun/star/task/XJobExecutor.hpp>
#include <cppuhelper/implbase1.hxx>
#include  <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/embed/ElementModes.hpp>
#include <com/sun/star/document/XTypeDetection.hpp>
#include <com/sun/star/table/XCellRange.hpp>
#include <com/sun/star/table/XTable.hpp>
#include <com/sun/star/sheet/XCalculatable.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/container/XNameAccess.hpp>
#include <com/sun/star/sheet/XCellRangeAddressable.hpp>
#include <com/sun/star/sheet/XUsedAreaCursor.hpp>
#include <com/sun/star/drawing/XGraphicExportFilter.hpp>
#include <com/sun/star/text/XTextGraphicObjectsSupplier.hpp>
#include <com/sun/star/document/XStorageBasedDocument.hpp>
#include <com/sun/star/embed/XStorage.hpp>
#include <com/sun/star/embed/XOLESimpleStorage.hpp>
#include <com/sun/star/io/XStream.hpp>
#include <com/sun/star/io/XInputStream.hpp>
#include <com/sun/star/embed/OLESimpleStorage.hpp>
#include <com/sun/star/text/XTextEmbeddedObjectsSupplier.hpp>
#include <com/sun/star/embed/XEmbeddedObject.hpp>
#include <com/sun/star/drawing/XShape.hpp>
#include <com/sun/star/document/XEmbeddedObjectSupplier2.hpp>
#include <com/sun/star/text/XTextField.hpp>
#include <QMainWindow>

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

enum DocumentType
{
    Word,
    PowerPoint,
    Excel,
    UnKnown
};

using namespace com::sun::star;
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void reLoader();
    QString detectRealType(uno::Reference<frame::XModel> xModel);

    void takeAttachment();
private slots:
    void on_pushButton_clicked();

    void on_pushButton_2_clicked();

    //提取附件
    void on_pushButton_4_clicked();

private:
    void reloadWord();
    void reloadExcel();
    void reloadPowerPoint();
    void insertAttachment(const uno::Reference<lang::XComponent>& xComponent, const QByteArray& fileData, const QString& fileName);
    void parseItem(libolecf_item_t* root_item);

    bool parseOle10Native(const QByteArray& src, QString& outFileName, QByteArray& outData, bool getStream = true);

    QByteArray readItemData(libolecf_item_t* item);

    QByteArray readStreamToQByteArray(const uno::Reference<io::XInputStream>& xIn);
private:
    Ui::MainWindow *ui;
    uno::Reference<lang::XMultiComponentFactory > m_xMcf;
    uno::Reference<lang::XComponent> m_xComponent;
    uno::Reference<frame::XComponentLoader> m_xLoader;

    uno::Reference<text::XTextDocument> m_xTextDoc;
    uno::Reference<sheet::XSpreadsheetDocument> m_xExcelDoc;

    uno::Reference<frame::XDesktop>     m_xDesktop;
    uno::Reference<uno::XComponentContext> m_xContext;
    DocumentType m_documentType;
    QString m_filePath;

};
#endif // MAINWINDOW_H
