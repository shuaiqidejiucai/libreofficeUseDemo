#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QDebug>
#include <QImage>
#include <QLabel>


class TerminationListener : public cppu::WeakImplHelper1<com::sun::star::frame::XTerminateListener> {
public:
    virtual void SAL_CALL queryTermination(const com::sun::star::lang::EventObject&) override
    {

        // if (atWork)
        // {
        //     std::cout << "Terminate while we are at work? You can't be serious ;-)!\n";
        //     throw TerminationVetoException();
        // }
    }

    virtual void SAL_CALL notifyTermination(const com::sun::star::lang::EventObject&) override {
        qDebug() << "LibreOffice service terminated successfully";
    }

    virtual void SAL_CALL disposing(const com::sun::star::lang::EventObject&) override
    {
        qDebug()<<"disposing";
    }
};

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    com::sun::star::uno::Reference<com::sun::star::uno::XComponentContext> xContext = cppu::bootstrap();
    // 获取服务工厂
    com::sun::star::uno::Reference<com::sun::star::lang::XMultiComponentFactory > xMcf(xContext->getServiceManager());
    uno::Sequence<rtl::OUString > serviceNameQuence = xMcf->getAvailableServiceNames();
    int nLength = serviceNameQuence.getLength();
    rtl::OUString * serviceNameArray = serviceNameQuence.getArray();
    for(int i = 0; i < nLength; ++i)
    {
        rtl::OUString tString = *(serviceNameArray + i);
        const char * ch = rtl::OUStringToOString(tString, RTL_TEXTENCODING_UTF8).getStr();
        qDebug()<< QString::fromUtf8(ch);
    }
    com::sun::star::uno::Reference<com::sun::star::frame::XDesktop> xDesktopT(
        xMcf->createInstanceWithContext("com.sun.star.frame.Desktop", xContext),
        com::sun::star::uno::UNO_QUERY_THROW
        );
    xDesktop = xDesktopT;
    com::sun::star::uno::Reference<com::sun::star::frame::XComponentLoader> xLoaderT(xDesktop, com::sun::star::uno::UNO_QUERY_THROW);
    xLoader = xLoaderT;

    com::sun::star::uno::Reference<com::sun::star::frame::XTerminateListener> xListener(new TerminationListener());
    xDesktop->addTerminateListener(xListener);
}

MainWindow::~MainWindow()
{
    if(xDesktop.is())
    {
        xDesktop->terminate();
    }

    delete ui;
}

void MainWindow::reLoader()
{
    ui->listWidget->clear();
    ui->textEdit->clear();
    com::sun::star::uno::Reference<com::sun::star::text::XText> xText = xTextDoc->getText();
    com::sun::star::uno::Reference<container::XEnumerationAccess> xParaAccess(xText, uno::UNO_QUERY);
    uno::Reference<container::XEnumeration> xParaEnum = xParaAccess->createEnumeration();

    QString qsStrText;
    //遍历段落
    while (xParaEnum->hasMoreElements())
    {
        uno::Reference<text::XTextRange> xPara(xParaEnum->nextElement(), uno::UNO_QUERY);
        rtl::OUString paragraphText = xPara->getString();
        const char * ch = rtl::OUStringToOString(paragraphText, RTL_TEXTENCODING_UTF8).getStr();
        qsStrText += QString::fromUtf8(ch);
    }
    ui->textEdit->setText(qsStrText);
    uno::Reference<drawing::XDrawPageSupplier> xDrawSupplier(xComponent, uno::UNO_QUERY);
    uno::Reference<drawing::XDrawPage> xDrawPage = xDrawSupplier->getDrawPage();

    for (sal_Int32 i = 0; i < xDrawPage->getCount(); ++i)
    {
        uno::Reference<drawing::XShape> xShape(xDrawPage->getByIndex(i), uno::UNO_QUERY);
        if (true) {
            uno::Reference<beans::XPropertySet> xPropSet(xShape, uno::UNO_QUERY);
            uno::Reference<graphic::XGraphic> xGraphic;
            xPropSet->getPropertyValue(rtl::OUString("Graphic")) >>= xGraphic;

            if (xGraphic.is()) {
                uno::Reference<awt::XBitmap> xBitmap(xGraphic, uno::UNO_QUERY);
                if(xBitmap.is())
                {
                    awt::Size size = xBitmap->getSize();
                    QImage image(size.Width, size.Height, QImage::Format_ARGB32);
                    uno::Sequence<::sal_Int8> uInt8 = xBitmap->getDIB();
                    image.loadFromData((uchar*)uInt8.getArray(), size.Width*size.Height * 8);
                    image = image.scaled(50,100,Qt::AspectRatioMode::IgnoreAspectRatio);
                    QListWidgetItem * item = new QListWidgetItem;
                    item->setSizeHint(QSize(50,100));
                    ui->listWidget->addItem(item);
                    QLabel * label = new QLabel();
                    label->setPixmap(QPixmap::fromImage(image));
                    ui->listWidget->setItemWidget(item, label);
                    label->show();
                }
            }
        }
    }

}

void MainWindow::on_pushButton_clicked()
{
    QString qsFileName = QFileDialog::getOpenFileName(this, "choose file",
                                                      "E:/QtProject/wpsfile",
                                                      "*.wps *.doc");
    if(qsFileName.isEmpty())
    {
        return;
    }
    if(xComponent.is())
    {
        uno::Reference<util::XCloseable> xCloseable(xComponent, uno::UNO_QUERY);
        if(xCloseable.is())
        {
            xCloseable->close(true);
        }
    }

    ui->lineEdit->setText(qsFileName);

    qsFileName = "file:///" + qsFileName;
    wchar_t wChFilePath[256] = {0};
    qsFileName.toWCharArray(wChFilePath);
    //rtl::OUString sDocUrl(L"file:///E:/QtProject/wpsfile/testwww.doc");
    rtl::OUString sDocUrl(wChFilePath);
    com::sun::star::uno::Sequence<com::sun::star::beans::PropertyValue> loadProps(1);
    loadProps[0].Name = "Hidden";
    loadProps[0].Value <<= true;
    xComponent = xLoader->loadComponentFromURL(
        sDocUrl, "_blank", 0, loadProps
        );

    com::sun::star::uno::Reference<com::sun::star::text::XTextDocument> xTextDocT(xComponent, com::sun::star::uno::UNO_QUERY_THROW);
    xTextDoc = xTextDocT;

    reLoader();
}


void MainWindow::on_pushButton_2_clicked()
{
    if(!xTextDoc.is())
    {
        return;
    }
    if(false)
    {
        //插入文字
        com::sun::star::uno::Reference<com::sun::star::text::XText> xText = xTextDoc->getText();
        uno::Reference<text::XTextCursor> xCursor = xText->createTextCursor();
        xCursor->gotoStart(false);
        bool ok = xCursor->goRight(2, false);
        xCursor->goRight(2, true);
        uno::Reference<text::XTextRange> textRange = xCursor->getStart();
        if(!textRange.is())
        {
            qDebug()<<"error";
        }
        rtl::OUString xtRtext = xCursor->getString();

        const char * ch2 = rtl::OUStringToOString(xtRtext, RTL_TEXTENCODING_UTF8).getStr();
        qDebug()<<"mjc2"<<QString::fromUtf8(ch2);


        uno::Reference<text::XTextRange> uiuu(xCursor, uno::UNO_QUERY);

        rtl::OUString xtRtext2(L"玉泽");
        xText->insertString(uiuu, xtRtext2, true);
        uno::Reference<frame::XStorable> xStorable(xTextDoc, uno::UNO_QUERY);
        xStorable->store();  // 显式保存
    }


    uno::Reference<util::XSearchable> xSearchable(xTextDoc, uno::UNO_QUERY_THROW);
    uno::Reference<util::XReplaceable> xReplaceable(xTextDoc, uno::UNO_QUERY_THROW);

    QString qsSrcText = ui->lineEdit_2->text();
    QString qsDestText = ui->lineEdit_3->text();

    wchar_t wSrc[1024] = {0};
    wchar_t wDest[1024] = {0};

    qsSrcText.toWCharArray(wSrc);
    qsDestText.toWCharArray(wDest);


    // 执行替换
    uno::Reference<util::XReplaceDescriptor> xReplaceDesc = xReplaceable->createReplaceDescriptor();
    xReplaceDesc->setSearchString(rtl::OUString(wSrc));
    xReplaceDesc->setPropertyValue("SearchCaseSensitive", uno::Any(false));
    xReplaceDesc->setPropertyValue("SearchWords", uno::Any(false));
    xReplaceDesc->setReplaceString(rtl::OUString(wDest));

    uno::Reference<util::XSearchDescriptor> xSearchDesc(xReplaceDesc, uno::UNO_QUERY);
    xReplaceable->replaceAll(xSearchDesc);

    uno::Reference<frame::XStorable> xStorable(xComponent, uno::UNO_QUERY);
    if (xStorable.is())
    {
        xStorable->store();
    }

    reLoader();
}

