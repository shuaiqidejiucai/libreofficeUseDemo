#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <com/sun/star/text/XTextContent.hpp>
#include <quuid.h>
#include <qfile.h>
#include <qfileinfo.h>
#include <qdebug.h>
#include <qfiledialog.h>
#include <libolecf.h>
#include <libbfio_handle.h>
#include <libbfio_memory_range.h>
#include <com/sun/star/container/XNamed.hpp>
#include <com/sun/star/connection/NoConnectException.hpp>
#include <com/sun/star/connection/ConnectionSetupException.hpp>
#include <com/sun/star/bridge/XUnoUrlResolver.hpp>
#include <com/sun/star/embed/XEmbedPersist.hpp>
#include <com/sun/star/io/XObjectInputStream.hpp>
#include <com/sun/star/document/XLinkTargetSupplier.hpp>
#include <com/sun/star/io/IOException.hpp>
#include <com/sun/star/embed/InvalidStorageException.hpp>

QString OUStringToQString(const rtl::OUString& oustring)
{
	const char* ch2 = rtl::OUStringToOString(oustring, RTL_TEXTENCODING_UTF8).getStr();
	return QString::fromUtf8(ch2);
}
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
		qDebug() << "disposing";
	}
};

MainWindow::MainWindow(QWidget* parent)
	: QMainWindow(parent)
	, ui(new Ui::MainWindow)
	, m_documentType(UnKnown)
{
	ui->setupUi(this);

	try
	{
		uno::Reference<uno::XComponentContext> xContext2 = cppu::bootstrap();
		m_xContext = xContext2;
	}
	catch(const uno::Exception& e)
	{
		qDebug() << "cppu::bootstrap() error:" << OUStringToQString(e.Message);
	}
    
	try
	{
		// 获取服务工厂
		uno::Reference<lang::XMultiComponentFactory > xMcf(m_xContext->getServiceManager());
		m_xMcf = xMcf;
	}
	catch (const uno::Exception& e)
	{
		qDebug() << "getServiceManager() error:" << OUStringToQString(e.Message);
	}
	
	try
	{
		uno::Reference<frame::XDesktop> xDesktopT(
			m_xMcf->createInstanceWithContext("com.sun.star.frame.Desktop", m_xContext),
			com::sun::star::uno::UNO_QUERY_THROW
		);
		m_xDesktop = xDesktopT;
	}
	catch (const uno::Exception& e)
	{
		qDebug() << "getServiceManager() error:" << OUStringToQString(e.Message);
	}

	try
	{
		com::sun::star::uno::Reference<com::sun::star::frame::XComponentLoader> xLoaderT(m_xDesktop, com::sun::star::uno::UNO_QUERY_THROW);
		m_xLoader = xLoaderT;
	}
	catch (const uno::Exception& e)
	{
		qDebug() << "XComponentLoader create fail:" << OUStringToQString(e.Message);
	}

	try
	{
		com::sun::star::uno::Reference<com::sun::star::frame::XTerminateListener> xListener(new TerminationListener());
		m_xDesktop->addTerminateListener(xListener);
	}
	catch (const uno::Exception& e)
	{
		qDebug() << "XTerminateListener create fail:" << OUStringToQString(e.Message);
	}
}

MainWindow::~MainWindow()
{
	if (m_xDesktop.is())
	{
		m_xDesktop->terminate();
	}

	delete ui;
}

void MainWindow::reLoader()
{
	ui->listWidget->clear();
	ui->textEdit->clear();

	switch (m_documentType)
	{
	case Word:reloadWord();
		break;
	case PowerPoint:reloadPowerPoint();
		break;
	case Excel:reloadExcel();
		break;
	default:
		break;
	}
}

QString MainWindow::detectRealType(uno::Reference<frame::XModel> xModel) {

	// 获取类型检测服务
	uno::Reference<document::XTypeDetection> xDetector(
		m_xContext->getServiceManager()->createInstanceWithContext(
			"com.sun.star.document.TypeDetection", m_xContext),
		uno::UNO_QUERY_THROW
	);

	// 基于内容而非后缀检测
	rtl::OUString typeName = xDetector->queryTypeByURL(xModel->getURL());
	const char* ch2 = rtl::OUStringToOString(typeName, RTL_TEXTENCODING_UTF8).getStr();
	return QString::fromUtf8(ch2);
}

void MainWindow::takeAttachment()
{
    // QString attachmentPath = m_filePath.replace("/", "\\\\");
    // libolecf_file_t *olecf_file = nullptr;
    // libolecf_item_t *root_item = nullptr;

    // if (libolecf_file_initialize(&olecf_file, nullptr) != 1)
    // {
    //     qCritical() << "Unable to initialize libolecf.";
    //     return;
    // }
    // libolecf_error_t *error = nullptr;
    // if (libolecf_file_open_wide(olecf_file, attachmentPath.toStdWString().c_str(),
    //                        LIBOLECF_OPEN_READ, &error) != 1)
    // {
    //     //qCritical() << "Unable to open OLECF file:" << inputFilePath;

    //     if (error) {
    //         char error_string[1024] = {0};
    //         libolecf_error_backtrace_sprint(error, error_string, sizeof(error_string));
    //         qCritical() << "libolecf error:" << error_string;
    //         libolecf_error_free(&error);
    //     } else {
    //         qCritical() << "libolecf error object is NULL (likely not an OLECF file?)";
    //     }
    //     return;
    // }

    // if (libolecf_file_get_root_item(olecf_file, &root_item, nullptr) != 1)
    // {
    //     qCritical() << "Unable to get root item.";
    //     return;
    // }

    // QStringList fileNameList;
    // parseItem(root_item, fileNameList);
}

void MainWindow::on_pushButton_clicked()
{
	QString qsFileName = QFileDialog::getOpenFileName(this, "choose file",
		"E:/QtProject/wpsfile",
		"*.wps *.doc *.docx *.xlsx *.xls *.et *.ppt *.pptx *.dps");
	if (qsFileName.isEmpty())
	{
		return;
	}
	if (m_xComponent.is())
	{
		uno::Reference<util::XCloseable> xCloseable(m_xComponent, uno::UNO_QUERY);
		if (xCloseable.is())
		{
			xCloseable->close(true);
		}
	}

	ui->lineEdit->setText(qsFileName);
    m_filePath = qsFileName;
	qsFileName = "file:///" + qsFileName;
	wchar_t wChFilePath[256] = { 0 };
	qsFileName.toWCharArray(wChFilePath);
	rtl::OUString sDocUrl(wChFilePath);

    com::sun::star::uno::Sequence<com::sun::star::beans::PropertyValue> loadProps(1);
    loadProps[0].Name = "Hidden";
    loadProps[0].Value <<= true;
	m_xComponent = m_xLoader->loadComponentFromURL(
        sDocUrl, "_blank", 0, loadProps
	);

	uno::Reference<frame::XModel> xModel(m_xComponent, uno::UNO_QUERY_THROW);
	if (!xModel.is())
	{
		return;
	}

	QString realType = detectRealType(xModel);
	if (realType == "writer_MS_Works_Document" || realType == "writer_MS_Word_97")
	{
		m_documentType = Word;
		uno::Reference<text::XTextDocument> xTextDocT(m_xComponent, uno::UNO_QUERY_THROW);

		m_xTextDoc = xTextDocT;
	}
	else if (realType == "impress_MS_PowerPoint_97" || realType == "Office Open XML Presentation")
	{
		m_documentType = PowerPoint;
	}
	else if (realType == "calc_MS_Excel_97" || realType == "MS Excel 2007 XML")
	{
		m_documentType = Excel;
		uno::Reference<sheet::XSpreadsheetDocument> xExcelDocT(m_xComponent, com::sun::star::uno::UNO_QUERY_THROW);
		m_xExcelDoc = xExcelDocT;
	}
	else
	{
		m_documentType = UnKnown;
	}
    reLoader();
}


void MainWindow::on_pushButton_2_clicked()
{
	if (!m_xTextDoc.is())
	{
		return;
	}
	if (false)
	{
		//插入文字
		com::sun::star::uno::Reference<com::sun::star::text::XText> xText = m_xTextDoc->getText();
		uno::Reference<text::XTextCursor> xCursor = xText->createTextCursor();
		xCursor->gotoStart(false);
		bool ok = xCursor->goRight(2, false);
		xCursor->goRight(2, true);
		uno::Reference<text::XTextRange> textRange = xCursor->getStart();
		if (!textRange.is())
		{
			qDebug() << "error";
		}
		rtl::OUString xtRtext = xCursor->getString();

		const char* ch2 = rtl::OUStringToOString(xtRtext, RTL_TEXTENCODING_UTF8).getStr();
		qDebug() << "mjc2" << QString::fromUtf8(ch2);


		uno::Reference<text::XTextRange> uiuu(xCursor, uno::UNO_QUERY);

		rtl::OUString xtRtext2(L"玉泽");
		xText->insertString(uiuu, xtRtext2, true);
		uno::Reference<frame::XStorable> xStorable(m_xTextDoc, uno::UNO_QUERY);
		xStorable->store();  // 显式保存
	}


	uno::Reference<util::XSearchable> xSearchable(m_xTextDoc, uno::UNO_QUERY_THROW);
	uno::Reference<util::XReplaceable> xReplaceable(m_xTextDoc, uno::UNO_QUERY_THROW);

	QString qsSrcText = ui->lineEdit_2->text();
	QString qsDestText = ui->lineEdit_3->text();

	wchar_t wSrc[1024] = { 0 };
	wchar_t wDest[1024] = { 0 };

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

	uno::Reference<container::XIndexAccess> xResults = xSearchable->findAll(xSearchDesc);
	int resultCount = xResults->getCount();

	uno::Reference<frame::XStorable> xStorable(m_xComponent, uno::UNO_QUERY);
	if (xStorable.is())
	{
		xStorable->store();
	}
    reLoader();
}

void MainWindow::reloadWord()
{
	com::sun::star::uno::Reference<com::sun::star::text::XText> xText = m_xTextDoc->getText();
	com::sun::star::uno::Reference<container::XEnumerationAccess> xParaAccess(xText, uno::UNO_QUERY);
	uno::Reference<container::XEnumeration> xParaEnum = xParaAccess->createEnumeration();

	QString qsStrText;
	//遍历段落
	while (xParaEnum->hasMoreElements())
	{
		uno::Reference<text::XTextRange> xPara(xParaEnum->nextElement(), uno::UNO_QUERY);
		if (!xPara.is())
		{
			break;
		}
		rtl::OUString paragraphText = xPara->getString();
		const char* ch = rtl::OUStringToOString(paragraphText, RTL_TEXTENCODING_UTF8).getStr();
		qsStrText += QString::fromUtf8(ch);
	}
	ui->textEdit->setText(qsStrText);
	uno::Reference<text::XTextGraphicObjectsSupplier> obj (m_xComponent, uno::UNO_QUERY);
	uno::Reference<container::XNameAccess> nameAccess = obj->getGraphicObjects();
	uno::Sequence<rtl::OUString> eleNames = nameAccess->getElementNames();
	int eleCount = eleNames.getLength();
	rtl::OUString* nameStr = eleNames.getArray();
	for (int eleIndex = 0; eleIndex < eleCount; ++eleIndex)
	{
		rtl::OUString nameStrObj = nameStr[eleIndex];
		uno::Any varant = nameAccess->getByName(nameStrObj);
		uno::Reference<drawing::XShape> xShape(varant, uno::UNO_QUERY);
		uno::Reference<beans::XPropertySet> xPropSet(xShape, uno::UNO_QUERY);

		uno::Any graphicAny = xPropSet->getPropertyValue("Graphic");
		uno::Reference<graphic::XGraphic> xGraphic;
		graphicAny >>= xGraphic;
		uno::Reference<awt::XBitmap> xBitmap(xGraphic, uno::UNO_QUERY);
		if (xBitmap.is())
		{
			awt::Size size = xBitmap->getSize();
			QImage image(size.Width, size.Height, QImage::Format_ARGB32);
			uno::Sequence<::sal_Int8> uInt8 = xBitmap->getDIB();
			image.loadFromData((uchar*)uInt8.getArray(), size.Width * size.Height * 8);
			image = image.scaled(50, 100, Qt::AspectRatioMode::IgnoreAspectRatio);
			QListWidgetItem* item = new QListWidgetItem;
			item->setSizeHint(QSize(50, 100));
			ui->listWidget->addItem(item);
			QLabel* label = new QLabel();
			label->setPixmap(QPixmap::fromImage(image));
			ui->listWidget->setItemWidget(item, label);
			label->show();
		}
		qDebug() << OUStringToQString(nameStrObj);
	}

	ui->listWidget_2->clear();



	libolecf_file_t* olecf_file = nullptr;
	//遍历附件
	if (!olecf_file)
	{
		if (libolecf_file_initialize(&olecf_file, NULL) != 1)
		{
			qCritical() << "Unable to initialize libolecf.";
			return;
		}
	}
	libolecf_item_t* root_item = NULL;
	QString attachment = m_filePath.replace("/", "\\\\");
	libolecf_error_t* error = NULL;
	if (libolecf_file_open(olecf_file, attachment.toUtf8().data(), LIBOLECF_OPEN_READ, &error) != 1) {
		qCritical() << "Unable to open OLECF file:" << attachment;

		if (error) {
			char error_string[1024] = { 0 };
			libolecf_error_backtrace_sprint(error, error_string, sizeof(error_string));
			qCritical() << "libolecf error:" << error_string;
			libolecf_error_free(&error);
		}
		else {
			qCritical() << "libolecf error object is NULL (likely not an OLECF file?)";
		}
		return;
	}

	if (libolecf_file_get_root_item(olecf_file, &root_item, NULL) != 1) {
		qCritical() << "Unable to get root item.";
		return;
	}

	QHash<QString, QByteArray> oleFileHash;
	parseItem(root_item, oleFileHash);
	// 遍历子项，查找 Ole10Native 流
	// 清理
	libolecf_item_free(&root_item, NULL);
	libolecf_file_free(&olecf_file, NULL);
	olecf_file = nullptr;
	QHash<QString, QByteArray>::iterator it;

	for (it = oleFileHash.begin(); it != oleFileHash.end(); ++it)
	{
		QListWidgetItem* item = new QListWidgetItem(ui->listWidget_2);
		item->setText(it.key());
		item->setData(Qt::UserRole + 1, it.value());
		ui->listWidget_2->addItem(item);
	}
}

void MainWindow::reloadExcel()
{
	uno::Reference<sheet::XSpreadsheets> excleSheets = m_xExcelDoc->getSheets();
	uno::Reference <container::XNameAccess> nameAccess(excleSheets, uno::UNO_QUERY);
	uno::Sequence<rtl::OUString > ousStringQuence = nameAccess->getElementNames();
	rtl::OUString* oustringPtr = ousStringQuence.getArray();
	int arryCount = ousStringQuence.getLength();
	for (int i = 0; i < arryCount; ++i)
	{
		rtl::OUString tempStr = *(oustringPtr + i);
		uno::Reference<sheet::XSpreadsheet>  excleSheet;
		excleSheets->getByName(tempStr) >>= excleSheet;
		if (excleSheet.is())
		{
			// 获取实际使用区域
			uno::Reference<sheet::XSheetCellCursor> xCursor = excleSheet->createCursor();
			uno::Reference<sheet::XUsedAreaCursor> xUsedCursor(xCursor, uno::UNO_QUERY);
			if (xUsedCursor.is()) {
				xUsedCursor->gotoStartOfUsedArea(false);
				xUsedCursor->gotoEndOfUsedArea(true);

				uno::Reference<table::XCellRange> xRange(xUsedCursor, uno::UNO_QUERY);
				table::CellRangeAddress rangeAddr = uno::Reference<sheet::XCellRangeAddressable>(xRange, uno::UNO_QUERY)->getRangeAddress();

				int rowCount = rangeAddr.EndRow - rangeAddr.StartRow + 1;
				int colCount = rangeAddr.EndColumn - rangeAddr.StartColumn + 1;

				for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
				{
					for (int columnIndex = 0; columnIndex < colCount; ++columnIndex)
					{
						uno::Reference<table::XCell > xCell = xRange->getCellByPosition(columnIndex, rowIndex);
						uno::Reference<text::XText> cellText(xCell, uno::UNO_QUERY);
						if (cellText.is())
						{
							rtl::OUString xtRtext = cellText->getString();
							const char* ch2 = rtl::OUStringToOString(xtRtext, RTL_TEXTENCODING_UTF8).getStr();
							QString::fromUtf8(ch2);
						}
					}
				}
			}
		}
	}
}

void MainWindow::reloadPowerPoint()
{

}

int littleToBigByte()
{
	return 0;
}

void MainWindow::insertAttachment(const uno::Reference<lang::XComponent>& xComponent, const QByteArray& fileData, const QString& fileName)
{
	QFile file(m_filePath);
	if (file.open(QIODevice::ReadOnly))
	{
		QByteArray data = file.readAll();
		// 1. 获取文档的多服务工厂
		uno::Reference<lang::XMultiServiceFactory> xFactory(xComponent, uno::UNO_QUERY);
		if (!xFactory.is()) return;
		// 2. 创建嵌入对象
		rtl::OUString objectName("EmbeddedObject");
		uno::Reference<beans::XPropertySet> xEmbeddedProps(
			xFactory->createInstance("com.sun.star.text.TextEmbeddedObject"),
			uno::UNO_QUERY);

		if (!xEmbeddedProps.is()) return;

		uno::Reference<document::XStorageBasedDocument> xStorageDoc(xComponent, uno::UNO_QUERY);
		if (!xStorageDoc.is()) return;

		uno::Reference<embed::XStorage> xDocStorage = xStorageDoc->getDocumentStorage();
		if (!xDocStorage.is()) return;
		// 6. 创建新的存储
		QString qsUuid = QUuid::createUuid().toString();
		rtl::OUString ouUUID = rtl::OUString::createFromAscii(qsUuid.toUtf8().data());

		uno::Reference<embed::XStorage> xObjStorage =
			xDocStorage->openStorageElement(ouUUID,
				embed::ElementModes::READWRITE);

		if (!xObjStorage.is()) return;

		// 7. 创建 Ole10Native 流
		uno::Reference<io::XStream> xOleStream =
			xObjStorage->openStreamElement(
				rtl::OUString::createFromAscii("Ole10Native"),
				embed::ElementModes::READWRITE);

		if (!xOleStream.is()) return;
		// 8. 写入 Ole10Native 数据
		uno::Reference<io::XOutputStream> xOutput = xOleStream->getOutputStream();
		if (xOutput.is())
		{
			// 构建 Ole10Native 格式数据
			QString fileTempPath;
			QByteArray oleData;
			// 总长度（4字节）
			QFileInfo info(m_filePath);
			QString fileNameS = info.fileName();
			quint32 totalSize = 4 + 2 + fileNameS.toUtf8().size() + 1 + info.absoluteFilePath().toUtf8().size() +
				1 + 4 + 4 + fileTempPath.toUtf8().size() + 1 + 4 + data.size();

			oleData.append((char*)&totalSize, 4);
			//保留字节
			oleData.append('\x2');
			oleData.append('\0');
			// 文件名
			QByteArray fileNameUtf8 = fileNameS.toUtf8();
			oleData.append(fileNameUtf8);
			oleData.append('\0');
			//全路径
			QByteArray fileNameFullUtf8 = info.absoluteFilePath().toUtf8();
			oleData.append(fileNameFullUtf8);
			oleData.append('\0');
			//mark
			oleData.append('\0');
			oleData.append('\0');
			oleData.append('\x3');
			oleData.append('\0');
			//临时路径size
			int fileTempSize = fileTempPath.toUtf8().size();
			oleData.append((char*)&fileTempSize, 4);
			//临时路径
			QByteArray fileNameTempUtf8 = fileTempPath.toUtf8();
			oleData.append(fileNameTempUtf8);
			oleData.append('\0');
			// 添加文件大小
			quint32 fileSize = data.size();
			oleData.append((char*)&fileSize, 4);

			// 添加文件数据
			oleData.append(data);

			// 写入数据
			uno::Sequence<sal_Int8> buffer(
				reinterpret_cast<const sal_Int8*>(oleData.constData()),
				oleData.size());
			xOutput->writeBytes(buffer);
			xOutput->flush();
		}
		// 9. 提交存储
		uno::Reference<embed::XTransactedObject> xTransact(xObjStorage, uno::UNO_QUERY);
		if (xTransact.is())
		{
			xTransact->commit();
		}
		// 10. 插入到文档中
		uno::Reference<text::XTextDocument> xTextDoc(xComponent, uno::UNO_QUERY);
		if (xTextDoc.is())
		{
			uno::Reference<text::XText> xText = xTextDoc->getText();
			uno::Reference<text::XTextCursor> xCursor = xText->createTextCursor();

			// 移动到文档末尾
			xCursor->gotoEnd(false);

			uno::Reference<text::XTextContent> xContent(xEmbeddedProps, uno::UNO_QUERY);
			if (xContent.is())
			{
				xText->insertTextContent(xCursor->getEnd(), xContent, false);
			}
			// 11. 保存文档
			uno::Reference<frame::XStorable> xStorable(xComponent, uno::UNO_QUERY);
			if (xStorable.is())
			{
				xStorable->store();
			}
		}
		file.close();
	}
}

void outputPropertySet(uno::Reference<beans::XPropertySet> propertySet)
{
	if (propertySet.is())
	{
		uno::Reference<beans::XPropertySetInfo> setInfo = propertySet->getPropertySetInfo();
		if (setInfo.is())
		{
			uno::Sequence<beans::Property> propertyQuence = setInfo->getProperties();
			int length = propertyQuence.getLength();
			beans::Property* properArray = propertyQuence.getArray();
			for (int i = 0; i < length; ++i)
			{
				beans::Property proper = properArray[i];
				qDebug() << OUStringToQString(proper.Name);
			}
		}
	}
}

void outNameContainter(uno::Reference<container::XNameContainer> xContainer3)
{
	uno::Sequence<rtl::OUString> containter3EleNames = xContainer3->getElementNames();
	int containter3Count = containter3EleNames.getLength();
	rtl::OUString* containter3Str = containter3EleNames.getArray();
	for (int p = 0; p < containter3Count; ++p)
	{
		rtl::OUString container3Str = containter3Str[p];
		qDebug() << OUStringToQString(container3Str);
		uno::Any container3Any = xContainer3->getByName(container3Str);
		qDebug() << OUStringToQString(container3Any.getValueTypeName());
	}
}

bool MainWindow::attachmentName(const QByteArray& srcData, QString &fileName)
{
	bool successful = false;
	QByteArray buffer = srcData;
	libbfio_handle_t* bfio_handle = nullptr;
	libcerror_error_t* bfio_error = nullptr;

	libcerror_error_t* rangeBfio_error = nullptr;
	if (libbfio_memory_range_initialize(&bfio_handle, &rangeBfio_error) != 1)
	{
		return successful;
	}

	// 2. 设置内存数据
	if (libbfio_memory_range_set(
		bfio_handle,
		reinterpret_cast<uint8_t*>(buffer.data()),
		buffer.size(),
		&bfio_error) != 1) {
		libbfio_handle_free(&bfio_handle, nullptr);
		return successful;
	}

	// 初始化 libolecf 对象
	libolecf_error_t* error = nullptr;
	libolecf_file_t* olecf2_file = nullptr;
	if (libolecf_file_initialize(&olecf2_file, nullptr) != 1)
	{
		qCritical() << "Unable to initialize libolecf.";
		libbfio_handle_free(&bfio_handle, nullptr);
		return successful;
	}

	// 使用内存句柄打开 OLECF
	if (libolecf_file_open_file_io_handle(
		olecf2_file,
		bfio_handle,
		LIBOLECF_OPEN_READ,
		&error) != 1)
	{
		qCritical() << "Unable to open OLECF from memory.";
		libolecf_file_free(&olecf2_file, nullptr);
		libbfio_handle_free(&bfio_handle, nullptr);
		return successful;
	}
	// 获取根项并解析
	libolecf_item_t* root_item = nullptr;
	if (libolecf_file_get_root_item(olecf2_file, &root_item, nullptr) == 1)
	{
		getFileName(root_item, fileName);
		if (!fileName.isEmpty())
		{
			successful = true;
		}

		libolecf_item_free(&root_item, nullptr);
	}

	// 清理资源
	libolecf_file_free(&olecf2_file, nullptr);
	libbfio_handle_free(&bfio_handle, nullptr);
	return successful;
}

void MainWindow::removeAttachment(const QString& name)
{
	uno::Reference<document::XStorageBasedDocument> xDocBaseStorage(m_xTextDoc, uno::UNO_QUERY_THROW);
	uno::Reference<embed::XStorage> xDocStorage = xDocBaseStorage->getDocumentStorage();
	uno::Reference<text::XTextEmbeddedObjectsSupplier> xSupplier(m_xTextDoc, uno::UNO_QUERY);
	uno::Reference<container::XNameAccess> xEmbeddedObjects = xSupplier->getEmbeddedObjects();

	if (xEmbeddedObjects.is())
	{
		uno::Sequence<rtl::OUString> elementNameQuence = xEmbeddedObjects->getElementNames();
		int length = elementNameQuence.getLength();
		rtl::OUString* elementNameArray = elementNameQuence.getArray();
		for (int i = 0; i < length; ++i)
		{
			rtl::OUString elementName = elementNameArray[i];
			uno::Any any = xEmbeddedObjects->getByName(elementName);
			uno::Reference<beans::XPropertySet> xProps(any, uno::UNO_QUERY);
			rtl::OUString streamName;
			xProps->getPropertyValue("StreamName") >>= streamName;
			uno::Reference<io::XStream> sStream;
			if (xDocStorage->isStorageElement(streamName))
			{
				xDocStorage->openStorageElement(streamName, embed::ElementModes::READWRITE);
			}
			
			if (xDocStorage->isStreamElement(streamName))
			{
				try
				{
					sStream = xDocStorage->cloneStreamElement(streamName);
				}
				catch(const uno::Exception& e)
				{
					qDebug() << OUStringToQString(e.Message);
				}
			}

			if (sStream.is())
			{
				uno::Reference<io::XInputStream> inputStream = sStream->getInputStream();
				if (inputStream.is())
				{
					QByteArray buffer = readStreamToQByteArray(inputStream);
					QString fileName;
					bool successful = attachmentName(buffer, fileName);
					if (successful && fileName == name)
					{
						uno::Reference<text::XTextContent> xContent(any, uno::UNO_QUERY);
						if (xContent.is())
						{
							removeStream(fileName);
							uno::Reference<text::XText> xText = m_xTextDoc->getText();
							xText->removeTextContent(xContent);
							uno::Reference<frame::XStorable> xStorable(m_xComponent, uno::UNO_QUERY);
							if (xStorable.is())
							{
								xStorable->store();
							}
						}
					}
				}
			}
		}
	}
}

void MainWindow::removeStream(const QString& name)
{
	uno::Reference<document::XStorageBasedDocument> xDocBaseStorage(m_xTextDoc, uno::UNO_QUERY_THROW);
	uno::Reference<embed::XStorage> xDocStorage = xDocBaseStorage->getDocumentStorage();
	uno::Sequence<rtl::OUString> elementQunence = xDocStorage->getElementNames();
	int length = elementQunence.getLength();
	rtl::OUString* ustringArray = elementQunence.getArray();
	for (int i = 0; i < length; ++i)
	{
		rtl::OUString ustring = ustringArray[i];
		uno::Reference<io::XStream> xstream;
		try
		{
			if (xDocStorage->isStorageElement(ustring))
			{
				//TODO:
			}
			else if(xDocStorage->isStreamElement(ustring))
			{
				xstream = xDocStorage->openStreamElement(ustring, embed::ElementModes::READWRITE);
			}
			if (xstream.is())
			{
				uno::Reference<io::XInputStream> inputStream = xstream->getInputStream();
				if (inputStream.is())
				{
					QByteArray buffer = readStreamToQByteArray(inputStream);
					QString fileName;
					bool successful = attachmentName(buffer, fileName);
					if (successful && fileName == name)
					{
						xDocStorage->removeElement(ustring);
						uno::Reference<embed::XTransactedObject> xTransact(xDocStorage, uno::UNO_QUERY);
						if (xTransact.is())
						{
							xTransact->commit();
						}
					}
				}
			}
		}
		catch (const uno::Exception& e)
		{
			continue;
		}
		
	}
}

void MainWindow::parseItem(libolecf_item_t *root_item, QHash<QString, QByteArray>& oleFileHash)
{
	int number_of_sub_items = 0;
	libolecf_item_get_number_of_sub_items(root_item, &number_of_sub_items, NULL);
	for (int i = 0; i < number_of_sub_items; ++i)
	{
		libolecf_item_t* sub_item = NULL;
		libolecf_item_get_sub_item(root_item, i, &sub_item, NULL);
		int childItemCount = 0;
		libolecf_item_get_number_of_sub_items(sub_item, &childItemCount, NULL);
		size_t name_size = 0;
		libolecf_item_get_utf8_name_size(sub_item, &name_size, NULL);
		char* name = new char[name_size];
		libolecf_item_get_utf8_name(sub_item, (uint8_t*)name, name_size, NULL);
		QString itemName(name);
		delete[] name;
		if (itemName.endsWith("Ole10Native"))
		{
			QByteArray ole10 = readItemData(sub_item);
			QString fileName;
			QByteArray fileData;
			if (parseOle10Native(ole10, fileName, fileData))
			{
				oleFileHash[fileName] = fileData;
			}
		}
		if (childItemCount > 0)
		{
			parseItem(sub_item, oleFileHash);
		}
		libolecf_item_free(&sub_item, nullptr);
	}
}

void MainWindow::getFileName(libolecf_item_t* root_item, QString& rootName)
{
	int number_of_sub_items = 0;
	libolecf_item_get_number_of_sub_items(root_item, &number_of_sub_items, NULL);
	for (int i = 0; i < number_of_sub_items; ++i)
	{
		libolecf_item_t* sub_item = NULL;
		libolecf_item_get_sub_item(root_item, i, &sub_item, NULL);
		int childItemCount = 0;
		libolecf_item_get_number_of_sub_items(sub_item, &childItemCount, NULL);
		size_t name_size = 0;
		libolecf_item_get_utf8_name_size(sub_item, &name_size, NULL);
		char* name = new char[name_size];
		libolecf_item_get_utf8_name(sub_item, (uint8_t*)name, name_size, NULL);
		QString itemName(name);
		delete[] name;
		if (itemName.endsWith("Ole10Native"))
		{
			QByteArray ole10 = readItemData(sub_item);
			QString fileName;
			QByteArray fileData;
			if (parseOle10Native(ole10, fileName, fileData))
			{
				rootName = fileName;
				libolecf_item_free(&sub_item, nullptr);
				return;
			}
		}
		if (childItemCount > 0)
		{
			getFileName(sub_item, rootName);
		}
		libolecf_item_free(&sub_item, nullptr);
	}
	return;
}

bool MainWindow::parseOle10Native(const QByteArray &src, QString &outFileName, QByteArray &outData, bool getStream)
{
    outFileName.clear();
    outData.clear();
    if (src.size() < 12) return false;

    auto rdLE32 = [&](int off)->quint32 {
        if (off + 4 > src.size()) return 0;
        const uchar* p = reinterpret_cast<const uchar*>(src.constData() + off);
        return (quint32)p[0] | ((quint32)p[1] << 8) | ((quint32)p[2] << 16) | ((quint32)p[3] << 24);
    };

    quint32 totalSize = rdLE32(0);
    int off = 4;

    auto readZ = [&](int& pos)->QByteArray {
        int start = pos;
        while(src[pos] == '\0')
        {
            ++pos;
        }
        while (pos < src.size() && src[pos] != '\0') ++pos;
        if (pos >= src.size()) return {};
        QByteArray s = src.mid(start, pos - start);
        ++pos;
        return s;
    };

    bool ok = false;
    for (int tryOff : {4}) {
        if (tryOff >= src.size()) break;
        int p = tryOff;
        QByteArray label = readZ(p);
        if (label.isEmpty()) continue;

        QByteArray file = readZ(p);
        outFileName = file;
        if (file.isEmpty()) continue;

        QByteArray cmd = readZ(p);
        if (cmd.isEmpty()) continue;

        quint32 mark = rdLE32(p);
        p = p + 4;
        quint32 tempPathLen = rdLE32(p);
        p = p + 4;
        QByteArray tempPathBa = readZ(p);
        if (tempPathBa.isEmpty()) continue;

        quint32 datasize = rdLE32(p);
        p = p + 4;

        QByteArray srcData = src.mid(p, datasize);
        outData = srcData;
        ok = true;
    }

    return ok;
}

QByteArray MainWindow::readItemData(libolecf_item_t* item)
{
	uint32_t size = 0;
	if (libolecf_item_get_size(item, &size, NULL) != 1 || size == 0)
		return {};

	QByteArray buf;
	buf.resize((int)size);

	ssize_t read_count = libolecf_stream_read_buffer(
		item,
		reinterpret_cast<uint8_t*>(buf.data()),
		(size_t)size,
		NULL
	);
	if (read_count < 0) return {};
	buf.resize((int)read_count);
	return buf;
}

QByteArray MainWindow::readStreamToQByteArray(const uno::Reference<io::XInputStream>& xIn)
{
	QByteArray result;
	if (!xIn.is()) return result;

	const sal_Int32 bufSize = 4096;
	uno::Sequence<sal_Int8> buffer(bufSize);

	while (true) 
	{
		sal_Int32 nRead = xIn->readBytes(buffer, bufSize);
		if (nRead <= 0) break;
		result.append(reinterpret_cast<const char*>(buffer.getConstArray()), nRead);
	}
	int size = result.size();
	return result;
}

void MainWindow::on_pushButton_4_clicked()
{
	QString qsFilePath = QFileDialog::getExistingDirectory(nullptr, "chioce Directory", "");
	if (qsFilePath.isEmpty())
	{
		return;
	}
	
	QList<QListWidgetItem*> selectItemList = ui->listWidget_2->selectedItems();
	for (int i = 0; i < selectItemList.count(); ++i)
	{
		QListWidgetItem* item = selectItemList.at(i);
		QString fileName = item->text();
		QByteArray data = item->data(Qt::UserRole + 1).toByteArray();
		QString qsFilePath2 = qsFilePath + "/" + fileName;
		QFile file(qsFilePath2);
		if (file.open(QIODevice::WriteOnly))
		{
			file.write(data);
			file.close();
		}
	}
}


void MainWindow::on_pushButton_5_clicked()
{
	QList<QListWidgetItem*> selectItemList = ui->listWidget_2->selectedItems();
	for (int i = 0; i < selectItemList.count(); ++i)
	{
		QListWidgetItem* item = selectItemList.at(i);
		QString fileName = item->text();
		removeAttachment(fileName);
	}
	reLoader();
}

