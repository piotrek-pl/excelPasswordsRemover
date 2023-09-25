#include <QFileDialog>
#include <QDomDocument>
#include <QRegExp>
#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->openFileButton, &QPushButton::clicked,
            this, &MainWindow::openFileButtonClicked);
    connect(ui->removePasswordsButton, &QPushButton::clicked,
            this, &MainWindow::removePasswordsButtonClicked);
    connect(ui->saveAsButton, &QPushButton::clicked,
            this, &MainWindow::saveAsButtonClicked);

    workbookTagNames << "workbookProtection";
    sheetTagNames << "sheetProtection";

    ui->removePasswordsButton->setEnabled(false);
    ui->saveAsButton->setEnabled(false);

    previousSelectedFilePathToRead = "";
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::openFileButtonClicked()
{
    openFile();

    if (!selectedFilePathToRead.isEmpty())
    {
        if (selectedFilePathToRead != previousSelectedFilePathToRead)
        {
            previousSelectedFilePathToRead = selectedFilePathToRead;
            ui->removePasswordsButton->setEnabled(true);
            ui->saveAsButton->setEnabled(false);

            qDebug() << "Selected file for reading: " << selectedFilePathToRead;
        }
    }
}

void MainWindow::openFile()
{
    QString desktopPath = QDir::homePath() + "/Desktop";

    selectedFilePathToRead = QFileDialog::getOpenFileName(
        this,
        "Select a File",
        desktopPath,
        "Excel Files (*.xlsx);"
    );
}

void MainWindow::removePasswordsButtonClicked()
{
    copyAndModifyXlsxFile();

    ui->removePasswordsButton->setEnabled(false);
    ui->saveAsButton->setEnabled(true);
}

void MainWindow::copyAndModifyXlsxFile()
{
    zipReader = QSharedPointer<QXlsx::ZipReader>::create(selectedFilePathToRead);

    tempFilePath = setTempFilePath(selectedFilePathToRead);
    qDebug() << "Temp file path: " << tempFilePath;

    QString targetXmlFileName = "workbook.xml";

    if (zipReader->exists())
    {
        QStringList filePaths = zipReader->filePaths();

        zipWriter = QSharedPointer<QXlsx::ZipWriter>::create(tempFilePath);

        setFileAttributes(tempFilePath, FILE_ATTRIBUTE_HIDDEN);

        qDebug() << "File list:";

        foreach (const QString &filePath, filePaths)
        {
            QFileInfo fileInfo(filePath);
            QString fileName = fileInfo.fileName();

            qDebug() << "\t" << fileName;

            if (fileName.compare(targetXmlFileName, Qt::CaseInsensitive) == 0)
            {
                editAndSaveXmlFile(filePath, workbookTagNames);
                qDebug() << "\t\t" << "File was edited.";
            }
            else if (filePath.endsWith(".xml", Qt::CaseInsensitive) && QRegExp("sheet\\d+", Qt::CaseInsensitive).indexIn(fileName) != -1)
            {
                editAndSaveXmlFile(filePath, sheetTagNames);
            }
            else
            {
                QByteArray fileData = zipReader->fileData(filePath);
                zipWriter->addFile(filePath, fileData);
            }
        }
        zipWriter->close();
        zipReader.clear();
    }
}

QString MainWindow::setTempFilePath(QString filePath)
{
    QFileInfo filePathInfo(filePath);
    QString fileName = filePathInfo.fileName();
    QString fileDirectoryPath = filePathInfo.path();
    QString tempFileName = "temp_" + fileName;
    QString tempFilePath = fileDirectoryPath + "/" + tempFileName;

    return tempFilePath;
}

void MainWindow::editAndSaveXmlFile(QString filePath, QStringList tagNames)
{
    QByteArray xmlData = zipReader->fileData(filePath);
    QString xmlString = QString::fromUtf8(xmlData);
    QString modifiedXmlString = removeXmlNodes(xmlString, tagNames);
    QByteArray newXmlData = modifiedXmlString.toUtf8();

    zipWriter->addFile(filePath, newXmlData);
}

QString MainWindow::removeXmlNodes(QString xmlString, QStringList tagNames)
{
    QDomDocument doc;
    doc.setContent(xmlString);

    foreach (const QString &tagName, tagNames)
    {
        QDomNodeList nodesToRemove = doc.elementsByTagName(tagName);
        for (int i = 0; i < nodesToRemove.size(); ++i)
        {
            QDomNode node = nodesToRemove.at(i);
            node.parentNode().removeChild(node);
        }
    }

    return doc.toString();
}


void MainWindow::saveAsButtonClicked()
{
    QFileInfo filePathInfo(selectedFilePathToRead);
    QString fileDirectoryPath = filePathInfo.path();

    QString fileName = QFileDialog::getSaveFileName(
        this,
        "Save File",
        fileDirectoryPath,
        "Excel Files (*.xlsx);"
        );

    if (fileName.isEmpty())
    {
        return;
    }
    else
    {
        qDebug() << "Selected file for saving: " << fileName;
    }

    QFile sourceFile(tempFilePath);
    if (!sourceFile.exists())
    {
        qDebug() << "Source file doesn't exist: " << tempFilePath;
        return;
    }

    if (tempFilePath == fileName)
    {
        setFileAttributes(fileName, FILE_ATTRIBUTE_NORMAL);
        updateUIAfterSave();
        return;
    }

    QFile targetFile(fileName);
    if (targetFile.exists() && !QFile::remove(fileName))
    {
        qDebug() << "Error while removing existing target file: " << targetFile.errorString();
        return;
    }

    if (!QFile::rename(tempFilePath, fileName))
    {
        qDebug() << "Error while moving the file: " << sourceFile.errorString();
        return;
    }

    setFileAttributes(fileName, FILE_ATTRIBUTE_NORMAL);
    updateUIAfterSave();
}

void MainWindow::setFileAttributes(const QString &filePath, DWORD fileAttributes)
{
    if (SetFileAttributes(reinterpret_cast<const WCHAR*>(filePath.utf16()), fileAttributes))
    {
        qDebug() << "File attributes have been changed:" << filePath;
    }
    else
    {
        qDebug() << "Error while changing file attributes:" << filePath;
    }
}

void MainWindow::updateUIAfterSave()
{
    ui->saveAsButton->setEnabled(false);
    ui->openFileButton->setEnabled(true);
    previousSelectedFilePathToRead = "";
}
