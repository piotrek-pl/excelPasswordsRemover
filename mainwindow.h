#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QSharedPointer>
#include <xlsxzipreader_p.h>
#include <xlsxzipwriter_p.h>
#include <windows.h>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void openFileButtonClicked();
    void removePasswordsButtonClicked();
    void saveAsButtonClicked();

private:
    Ui::MainWindow *ui;
    QString selectedFilePathToRead;
    QString previousSelectedFilePathToRead;
    QString tempFilePath;
    QStringList workbookTagNames;
    QStringList sheetTagNames;
    QSharedPointer<QXlsx::ZipReader> zipReader;
    QSharedPointer<QXlsx::ZipWriter> zipWriter;

    void openFile();
    void copyAndModifyXlsxFile();
    QString setTempFilePath(QString filePath);
    void editAndSaveXmlFile(QString filePath, QStringList tagNamesToRemove);
    QString removeXmlNodes(QString xmlString, QStringList tagNames);
    void setFileAttributes(const QString &filePath, DWORD fileAttributes);
    void updateUIAfterSave();
};
#endif // MAINWINDOW_H
