#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QSettings>
#include "xlsxdocument.h"
#include "laborman.h"
#include "machine.h"
#include "materials.h"
#include "equipment.h"

QTXLSX_USE_NAMESPACE

namespace Ui {
    class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
    LaborMan laborSection;
    Machine machineSection;
    Materials materialSection;
    Equipment equipmentSection;

private:
    Ui::MainWindow *ui;
    Document *doc;
    Document *xlsx;
    QString fileName;

    void loadconfig();
    void saveconfig();
    void setColumnRowDim(Document *xlsx);
    void createHeader(Document *xlsx);

private slots:
    void onpushbutton();
    void LoadExcel();
    void SaveExcel();
signals:
    void endInit();
};

#endif // MAINWINDOW_H
