#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QMessageBox>
#include <cmath>
#include <iostream>
#include <QCloseEvent>
#include <QDebug>
#include <QDate>
#include <QString>
#include <QFileInfo>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->pbLoad, SIGNAL(clicked()), this, SLOT(onpushbutton()));
    connect(ui->pushButtonSave, SIGNAL(clicked()), this, SLOT(SaveExcel()));
    connect(this, SIGNAL(endInit()), this, SLOT(LoadExcel()));
    setting = new QSettings("ipigaz", "resursvedconvert", this);
    loadconfig();
}

MainWindow::~MainWindow()
{
    delete setting;
    delete ui;
}

void MainWindow::closeEvent(QCloseEvent *event)
{
    QMessageBox::StandardButton mesRepl;
    mesRepl = QMessageBox::question(this, "Заголовок", "Вы действительно хотите выйти из программы?",
                                    QMessageBox::Yes | QMessageBox::No);
    if(mesRepl == QMessageBox::Yes) {
        saveconfig();
        event->accept();
    } else {
        event->ignore();
    }
}

void MainWindow::loadconfig()
{
    ui->dspProcMash->setValue(setting->value("procMash", 4.0).toString().replace(",",".").toDouble());
    ui->dspProcMat->setValue(setting->value("procMat", 4.0).toString().replace(",",".").toDouble());
    ui->dateEdit->setDate(QDate::fromString(setting->value("fromyear", 2014).toString(), "yyyy"));
    ui->textEditObjectName->insertPlainText(setting->value("project", "").toString());
    loadPath = setting->value("loadpath", "").toString();
    savePath = setting->value("savepath", "").toString();
}

void MainWindow::saveconfig()
{
    setting->setValue("procMash", ui->dspProcMash->text());
    setting->setValue("procMat", ui->dspProcMat->text());
    setting->setValue("fromyear", ui->dateEdit->date().year());
    setting->setValue("project", ui->textEditObjectName->toPlainText());
    setting->setValue("loadpath", loadPath);
    setting->setValue("savepath", savePath);
    setting->sync();
}

void MainWindow::onpushbutton()
{
    fileName = QFileDialog::getOpenFileName(this, tr("Open File"), loadPath, tr("Excel File (*.xlsx)"));
    if (!fileName.isEmpty()) {
        QFileInfo loadFileInfo(fileName);
        loadPath = loadFileInfo.absoluteFilePath();
    }
    emit endInit();
}


void MainWindow::LoadExcel()
{

    if (!fileName.isEmpty()) {
        doc = new Document(fileName, this);
        ui->textEdit->clear();

        // Обработка трудозатрат. Сделано!
        laborSection.loadData(doc);
        laborSection.compactToCode();
        foreach (const sampleRow r, laborSection.dstListRows) {
            ui->textEdit->insertPlainText(r.cod.leftJustified(30, ' ') + "\t" + r.name + "\t" + r.measure + "\t" + r.amt + "\t"
                                          + "\t" + QString(" = %1 \t %2 \t %3\n")
                                          .arg(r.amt, 2).arg(r.prOne, 2).arg(r.prSum, 0, 'f', 2));
        }
        ui->textEdit->insertPlainText("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\n");

        // Обработка машин и механизмов. Сделано!
        machineSection.loadData(doc);
        machineSection.compactToSummMech(ui->dspProcMash->value());
        foreach (const sampleRow r, machineSection.dstListRows) {
            ui->textEdit->insertPlainText(r.cod.leftJustified(30, ' ') + "\t" + r.name + "\t" + r.measure + "\t" + r.amt + "\t"
                                          + "\t" + QString(" = %1 \t %2 \t %3\n")
                                          .arg(r.amt, 2).arg(r.prOne, 2).arg(r.prSum, 0, 'f', 2));
        }
        ui->textEdit->insertPlainText("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\n");

        // Обработка материалов. Сделано!
        materialSection.loadData(doc);
        //materialSection.compactToCode();
        materialSection.compactToSummMater(ui->dspProcMat->value());
        foreach (const sampleRow r, materialSection.dstListRows) {
            ui->textEdit->insertPlainText(r.cod.leftJustified(30, ' ') + "\t" + r.name + "\t" + r.measure + "\t" + r.amt + "\t"
                                          + "\t" + QString(" = %1 \t %2 \t %3\n")
                                          .arg(r.amt, 2).arg(r.prOne, 2).arg(r.prSum, 0, 'f', 2));
        }
        ui->textEdit->insertPlainText("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\n");
        // Обработка оборудования
        equipmentSection.loadData(doc);
        equipmentSection.compactToEquipment();
        foreach (const sampleRow r, equipmentSection.dstListRows) {
            ui->textEdit->insertPlainText(r.cod.leftJustified(30, ' ') + "\t" + r.name + "\t" + r.measure + "\t" + r.amt + "\t"
                                          + "\t" + QString(" = %1 \t %2 \t %3\n")
                                          .arg(r.amt, 2).arg(r.prOne, 2).arg(r.prSum, 0, 'f', 2));
        }
    }

}

void MainWindow::SaveExcel()
{
    QString fileName = QFileDialog::getSaveFileName(this, tr("Open File"), savePath, tr("Excel File (*.xlsx)"));
    if (!fileName.isEmpty()) {
        QFileInfo saveFileInfo(fileName);
        savePath = saveFileInfo.absoluteFilePath();
        xlsx = new Document();
        xlsx->addSheet("Выборка ресурсов");
        // Начнем создавать таблицу выборки ресурсов
        setColumnRowDim(xlsx); // Зададим ширину столбцов и высоту строк
        createHeader(xlsx);
        int delta = 14;

        Format fmt0, fmt1, fmt2, fmt3, fmt4, fmt5;
        fmt0.setHorizontalAlignment(Format::AlignHCenter);
        fmt0.setVerticalAlignment(Format::AlignTop);
        fmt0.setBorderStyle(Format::BorderThin);
        fmt0.setFontBold(true);
        fmt1.setHorizontalAlignment(Format::AlignRight);
        fmt1.setVerticalAlignment(Format::AlignTop);
        fmt1.setBorderStyle(Format::BorderThin);
        fmt1.setTextWarp(true);
        fmt1.setNumberFormat("### ### ### ### ##0.00");
        fmt2.setHorizontalAlignment(Format::AlignRight);
        fmt2.setVerticalAlignment(Format::AlignTop);
        fmt2.setFontBold(true);
        fmt2.setBorderStyle(Format::BorderThin);
        fmt2.setTextWarp(true);
        fmt2.setNumberFormat("### ### ### ### ##0.00");
        fmt3.setBorderStyle(Format::BorderThin);
        fmt3.setHorizontalAlignment(Format::AlignLeft);
        fmt3.setVerticalAlignment(Format::AlignTop);
        fmt3.setTextWarp(true);
        fmt4.setBorderStyle(Format::BorderThin);
        fmt4.setHorizontalAlignment(Format::AlignHCenter);
        fmt4.setVerticalAlignment(Format::AlignTop);
        fmt4.setTextWarp(true);
        fmt5.setHorizontalAlignment(Format::AlignLeft);
        fmt5.setVerticalAlignment(Format::AlignTop);
        fmt5.setBorderStyle(Format::BorderThin);
        fmt5.setFontBold(true);
        // Трудозатраты
        xlsx->mergeCells("A" + QString("%1").arg(delta - 1) + ":G" + QString("%1").arg(delta - 1), fmt5);
        xlsx->write(delta - 1, 1, "                    Трудозатраты");
        int x = 0;
        for (int i = 0; i < laborSection.dstListRows.size(); ++i) {
            if (laborSection.dstListRows.size() - 1 != i) {
                xlsx->write(i + delta, 1, i + 1, fmt4);
                xlsx->write(i + delta, 3, laborSection.dstListRows.at(i).name, fmt3);
                xlsx->write(i + delta, 5, laborSection.dstListRows.at(i).amt, fmt1);
                xlsx->write(i + delta, 6, laborSection.dstListRows.at(i).prOne, fmt1);
                xlsx->write(i + delta, 7, laborSection.dstListRows.at(i).prSum, fmt1);
            } else {
                xlsx->write(i + delta, 1, "", fmt0);
                xlsx->write(i + delta, 3, laborSection.dstListRows.at(i).name, fmt2);
                xlsx->write(i + delta, 5, "", fmt0);
                xlsx->write(i + delta, 6, "", fmt0);
                xlsx->write(i + delta, 7, laborSection.dstListRows.at(i).prSum, fmt2);
            }
            xlsx->write(i + delta, 2, laborSection.dstListRows.at(i).cod, fmt1);
            xlsx->write(i + delta, 4, laborSection.dstListRows.at(i).measure, fmt4 );
            x = i;
        }
        x += 1;
        // Машины и механизмы
        int delta1 = delta + x + 1;
        xlsx->mergeCells("A" + QString("%1").arg(delta1 - 1) + ":G" + QString("%1").arg(delta1 - 1), fmt5);
        xlsx->write(delta1 - 1, 1, "                    Машины и механизмы");
        for (int i = 0; i < machineSection.dstListRows.size(); ++i) {
            if (machineSection.dstListRows.size() - 2 > i) {
                xlsx->write(i + delta1, 1, x, fmt4);
                xlsx->write(i + delta1, 3, machineSection.dstListRows.at(i).name, fmt3);
                xlsx->write(i + delta1, 5, machineSection.dstListRows.at(i).amt, fmt1);
                xlsx->write(i + delta1, 6, machineSection.dstListRows.at(i).prOne, fmt1);
                xlsx->write(i + delta1, 7, machineSection.dstListRows.at(i).prSum, fmt1);
                x += 1;
            } else {
                if (machineSection.dstListRows.size() - 1 > i) {
                    xlsx->write(i + delta1, 1, x, fmt4);
                    x += 1;
                } else {
                    xlsx->write(i + delta1, 1, "", fmt0);
                }
                xlsx->write(i + delta1, 3, machineSection.dstListRows.at(i).name, fmt2);
                if (machineSection.dstListRows.at(i).amt != 0) {
                    xlsx->write(i + delta1, 5, machineSection.dstListRows.at(i).amt, fmt1);
                } else {
                    xlsx->write(i + delta1, 5, "", fmt0);
                }
                xlsx->write(i + delta1, 6, "", fmt0);
                xlsx->write(i + delta1, 7, machineSection.dstListRows.at(i).prSum, fmt2);
            }
            xlsx->write(i + delta1, 2, machineSection.dstListRows.at(i).cod, fmt1);
            xlsx->write(i + delta1, 4, machineSection.dstListRows.at(i).measure, fmt4 );
        }
        // Материалы и транспортные
        int delta2 = delta + x + 3;
        xlsx->mergeCells("A" + QString("%1").arg(delta2 - 1) + ":G" + QString("%1").arg(delta2 - 1), fmt5);
        xlsx->write(delta2 - 1, 1, "                    Материалы");
        for (int i = 0; i < materialSection.dstListRows.size(); ++i) {
            if (materialSection.dstListRows.size() - 3 > i) {
                xlsx->write(i + delta2, 1, x, fmt4);
                xlsx->write(i + delta2, 3, materialSection.dstListRows.at(i).name, fmt3);
                xlsx->write(i + delta2, 5, materialSection.dstListRows.at(i).amt, fmt1);
                xlsx->write(i + delta2, 6, materialSection.dstListRows.at(i).prOne, fmt1);
                xlsx->write(i + delta2, 7, materialSection.dstListRows.at(i).prSum, fmt1);
                x += 1;
            } else {
                if (materialSection.dstListRows.at(i).cod.compare("Расчет") == 0) {
                    xlsx->write(i + delta2, 1, x, fmt4);
                    xlsx->write(i + delta2, 3, materialSection.dstListRows.at(i).name, fmt3);
                    xlsx->write(i + delta2, 5, "", fmt1);
                    xlsx->write(i + delta2, 6, "", fmt1);
                    xlsx->write(i + delta2, 7, materialSection.dstListRows.at(i).prSum, fmt1);
                    x += 1;
                } else {
                    if (materialSection.dstListRows.size() - 1 > i) {
                        xlsx->write(i + delta2, 1, x, fmt4);
                        x += 1;
                    } else {
                        xlsx->write(i + delta2, 1, "", fmt0);
                    }
                    xlsx->write(i + delta2, 3, materialSection.dstListRows.at(i).name, fmt2);
                    if (materialSection.dstListRows.at(i).amt != 0) {
                        xlsx->write(i + delta2, 5, materialSection.dstListRows.at(i).amt, fmt1);
                    } else {
                        xlsx->write(i + delta2, 5, "", fmt0);
                    }
                    xlsx->write(i + delta2, 6, "", fmt0);
                    xlsx->write(i + delta2, 7, materialSection.dstListRows.at(i).prSum, fmt2);
                }
            }
            xlsx->write(i + delta2, 2, materialSection.dstListRows.at(i).cod, fmt1);
            xlsx->write(i + delta2, 4, materialSection.dstListRows.at(i).measure, fmt4 );
        }
        // Оборудование. Если этот раздел присутствует...
        if (!equipmentSection.srcListRows.empty()) {
            int delta3 = delta + x + 5;
            xlsx->mergeCells("A" + QString("%1").arg(delta3 - 1) + ":G" + QString("%1").arg(delta3 - 1), fmt5);
            xlsx->write(delta3 - 1, 1, "                    Оборудование");
            for (int i = 0; i < equipmentSection.dstListRows.size(); ++i) {
                if (equipmentSection.dstListRows.size() - 1 != i) {
                    xlsx->write(i + delta3, 1, x, fmt4);
                    xlsx->write(i + delta3, 3, equipmentSection.dstListRows.at(i).name, fmt3);
                    xlsx->write(i + delta3, 5, equipmentSection.dstListRows.at(i).amt, fmt1);
                    xlsx->write(i + delta3, 6, equipmentSection.dstListRows.at(i).prOne, fmt1);
                    xlsx->write(i + delta3, 7, equipmentSection.dstListRows.at(i).prSum, fmt1);
                    x += 1;
                } else {
                    xlsx->write(i + delta3, 1, "", fmt0);
                    xlsx->write(i + delta3, 3, equipmentSection.dstListRows.at(i).name, fmt2);
                    xlsx->write(i + delta3, 5, "", fmt0);
                    xlsx->write(i + delta3, 6, "", fmt0);
                    xlsx->write(i + delta3, 7, equipmentSection.dstListRows.at(i).prSum, fmt2);
                }
                xlsx->write(i + delta3, 2, equipmentSection.dstListRows.at(i).cod, fmt1);
                xlsx->write(i + delta3, 4, equipmentSection.dstListRows.at(i).measure, fmt4 );
            }
        }
        // Итоговая сумма
        double ammountSumm = 0.0;
        int delta4 = 0;
        // Если раздел оборудование есть/нет
        if (equipmentSection.srcListRows.empty()) {
            ammountSumm = laborSection.dstListRows.at(laborSection.dstListRows.size() - 1).prSum +
                machineSection.dstListRows.at(machineSection.dstListRows.size() - 1).prSum +
               materialSection.dstListRows.at(materialSection.dstListRows.size() - 1).prSum;
            delta4 = delta + x + 4;
        } else {
            ammountSumm = laborSection.dstListRows.at(laborSection.dstListRows.size() - 1).prSum +
                machineSection.dstListRows.at(machineSection.dstListRows.size() - 1).prSum +
               materialSection.dstListRows.at(materialSection.dstListRows.size() - 1).prSum +
              equipmentSection.dstListRows.at(equipmentSection.dstListRows.size() - 1).prSum;
            delta4 = delta + x + 6;
        }
        for (int i = 1; i < 8; ++i) {
            xlsx->write(delta4, i, "", fmt0);
        }
        xlsx->write(delta4, 3, "Итого :", fmt2);
        xlsx->write(delta4, 7, ammountSumm, fmt2);
        xlsx->saveAs(fileName);
    }
}

void MainWindow::setColumnRowDim(Document *xlsx) {
    xlsx->setColumnWidth(1,  5.57);
    xlsx->setColumnWidth(2, 17.14);
    xlsx->setColumnWidth(3, 40.57);
    xlsx->setColumnWidth(4,  9.86);
    xlsx->setColumnWidth(5, 11.43);
    xlsx->setColumnWidth(6,  13);
    xlsx->setColumnWidth(7, 23.0);
}

void MainWindow::createHeader(Document *xlsx)
{
    Format fmt0;
    fmt0.setFontSize(14);
    fmt0.setFontName("Arial");
    fmt0.setHorizontalAlignment(Format::AlignHCenter);
    xlsx->mergeCells("C4:E4", fmt0);
    xlsx->setRowHeight(4, 4, 18.0);
    xlsx->write(4, 3, "УКРУПНЕННАЯ ВЫБОРКА РЕСУРСОВ", fmt0);
    fmt0.setFontSize(11);
    fmt0.setFontBold(true);
    fmt0.setVerticalAlignment(Format::AlignVCenter);
    fmt0.setTextWarp(true);
    xlsx->setRowHeight(7, 7, 21.0);
    xlsx->mergeCells("B6:F7", fmt0);
    xlsx->write(6, 2, "ОБЪЕКТ: " + ui->textEditObjectName->toPlainText(), fmt0);
    fmt0.setFontSize(10);
    fmt0.setFontBold(false);
    fmt0.setBorderStyle(Format::BorderThin);
    xlsx->mergeCells("A9:A11", fmt0);
    xlsx->write(9, 1, "№ п/п");
    xlsx->mergeCells("B9:B11", fmt0);
    xlsx->write(9, 2, "Код ресурса");
    xlsx->mergeCells("C9:C11", fmt0);
    xlsx->write(9, 3, "Наименование");
    xlsx->mergeCells("D9:D11", fmt0);
    xlsx->write(9, 4, "Ед.изм.");
    xlsx->mergeCells("E9:E11", fmt0);
    xlsx->write(9, 5, "Кол-во");
    xlsx->mergeCells("F10:F11", fmt0);
    xlsx->write(10, 6, "На единицу");
    xlsx->mergeCells("G10:G11", fmt0);
    xlsx->write(10, 7, "Всего");
    xlsx->mergeCells("F9:G9", fmt0);
    xlsx->write(9, 6, "Стоимость в ценах на " + QString("%1").arg(ui->dateEdit->date().year()) + " г., руб.");
    fmt0.setFontSize(8);
    for (int i = 1; i < 8; ++i) {
        xlsx->write(12, i, i, fmt0);
    }
    xlsx->setRowHeight(9, 9, 30.0);
}
