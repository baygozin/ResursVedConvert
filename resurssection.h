#ifndef RESURSSECTION_H
#define RESURSSECTION_H

#include <QObject>
#include "xlsxdocument.h"
#include "bovutils.h"

QTXLSX_USE_NAMESPACE

bool RowCodeLess(const sampleRow a, const sampleRow b);
bool RowSummLess(const sampleRow a, const sampleRow b);

class ResursSection : public QObject
{
    Q_OBJECT
public:
    explicit ResursSection(QObject *parent = 0);
    ~ResursSection();
    QString strBegin, strEnd;
    QList<sampleRow> srcListRows, dstListRows;
    int rowBegin;
    int rowEnd;

    bool loadData(Document *doc);
    bool workProcess(Document *doc);
    void compactToCode();
    void compactToSummMech(double proc);
    void compactToSummMater(double proc);
    void compactToEquipment();
    void summRowCode(const sampleRow *src, sampleRow *dst);
    void summRowCodeMaterials(const sampleRow *src, sampleRow *dst);
    void sortingRowCode(QList<sampleRow> &list); // Сортировка по 1-й колонке "Код ресурса"
    void sortingRowSumm(); // Сортировка по 11-й колонке "Общая сметная стоимость"

};

#endif // RESURSSECTION_H
