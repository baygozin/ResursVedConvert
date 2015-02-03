#include "resurssection.h"

// Правило сортировки по 1-й колонке "Код ресурса"
bool RowCodeLess(const sampleRow a, const sampleRow b) {
    QString codeA = a.cod;
    codeA.remove(QChar('-'), Qt::CaseInsensitive).remove(QChar('_'), Qt::CaseInsensitive);
    QString codeB = b.cod;
    codeB.remove(QChar('-'), Qt::CaseInsensitive).remove(QChar('_'), Qt::CaseInsensitive);
    if (codeA.toLongLong() < codeB.toLongLong()) {return true;} else {return false;}
}

// Правило сортировки по 1-й колонке "Код ресурса"
bool RowSummLess(const sampleRow a, const sampleRow b) {
    double codeA = a.prSum;
    double codeB = b.prSum;
    if (codeA < codeB) {return true;} else {return false;}
}

ResursSection::ResursSection(QObject *parent) : QObject(parent)
{
    srcListRows = QList<sampleRow>();
    dstListRows = QList<sampleRow>();
    rowBegin = 0;
    rowEnd = 0;
}

ResursSection::~ResursSection()
{
}

bool ResursSection::loadData(Document *doc)
{
    QStringList listSheets = doc->sheetNames();
    if (listSheets.count() > 0) {
        foreach (const QString &name, listSheets) {
            if (name.compare("Мои данные",Qt::CaseSensitive) == 0) {
                doc->selectSheet(name);
                workProcess(doc);
                break;
            }
        }
    }
    return true;
}

bool ResursSection::workProcess(Document *doc)
{
    srcListRows.clear();
    Worksheet *ws = doc->currentWorksheet();
    QList<CellRange> listCells;
    listCells.append(ws->mergedCells());
    foreach (const CellRange range, listCells) {
        QString value = ws->cellAt(range.firstRow(), range.firstColumn())->value().toString().trimmed();
        if (!value.compare(strBegin)) {
            rowBegin = range.firstRow() + 1;
        }
        if (rowBegin != 0) {
            break;
        }
    }

    if (rowBegin != 0) {
        for (int i = rowBegin; i < ws->dimension().lastRow(); ++i) {
            QString value = ws->cellAt(i, 2)->value().toString().trimmed();
            if (!value.compare(strEnd)) {
                rowEnd = i;
                break;
            }
        }
        for (int i = rowBegin; i < rowEnd; ++i) {
            sampleRow temp;
            temp.cod = ws->cellAt(i, 1)->value().toString();
            temp.name = ws->cellAt(i, 2)->value().toString();
            temp.measure = ws->cellAt(i, 3)->value().toString();
            temp.amt = ws->cellAt(i, 4)->value().toDouble();
            temp.prOne = ws->cellAt(i, 7)->value().toDouble();
            temp.prSum = ws->cellAt(i, 11)->value().toDouble();
            srcListRows.append(temp);
        }
    }
    listCells.clear();
    return true;
}

// Сортируем по 1-й колонке "Код ресурса"
void ResursSection::sortingRowCode(QList<sampleRow> &list)
{
    qSort(list.begin(), list.end(), RowCodeLess);
}

// Сортируем по 11-й колонке "Общая сметная стоимость"
void ResursSection::sortingRowSumm()
{
    qSort(srcListRows.begin(), srcListRows.end(), RowSummLess);
}

// Ужимаем Трудозатраты
void ResursSection::compactToCode()
{
    dstListRows.clear();
    if (srcListRows.size() <= 0) {
        return;
    }
    sortingRowCode(srcListRows);
    double globalSumm = 0.0;
    sampleRow tmp_row;
    for (int i = 0; i < srcListRows.size(); ++i) {
        globalSumm += srcListRows.at(i).prSum;
    }
    for (int i = 0; i < srcListRows.size(); ++i) {
        dstListRows.append(srcListRows.at(i));
        for (int x = i + 1; x < srcListRows.size(); ++x) {
            if (srcListRows.at(x).cod.compare(dstListRows.last().cod) == 0) {
               summRowCode(&srcListRows.at(x), &dstListRows.last());
            } else {
                i = x - 1;
                break;
            }
        }
    }
    tmp_row.cod = "";
    tmp_row.name = "Итого \"Трудозатраты\"";
    tmp_row.measure = "";
    tmp_row.amt = 0.0;
    tmp_row.prOne = 0.0;
    tmp_row.prSum = globalSumm;
    dstListRows.append(tmp_row);

}

void ResursSection::compactToSummMech(double proc)
{
    dstListRows.clear();
    if (srcListRows.size() <= 0) {
        return;
    }
    sortingRowSumm();
    double summScan = 0.0;
    double globalSumm = 0.0;
    double endsumScan = 0.0;
    sampleRow tmp_row;
    for (int i = 0; i < srcListRows.size(); ++i) {
        globalSumm += srcListRows.at(i).prSum;
    }
    for (int i = 0; i < srcListRows.size(); ++i) {
        summScan += srcListRows.at(i).prSum;
        if (((summScan / globalSumm) * 100) > proc ) {
            dstListRows.append(srcListRows.at(i));
        } else {
            endsumScan = summScan;
        }
    }
    sortingRowCode(dstListRows);
    tmp_row.cod = "";
    tmp_row.name = "Прочие \"Машины и механизмы\"";
    tmp_row.measure = "%";
    tmp_row.amt = proc;
    tmp_row.prOne = 0.0;
    tmp_row.prSum = endsumScan;
    dstListRows.append(tmp_row);
    tmp_row.cod = "";
    tmp_row.name = "Итого \"Машины и механизмы\"";
    tmp_row.measure = "";
    tmp_row.amt = 0.0;
    tmp_row.prOne = 0.0;
    tmp_row.prSum = globalSumm;
    dstListRows.append(tmp_row);
}

void ResursSection::compactToSummMater(double proc)
{
    dstListRows.clear();
    if (srcListRows.size() <= 0) {
        return;
    }
    sortingRowCode(srcListRows);
    double globalSumm = 0.0;
    sampleRow tmp_row;
    for (int i = 0; i < srcListRows.size(); ++i) {
        globalSumm += srcListRows.at(i).prSum;
    }
    for (int i = 0; i < srcListRows.size(); ++i) {
        dstListRows.append(srcListRows.at(i));
        for (int x = i + 1; x < srcListRows.size(); ++x) {
            if (srcListRows.at(x).cod.compare(dstListRows.last().cod) == 0) {
               summRowCodeMaterials(&srcListRows.at(x), &dstListRows.last());
            } else {
                i = x - 1;
                break;
            }
        }
    }
    srcListRows.clear();
    srcListRows = dstListRows;
    dstListRows.clear();
    sortingRowSumm();
    double summScan = 0.0;
    double summTrans = 0.0;
    globalSumm = 0.0;
    double endSumScan = 0.0;
    for (int i = 0; i < srcListRows.size(); ++i) {
        globalSumm += srcListRows.at(i).prSum;
    }
    for (int i = 0; i < srcListRows.size(); ++i) {
        if (srcListRows.at(i).cod.contains("Расчет №") != 0) {
            summTrans += srcListRows.at(i).prSum;
        } else {
            summScan += srcListRows.at(i).prSum;
            if (((summScan / globalSumm) * 100) > proc ) {
                dstListRows.append(srcListRows.at(i));
            } else {
                endSumScan = summScan;
            }
        }
    }
    //endsumScan = summScan;
    sortingRowCode(dstListRows);
    tmp_row.cod = "Расчет";
    tmp_row.name = "Транспортные расходы";
    tmp_row.measure = "";
    tmp_row.amt = 0.0;
    tmp_row.prOne = 0.0;
    tmp_row.prSum = summTrans;
    dstListRows.append(tmp_row);
    tmp_row.cod = "";
    tmp_row.name = "Прочие материалы";
    tmp_row.measure = "%";
    tmp_row.amt = proc;
    tmp_row.prOne = 0.0;
    tmp_row.prSum = endSumScan;
    dstListRows.append(tmp_row);
    tmp_row.cod = "";
    tmp_row.name = "Итого \"Материалы\"";
    tmp_row.measure = "";
    tmp_row.amt = 0.0;
    tmp_row.prOne = 0.0;
    tmp_row.prSum = globalSumm;
    dstListRows.append(tmp_row);

}

void ResursSection::compactToEquipment()
{
    dstListRows.clear();
    if (srcListRows.size() <= 0) {
        return;
    }
    sampleRow tmp_row;
    double summScan = 0.0;
    for (int i = 0; i < srcListRows.size(); ++i) {
        dstListRows.append(srcListRows.at(i));
        summScan += srcListRows.at(i).prSum;
    }
    tmp_row.cod = "";
    tmp_row.name = "Итого \"Оборудование\"";
    tmp_row.measure = "";
    tmp_row.amt = 0.0;
    tmp_row.prOne = 0.0;
    tmp_row.prSum = summScan;
    dstListRows.append(tmp_row);
}

void ResursSection::summRowCode(const sampleRow *src, sampleRow *dst)
{
    dst->cod = dst->cod;
    dst->name = "Затраты труда рабочих...";
    dst->measure = src->measure;
    dst->amt += src->amt;
    dst->prOne = src->prOne;
    dst->prSum += src->prSum;
}

void ResursSection::summRowCodeMaterials(const sampleRow *src, sampleRow *dst)
{
    dst->cod = dst->cod;
    dst->name = src->name;
    dst->measure = src->measure;
    dst->amt += src->amt;
    dst->prOne = src->prOne;
    dst->prSum += src->prSum;
}

