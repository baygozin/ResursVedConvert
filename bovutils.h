#ifndef BOVUTILS
#define BOVUTILS
#include <QtCore>

struct sampleRow {
    QString     cod;        // Код ресурса (1)
    QString     name;       // Наименование (2)
    QString     measure;    // Ед.измерения (3)
    double      amt;        // Кол-во (4)
    double      prOne;      // Текущая сметная цена за единицу (7)
    double      prSum;      // Текущая сметная цена общая (11)
};

#endif // BOVUTILS

