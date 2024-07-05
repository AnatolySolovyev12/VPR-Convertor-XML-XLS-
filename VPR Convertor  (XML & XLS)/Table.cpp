#include "Table.h"
#include <QInputDialog>
#include <QElapsedTimer>
#include <QAxWidget>
#include <QTime>
#include <QMultiHash>
#include <QFile>

#include <QPair.h>

QTextStream out(stdout);

Table::Table(QWidget* parent)
    : QWidget(parent) {

    QHBoxLayout* Hbox = new QHBoxLayout(this);
    Vbox = new QVBoxLayout();
    QVBoxLayout* VboxButtons = new QVBoxLayout();

    VPR = new QPushButton("VPR", this);
   // VPR->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    connect(VPR, &QPushButton::clicked, this, &Table::myVPR);

    buttConvertToXML = new QPushButton("Convert Donor to XML", this);
    // VPR->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    connect(buttConvertToXML, &QPushButton::clicked, this, &Table::funcConvertToXML);

    donor = new QPushButton("AddDonor", this);
    //donor->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    connect(donor, &QPushButton::clicked, this, &Table::addDonor);

    recepient = new QPushButton("AddRecepient", this);
   // recepient->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    connect(recepient, &QPushButton::clicked, this, &Table::addRecepient);

    loadConfig = new QPushButton("Load config", this);
    // recepient->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    connect(loadConfig, &QPushButton::clicked, this, &Table::readFileConfig);

    paramMenu = new QPushButton("Selecting Options", this);
   // paramMenu->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    pm = new QMenu(paramMenu); // Инициализируем выпадающую кнопку

    pm->addAction("&Where find in Donor?", this, &Table::whatFind);
    pm->addAction("&What start Row find in Donor?", this, &Table::RowDoctor);
    pm->addAction("&Where find in Recepient?", this, &Table::whereFind);
    pm->addAction("&What start Row find in Recepient?", this, &Table::RowRecepient);
    pm->addAction("&Where Day/Night of Donor?", this, &Table::whereDayNightDonor);
    pm->addAction("&Where Day/Night of Recepient?", this, &Table::whereDayNightRecepient);
    pm->addAction("&What to insert in Donor?", this, &Table::whatToInsert);
    pm->addAction("&Where to insert in Recepient?", this, &Table::whereToInsert);
    pm->addAction("&Indent from last line with text in Donor?", this, &Table::lastLineInDonor);
    pm->addAction("&Indent from last line with text in Recepient?", this, &Table::lastLineInRecepient);
    pm->addAction("&What column for find negative values?", this, &Table::colorColumnRecepientFunc);


    paramMenu->setMenu(pm);

    
    savedConfig = new QPushButton("Save config", this);
    saveMenu = new QMenu(savedConfig);

    saveMenu->addAction("&Save current parameter as default", this, &Table::writeCurrent);
    saveMenu->addAction("&Save current parameter in other file", this, &Table::writeCurrentinOtherFile);

    savedConfig->setMenu(saveMenu);
    

    statusBar = new QStatusBar();
    statusBar->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);

    cb = new QCheckBox("Refresh Recepient table after VPR", this);
    connect(cb, &QCheckBox::stateChanged, this, &Table::checkStateForRefresh);

    dayNightCheck = new QCheckBox("Use Day/Night parameters", this);
    connect(dayNightCheck, &QCheckBox::stateChanged, this, &Table::checkDayNight);

    colorCheck = new QCheckBox("Find negative values", this);
    connect(colorCheck, &QCheckBox::stateChanged, this, &Table::checkColorRecepient);

    refresh = new QPushButton("Refresh", this);
    connect(refresh, &QPushButton::clicked, this, &Table::refreshAllButtons);



    VboxButtons->setSpacing(10); // расстояние между виджетами внутри вертикального бокса
    VboxButtons->addStretch(1); // равноудаляет от краёв или типо того
    VboxButtons->addWidget(cb);
    VboxButtons->addWidget(dayNightCheck);
    VboxButtons->addWidget(colorCheck);
    VboxButtons->addWidget(VPR);
    VboxButtons->addWidget(buttConvertToXML);
    VboxButtons->addWidget(donor);
    VboxButtons->addWidget(recepient);
    VboxButtons->addWidget(loadConfig);
    VboxButtons->addWidget(savedConfig);
    VboxButtons->addWidget(paramMenu);
    VboxButtons->addWidget(refresh);

    VboxButtons->addWidget(statusBar);
    VboxButtons->addStretch(1);

    Hbox->addLayout(Vbox, Qt::AlignRight);
    Hbox->addSpacing(10);
    Hbox->addLayout(VboxButtons, Qt::AlignLeft);

    readDefaultFileConfig();
}



void Table::myVPR()
{
    if (!Table::readyDonor || !Table::readyRecepient)
    {
        statusBar->showMessage("Add Donor first, recepient second!", 2000);

        return;
    }

    excelDonor = new QAxObject("Excel.Application", 0);// использование самого Excel. При использованиии ActiveX надо полагать что на всех целевыфх машинах будет установлен Excel. В общем указываем с каким приложением будем работать (к примеру могло быть "Outlook.Application")
    workbooksDonor = excelDonor->querySubObject("Workbooks"); // выбираем книгу
    workbookDonor = workbooksDonor->querySubObject("Open(const QString&)", addFileDonor); // выбираем файл с каким работать
    sheetsDonor = workbookDonor->querySubObject("Worksheets"); // обращаемся к листу
    sheetDonor = sheetsDonor->querySubObject("Item(int)", listDonor); // выбираем номер листа

    excelRecepient = new QAxObject("Excel.Application", 0);
    workbooksRecepient = excelRecepient->querySubObject("Workbooks");
    workbookRecepient = workbooksRecepient->querySubObject("Open(const QString&)", addFileRecepient);
    sheetsRecepient = workbookRecepient->querySubObject("Worksheets");
    sheetRecepient = sheetsRecepient->querySubObject("Item(int)", listRecepient);

    QElapsedTimer timer;

    int countTimer = 0;

    timer.start();

    QAxObject* copy = nullptr;
    QAxObject* compareDonor = nullptr;
    QAxObject* dayDonor = nullptr;
    QAxObject* compareRecepient = nullptr;
    QAxObject* paste = nullptr;
    QAxObject* dayRecepient = nullptr;
    QAxObject* negativeValue = nullptr;

    if (dayNightParametres)
    {
        // QList<vprStruct> tabelDonorFindAndDay;

        QMultiHash<QPair<QString, QString>, QVariant> tabelDonorFindAndDay; // профита нет

        for (int counter = memberRowFromFindDonor; counter <= (countRowsDonor - lastLineDonor); counter++)
        {
            compareDonor = sheetDonor->querySubObject("Cells(auto,auto)", counter, memberWhatFind);
            dayDonor = sheetDonor->querySubObject("Cells(auto,auto)", counter, memberwhereDayNightDonor);
            copy = sheetDonor->querySubObject("Cells(auto,auto)", counter, memberWhatToInsert);
            
           // QVariant val1 = compareDonor->property("Value").toString();
           // QVariant val2 = dayDonor->property("Value").toString();
           // QVariant val3 = copy->property("Value").toString();
           // vprStruct some = { val1, val2, val3 };
           // tabelDonorFindAndDay.append(some);
            tabelDonorFindAndDay.insert(QPair<QString, QString>{compareDonor->property("Value").toString(), dayDonor->property("Value").toString()}, copy->property("Value").toString());
            delete compareDonor;
            delete copy;
            delete dayDonor;
            
        }

        countTimer = timer.elapsed();
        out << "Creating an array finished in = " << (double)countTimer / 1000 << " sec" << Qt::endl;

        workbookDonor->dynamicCall("Close()"); 
        excelDonor->dynamicCall("Quit()");
        delete workbookDonor;
        delete excelDonor;

        // QListIterator<vprStruct> it(tabelDonorFindAndDay);

        QMultiHashIterator<QPair<QString, QString>, QVariant> it(tabelDonorFindAndDay);

        int countDoingIterationForTime = 0;

        for (int counter = memberRowFromFindRecepient; counter <= (countRowsRecepient - lastLineRecepient); counter++)
        {
            compareRecepient = sheetRecepient->querySubObject("Cells(&int,&int)", counter, memberWhereFind);
            paste = sheetRecepient->querySubObject("Cells(&int,&int)", counter, memberWhereToInsert);
            dayRecepient = sheetRecepient->querySubObject("Cells(&int,&int)", counter, memberwhereDayNightRecepient);
            negativeValue = sheetRecepient->querySubObject("Cells(&int,&int)", counter, colorColumnRecepint);

            while (it.hasNext())
            {
               it.next();

               //vprStruct temporary = it.next();

               //if ((temporary.whatFindStruct == compareRecepient->property("Value").toString()) && (temporary.dayNightStruct == dayRecepient->property("Value").toString())) 

                if ((it.key().first == compareRecepient->property("Value").toString()) && (it.key().second == dayRecepient->property("Value").toString())) // надо сравнивать QVariant с переводом в QString иначе не сравнивает.
                {
                    ++countDoingIterationForTime;

                    paste->dynamicCall("SetValue(String)", it.value().toDouble());

                   // tabelDonorFindAndDay.remove(it.key(), it.value()); // удаление записей из хэша (непомогло ускорить процесс)

                    qDebug() << "DONE WITH PARAM" << counter; // tabelDonorFindAndDay.count(); - для подсчёта остатков после удаления из хэша записей

                    delete compareRecepient;
                    delete paste;
                    delete dayRecepient;

                    if (colorChecked)
                    {
                        if (negativeValue->property("Value").toDouble() < 0)
                        {
                            // получаем указатель на её фон
                            QAxObject* interior = negativeValue->querySubObject("Interior");
                            // устанавливаем цвет
                            interior->setProperty("Color", QColor("red"));
                            // освобождение памяти
                            delete interior;
                        }
                    }

                    delete negativeValue;

                    break;
                }

                if (colorChecked)
                {
                    if (negativeValue->property("Value").toDouble() < 0)
                    {
                        // получаем указатель на её фон
                        QAxObject* interior = negativeValue->querySubObject("Interior");
                        // устанавливаем цвет
                        interior->setProperty("Color", QColor("red"));
                        // освобождение памяти
                        delete interior;
                    }
                }
            }
            it.toFront();

            delete sheetRecepient;
            delete sheetsRecepient;
            sheetsRecepient = workbookRecepient->querySubObject("Worksheets");
            sheetRecepient = sheetsRecepient->querySubObject("Item(int)", listRecepient);

        }
        
        countTimer = timer.elapsed();
        out << "VPR finished in = " << (double)countTimer / 1000 << " sec" << Qt::endl;
    }

    if (!dayNightParametres)
    {
        QMultiHash< QString, QString> tabelDonorFindAndDay; // QMultiMap

        for (int counter = memberRowFromFindDonor; counter < (countRowsDonor - lastLineDonor); counter++)
        {
            compareDonor = sheetDonor->querySubObject("Cells(auto,auto)", counter, memberWhatFind);
            copy = sheetDonor->querySubObject("Cells(auto,auto)", counter, memberWhatToInsert);
            QString val1 = compareDonor->property("Value").toString();
            QString val2 = copy->property("Value").toString();
            tabelDonorFindAndDay.insert(val1, val2);

        }

        delete compareDonor;
        delete copy;

        workbookDonor->dynamicCall("Close()");
        excelDonor->dynamicCall("Quit()");
        delete workbookDonor;
        delete excelDonor;

        QMultiHashIterator<QString, QString> it(tabelDonorFindAndDay);

        for (int counter = memberRowFromFindRecepient; counter < (countRowsRecepient - lastLineRecepient); counter++)
        {
            compareRecepient = sheetRecepient->querySubObject("Cells(&int,&int)", counter, memberWhereFind);
            paste = sheetRecepient->querySubObject("Cells(&int,&int)", counter, memberWhereToInsert);
            negativeValue = sheetRecepient->querySubObject("Cells(&int,&int)", counter, colorColumnRecepint);

            while (it.hasNext())
            {
                it.next();

                if (it.key() == compareRecepient->property("Value").toString())
                {
                    paste->dynamicCall("SetValue(double)", it.value());

                    delete compareRecepient;
                    delete paste;

                    if (colorChecked)
                    {
                        if (negativeValue->property("Value").toDouble() < 0)
                        {
                            // получаем указатель на её фон
                            QAxObject* interior = negativeValue->querySubObject("Interior");
                            // устанавливаем цвет
                            interior->setProperty("Color", QColor("red"));
                            // освобождение памяти
                            delete interior;
                        }
                    }

                    delete negativeValue;

                    qDebug() << "DONE NO PARAM" << counter;
                    break;
                }

                if (colorChecked)
                {
                    if (negativeValue->property("Value").toDouble() < 0)
                    {
                        // получаем указатель на её фон
                        QAxObject* interior = negativeValue->querySubObject("Interior");
                        // устанавливаем цвет
                        interior->setProperty("Color", QColor("red"));
                        // освобождение памяти
                        delete interior;
                    }
                }

            }
            it.toFront();

            delete sheetRecepient;
            delete sheetsRecepient;
            sheetsRecepient = workbookRecepient->querySubObject("Worksheets");
            sheetRecepient = sheetsRecepient->querySubObject("Item(int)", listRecepient);
        }

        countTimer = timer.elapsed();

        out << "VPR finished in = " << (double)countTimer / 1000 << " sec" << Qt::endl;
    }

    if (!refreshChecked)
    {
        workbookRecepient->dynamicCall("Close()");
        excelRecepient->dynamicCall("Quit()");
        delete workbookRecepient;
        delete excelRecepient;
        return;
    }

    timer.restart();

    delete Table::table2;

    usedRangeColRecepient = sheetRecepient->querySubObject("UsedRange"); // так можем получить количество столбцов в документе
    columnsRecepient = usedRangeColRecepient->querySubObject("Columns");
    countColsRecepient = columnsRecepient->property("Count").toInt();

    table2 = new QTableWidget(20, countColsRecepient, this);
    Vbox->addWidget(table2);

    QAxObject* cell = nullptr;
    QTableWidgetItem* item = nullptr;

    for (int row = 0; row < 20; ++row) {
        for (int column = 0; column < countColsRecepient; ++column) {

            cell = sheetRecepient->querySubObject("Cells(int,int)", row + 1, column + 1); // так указываем с какой ячейкой работать
            item = new QTableWidgetItem(cell->property("Value").toString());
            table2->setItem(row, column, item);
        }
    }

    delete cell;
    delete item;
    cell = nullptr;
    item = nullptr;

    countTimer = timer.elapsed();

    out << "Refresh recepient table in = " << (double)countTimer / 1000 << " sec" << Qt::endl;

    workbookRecepient->dynamicCall("Close()");
    excelRecepient->dynamicCall("Quit()");
    delete workbookRecepient;
    delete excelRecepient;
    return;
}



void Table::addDonor() {

   if (Table::readyDonor && Table::readyRecepient)
    {
        statusBar->showMessage("Maybe enough!", 2000);

        return;
    }

   if (Table::readyDonor)
   {
       statusBar->showMessage("Now addFileDonor recepient!", 2000);

       return;
   }
    
    addFileDonor = QFileDialog::getOpenFileName(0, "Open donor file", "", "*.xls *.xlsx");

    if (Table::addFileDonor == "")
    {
        return;
    }

    QElapsedTimer timer;

    int countTimer = 0;

    timer.start();

    excelDonor = new QAxObject("Excel.Application", 0); 
    workbooksDonor = excelDonor->querySubObject("Workbooks"); 
    workbookDonor = workbooksDonor->querySubObject("Open(const QString&)", addFileDonor); // 
    sheetsDonor = workbookDonor->querySubObject("Worksheets");
   
    listDonor = sheetsDonor->property("Count").toInt(); // так можем получить количество листов в документе
    
    if (listDonor > 1)
    {
        do 
        {
            listDonor = QInputDialog::getInt(this, "Number of list", "What list do you need?");
            if (!listDonor)
            {
                return;
            }
        } 
        while (listDonor <= 0 || (listDonor > (sheetsDonor->property("Count").toInt())));
        
    }

    sheetDonor = sheetsDonor->querySubObject("Item(int)", listDonor);// Тут определяем лист с которым будем работаь

    readyDonor = true;

    usedRangeDonor = sheetDonor->querySubObject("UsedRange"); // так можем получить количество строк в документе
    rowsDonor = usedRangeDonor->querySubObject("Rows");
    countRowsDonor = rowsDonor->property("Count").toInt();

    usedRangeColDonor = sheetDonor->querySubObject("UsedRange"); // так можем получить количество столбцов в документе
    columnsDonor = usedRangeColDonor->querySubObject("Columns");
    countColsDonor = columnsDonor->property("Count").toInt();

    table = new QTableWidget(20, countColsDonor, this); // создаём тамблицу по размеру той которую открываем в excelDonor

    Vbox->addWidget(table);

    QAxObject* cell = nullptr;
    QTableWidgetItem* item = nullptr;

    for (int row = 0; row < 20; ++row) {
        for (int column = 0; column < countColsDonor; ++column) {

            cell = sheetDonor->querySubObject("Cells(int,int)", row + 1, column + 1); // так указываем с какой ячейкой работать
            item = new QTableWidgetItem(cell->property("Value").toString());
            table->setItem(row, column, item);   
        }  
    }

    delete cell;
    delete item;

    countTimer = timer.elapsed();

    out << "Add Donor table and file = " << (double)countTimer/1000  <<" sec" << Qt::endl;

    if (countRowsDonor < lastLineDonor)
    {
        lastLineDonor = 0;

        qDebug() << "Indent from last line with text in Donor is more than the number of lines of text in Donor. Used default cofiguration.";
    }

    workbookDonor->dynamicCall("Close()"); // обязательно используем в работе с Excel иначе документы будет фbоном открыт в системе
    excelDonor->dynamicCall("Quit()");

    delete workbookDonor;
    delete excelDonor;
};



void Table::addRecepient() {

    if (!Table::readyDonor)
    {
        statusBar->showMessage("Add Donor first!", 2000);

        return;
    }

    if (Table::readyDonor && Table::readyRecepient)
    {
        statusBar->showMessage("Maybe enough!", 2000);

        return;
    }

    addFileRecepient = QFileDialog::getOpenFileName(0, "Open donor file", "", "*.xls *.xlsx");


    if (Table::addFileRecepient == "")
    {
        return;
    }

    readyRecepient = true;

    QElapsedTimer timer;

    int countTimer = 0;

    timer.start();

    excelRecepient = new QAxObject("Excel.Application", 0); // использование самого Excel. При использованиии ActiveX надо полагать что на всех целевыфх машинах будет установлен Excel. В общем указываем с каким приложением будем работать (к примеру могло быть "Outlook.Application")
    workbooksRecepient = excelRecepient->querySubObject("Workbooks"); // Витдимо это орпеделённая API для работы с COM объектом. В Нашем случае с Excel
    workbookRecepient = workbooksRecepient->querySubObject("Open(const QString&)", addFileRecepient); // Для взаимодействия со вторым файлом обязательно переопредлелять
    sheetsRecepient = workbookRecepient->querySubObject("Worksheets");// Для взаимодействия со вторым файлом обязательно переопредлелять
   
    listRecepient = sheetsRecepient->property("Count").toInt(); // так можем получить количество листов в документе

    if (listRecepient > 1)
    {
        do
        {
            listRecepient = QInputDialog::getInt(this, "Number of list", "What list do you need?");

            if (!listRecepient)
            {
                return;
            };

        } while (listRecepient <= 0 || (listRecepient > (sheetsRecepient->property("Count").toInt())));

    }
    
    sheetRecepient = sheetsRecepient->querySubObject("Item(int)", listRecepient);// Для взаимодействия со вторым файлом обязательно переопредлелять

    usedRangeRecepient = sheetRecepient->querySubObject("UsedRange"); // так можем получить количество строк в документе
    rowsRecepient = usedRangeRecepient->querySubObject("Rows");
    countRowsRecepient = rowsRecepient->property("Count").toInt();

    usedRangeColRecepient = sheetRecepient->querySubObject("UsedRange"); // так можем получить количество столбцов в документе
    columnsRecepient = usedRangeColRecepient->querySubObject("Columns");
    countColsRecepient = columnsRecepient->property("Count").toInt();

    table2 = new QTableWidget(20, countColsRecepient, this);
    Vbox->addWidget(table2);

    QAxObject* cell = nullptr;
    QTableWidgetItem* item = nullptr;

    // Наполняем таблицу 2 значениями из файла 2
    for (int row = 0; row < 20; ++row) {
        for (int column = 0; column < countColsRecepient; ++column) {

            cell = sheetRecepient->querySubObject("Cells(int,int)", row + 1, column + 1); // так указываем с какой ячейкой работать
            item = new QTableWidgetItem(cell->property("Value").toString());
            table2->setItem(row, column, item);
           // delete item;
        }
    }

    delete cell;
    delete item;

    countTimer = timer.elapsed();

    out << "Add Recepient table and file = " << (double)countTimer / 1000 << " sec" << Qt::endl;

    if (countRowsRecepient < lastLineRecepient)
    {
        lastLineRecepient = 0;

        qDebug() << "Indent from last line with text in Recepient is more than the number of lines of text in Donor. Used default cofiguration.";
    }

    workbookRecepient->dynamicCall("Close()");
    excelRecepient->dynamicCall("Quit()");

    delete workbookRecepient;
    delete excelRecepient;
};



void Table::whatFind()
{
    // bool ok необязательный параметр для inputDialog.getInt(). Откликается на нажатие Ок и Cancel в окне ввода данных. Соответственно становится true или false в зависимости от нажатой кнопки. 
    // Обязательно надо в начале задать какое то из двух значений чтобы состояния переменно коректно изменялись при нажатии кнопок. Учавствует в качестве указателя в параметрах. Передаём по адресу.
    bool ok = true; 
    QInputDialog inputDialog;
    QString now = "Specify Search Values. Now ";
    now.append(QString::number(memberWhatFind));
    int whatFind = inputDialog.getInt(this, "What find?", now, memberWhatFind, 0, 30, 1, &ok); // принадлежность/приписка над строкой ввода/имя окна/значение сразу введённое в окне/мin/max/шаг изменения значения от нажатия стрелок/bool статус нажатия конкретной кнопки (очень удобно)
    memberWhatFind = whatFind;
}

void Table::RowDoctor()
{
    bool ok = true;
    QInputDialog inputDialog;
    QString now = "Specify Search Values. Now ";
    now.append(QString::number(memberRowFromFindDonor));
    int whatFind = inputDialog.getInt(this, "What find?", now, memberRowFromFindDonor, 0, 30, 1, &ok);
    memberRowFromFindDonor = whatFind;
}

void Table::whereFind()
{
    bool ok = true;
    QInputDialog inputDialog;
    QString now = "Specify Search Values. Now ";
    now.append(QString::number(memberWhereFind));
    int whatFind = inputDialog.getInt(this, "Where find?", now, memberWhereFind, 0, 30, 1, &ok);
    memberWhereFind = whatFind;
}

void Table::RowRecepient()
{
    bool ok = true;
    QInputDialog inputDialog;
    QString now = "Specify Search Values. Now ";
    now.append(QString::number(memberRowFromFindRecepient));
    int whatFind = inputDialog.getInt(this, "What find?", now, memberRowFromFindRecepient, 0, 30, 1, &ok);
    memberRowFromFindRecepient = whatFind;
}

void Table::whereDayNightDonor()
{
    bool ok = true;
    QInputDialog inputDialog;
    QString now = "Specify where tariffing. Now ";
    now.append(QString::number(memberwhereDayNightDonor));
    int whatFind = inputDialog.getInt(this, "Where Day/Night?", now, memberwhereDayNightDonor, 0, 30, 1, &ok);
    memberwhereDayNightDonor = whatFind;
}

void Table::whereDayNightRecepient()
{
    bool ok = true;
    QInputDialog inputDialog;
    QString now = "Specify where tariffing. Now ";
    now.append(QString::number(memberwhereDayNightRecepient));
    int whatFind = inputDialog.getInt(this, "Where Day/Night?", now, memberwhereDayNightRecepient, 0, 30, 1, &ok);
    memberwhereDayNightRecepient = whatFind;
}

void Table::whatToInsert()
{
    bool ok = true;
    QInputDialog inputDialog;
    QString now = "Specify what to insert. Now ";
    now.append(QString::number(memberWhatToInsert));
    int whatFind = inputDialog.getInt(this, "Where to insert?", now, memberWhatToInsert, 0, 30, 1, &ok);
    memberWhatToInsert = whatFind;
}

void Table::whereToInsert()
{
    bool ok = true;
    QInputDialog inputDialog;
    QString now = "Specify where to insert. Now ";
    now.append(QString::number(memberWhereToInsert));
    int whatFind = inputDialog.getInt(this, "Where to insert?", now, memberWhereToInsert, 0, 30, 1, &ok);
    memberWhereToInsert = whatFind;
}

void Table::lastLineInDonor()
{
    bool ok = true;
    int border = 100;
    if (readyDonor) border = countRowsDonor;
    QInputDialog inputDialog;
    QString now = "Specify indent from last line with text in Donor. Now ";
    now.append(QString::number(lastLineDonor));
    int whatFind = inputDialog.getInt(this, "Indent from last line with text?", now, lastLineDonor, 0, border, 1, &ok);
    lastLineDonor = whatFind;
}

void Table::lastLineInRecepient()
{
    bool ok = true;
    int border = 100;
    if (readyRecepient) border = countRowsRecepient;
    QInputDialog inputDialog;
    QString now = "Specify indent from last line with text in Recepient. Now ";
    now.append(QString::number(lastLineRecepient));
    int whatFind = inputDialog.getInt(this, "Indent from last line with text?", now, lastLineRecepient, 0, border, 1, &ok);
    lastLineRecepient = whatFind;
}

void Table::colorColumnRecepientFunc()
{
    bool ok = true;
    QInputDialog inputDialog;
    QString now = "Specify where find Negative. Now ";
    now.append(QString::number(colorColumnRecepint));
    int whatFind = inputDialog.getInt(this, "Where find Negative?", now, colorColumnRecepint, 0, 30, 1, &ok);
    colorColumnRecepint = whatFind;
}

void Table::checkStateForRefresh(int state) {

    if (state == Qt::Checked) {
        refreshChecked = true;
    }
    else {
        refreshChecked = false;
    }
}

void Table::checkDayNight(int myState) {

    if (myState == Qt::Checked) {
        dayNightParametres = true;
    }
    else {
        dayNightParametres = false;
    }
}

void Table::checkColorRecepient(int myState) {

    if (myState == Qt::Checked) {
        colorChecked = true;
    }
    else {
        colorChecked = false;
    }
}

void Table::readFileConfig()
{
    QString saved = QFileDialog::getOpenFileName(0, "Load parameters from other file", "", "*.txt");

    if (saved == "")
    {
        return;
    }
    QFile configFile(saved);

    if (!configFile.open(QIODevice::ReadOnly))
    {
        out << "Dont find config file. Used default cofiguration." << Qt::endl;
        return;
    }

    QTextStream in(&configFile);

    int countParam = 0;

    // Считываем файл строка за строкой
    while (!in.atEnd())
    { // метод atEnd() возвращает true, если в потоке больше нет данных для чтения
        QString line = in.readLine(); // метод readLine() считывает одну строку из потока
        ++countParam;
        QString temporary;

        for (auto& val : line)
        {
            if (val.isDigit())
            {
                temporary += val;
            }
        }

        switch (countParam)
        {

        case(1):
        {
            qDebug() << "Where find in Donor before load config = " << memberWhatFind;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            memberWhatFind = temporary.toInt();
            qDebug() << "Where find in Donor after load config = " << memberWhatFind;
            break;
        }
        case(2):
        {
            qDebug() << "Start Row find in Donor before load config = " << memberRowFromFindDonor;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            memberRowFromFindDonor = temporary.toInt();
            qDebug() << "Start Row find in Donor after load config = " << memberRowFromFindDonor;
            break;
        }
        case(3):
        {
            qDebug() << "Where find in Recepient before load config = " << memberWhereFind;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            memberWhereFind = temporary.toInt();
            qDebug() << "Where find in Recepient after load config = " << memberWhereFind;
            break;
        }
        case(4):
        {
            qDebug() << "Start Row find in Recepient before load config = " << memberRowFromFindRecepient;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            memberRowFromFindRecepient = temporary.toInt();
            qDebug() << "Start Row find in Recepient after load config = " << memberRowFromFindRecepient;
            break;
        }
        case(5):
        {
            qDebug() << "Where Day/Night of Donor before load config = " << memberwhereDayNightDonor;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            memberwhereDayNightDonor = temporary.toInt();
            qDebug() << "Where Day/Night of Donor after load config = " << memberwhereDayNightDonor;
            break;
        }
        case(6):
        {
            qDebug() << "Where Day/Night of Recepient before load config = " << memberwhereDayNightRecepient;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            memberwhereDayNightRecepient = temporary.toInt();
            qDebug() << "Where Day/Night of Recepient after load config = " << memberwhereDayNightRecepient;
            break;
        }
        case(7):
        {
            qDebug() << "What to insert in Donor before load config = " << memberWhatToInsert;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            memberWhatToInsert = temporary.toInt();
            qDebug() << "What to insert in Donor after load config = " << memberWhatToInsert;
            break;
        }
        case(8):
        {
            qDebug() << "Where to insert in Recepient before load config = " << memberWhereToInsert;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            memberWhereToInsert = temporary.toInt();
            qDebug() << "Where to insert in Recepient after load config = " << memberWhereToInsert;
            break;
        }
        case(9):
        {
            qDebug() << "Refresh function before load config = " << refreshChecked;
            if ((temporary.toInt() < 0) || (temporary.toInt() > 1))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            refreshChecked = temporary.toInt();
            qDebug() << "Refresh function after load config = " << refreshChecked;
            break;
        }
        case(10):
        {
            qDebug() << "Day/Night function before load config = " << dayNightParametres;
            if ((temporary.toInt() < 0) || (temporary.toInt() > 1))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            dayNightParametres = temporary.toInt();
            qDebug() << "Day/Night function after load config = " << dayNightParametres;
            break;
        }
        case(11):
        {
            int border = 100;
            if (readyDonor) border = countRowsDonor;
            qDebug() << "Indent from last line with text in Donor before load config = " << lastLineDonor;
            if ((temporary.toInt() < 0) || (temporary.toInt() > border))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            lastLineDonor = temporary.toInt();
            qDebug() << "Indent from last line with text in Donor after load config = " << lastLineDonor;
            break;
        }
        case(12):
        {
            int border = 100;
            if (readyRecepient) border = countRowsRecepient;
            qDebug() << "Indent from last line with text in Recepient before load config = " << lastLineRecepient;
            if ((temporary.toInt() < 0) || (temporary.toInt() > border))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            lastLineRecepient = temporary.toInt();
            qDebug() << "Indent from last line with text in Recepient after load config = " << lastLineRecepient;
            break;
        }
        case(13):
        {
            qDebug() << "Where find negative values = " << colorColumnRecepint;
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            colorColumnRecepint = temporary.toInt();
            qDebug() << "Where find negative values = " << colorColumnRecepint;
            break;
        }
        case(14):
        {
            qDebug() << "Negative find function before load config = " << colorChecked;
            if ((temporary.toInt() < 0) || (temporary.toInt() > 1))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            colorChecked = temporary.toInt();
            qDebug() << "Negative find after load config = " << colorChecked;
            break;
        }
        }
    }

    configFile.close();
}



void Table::readDefaultFileConfig()
{
    QString filename = "config.txt";
    QFile file(filename);

	if (!file.open(QIODevice::ReadOnly))
	{
		out << "Dont fide config file. Used default cofiguration." << Qt::endl;
		return;
	}

	QTextStream in(&file);

	int countParam = 0;

	// Считываем файл строка за строкой
	while (!in.atEnd())
	{ // метод atEnd() возвращает true, если в потоке больше нет данных для чтения
		QString line = in.readLine(); // метод readLine() считывает одну строку из потока
		++countParam;
		QString temporary;

		for (auto& val : line)
		{ 
			if (val.isDigit())
			{
				temporary += val;
			}
		}

		switch (countParam)
		{

		case(1):
		{
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
			memberWhatFind = temporary.toInt();
			break;
		}
        case(2):
        {
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
            memberRowFromFindDonor = temporary.toInt();
            break;
        }
		case(3):
		{
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
			memberWhereFind = temporary.toInt();
			break;
		}
        case(4):
        {
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
            memberRowFromFindRecepient = temporary.toInt();
            break;
        }
		case(5):
		{
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
			memberwhereDayNightDonor = temporary.toInt();
			break;
		}
		case(6):
		{
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
			memberwhereDayNightRecepient = temporary.toInt();
			break;
		}
		case(7):
		{
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
			memberWhatToInsert = temporary.toInt();
			break;
		}
		case(8):
		{
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
			memberWhereToInsert = temporary.toInt();
			break;
		}
        case(9):
        {
            if ((temporary.toInt() < 0) || (temporary.toInt() > 1))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            refreshChecked = temporary.toInt();
            cb->setChecked(refreshChecked);
            break;
        }
        case(10):
        {
            if ((temporary.toInt() < 0) || (temporary.toInt() > 1))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            dayNightParametres = temporary.toInt();
            dayNightCheck->setChecked(dayNightParametres);
            break;
        }
        case(11):
        {
            if ((temporary.toInt() < 0) || (temporary.toInt() > 100))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            lastLineDonor = temporary.toInt();
            break;
        }
        case(12):
        {
            if ((temporary.toInt() < 0) || (temporary.toInt() > 100))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            lastLineRecepient = temporary.toInt();
            break;
        }
        case(13):
        {
            if ((temporary.toInt() < 1) || (temporary.toInt() > 30))
            {
                qDebug() << "Parameter in file going beyond borders! Default value will be used.";
                break;
            }
            colorColumnRecepint = temporary.toInt();
            break;
        }
        case(14):
        {
            if ((temporary.toInt() < 0) || (temporary.toInt() > 1))
            {
                qDebug() << "Parameter in file going beyond borders! Old value will be used.";
                break;
            }
            colorChecked = temporary.toInt();
            colorCheck->setChecked(colorChecked);
            break;
        }
		}
	}
    file.close();
}

void Table::writeCurrent()
{
    QString filename = "config.txt";
    QFile file(filename);
    
    // Открываем файл в режиме "Только для записи"
    if (file.open(QIODevice::WriteOnly)) {
        QTextStream out(&file); // поток записываемых данных направляем в файл

        // Для записи данных в файл используем оператор <<
        out << "memberWhatFind = " << memberWhatFind << Qt::endl;
        out << "memberRowFromFindDonor = " << memberRowFromFindDonor << Qt::endl;
        out << "memberWhereFind = " << memberWhereFind << Qt::endl;
        out << "memberRowFromFindRecepient = " << memberRowFromFindRecepient << Qt::endl;
        out << "memberwhereDayNightDonor = " << memberwhereDayNightDonor << Qt::endl;
        out << "memberwhereDayNightRecepient = " << memberwhereDayNightRecepient << Qt::endl;
        out << "memberWhatToInsert = " << memberWhatToInsert << Qt::endl;
        out << "memberWhereToInsert = " << memberWhereToInsert << Qt::endl;
        out << "refreshChecked = " << refreshChecked << Qt::endl;
        out << "dayNightParametres = " << dayNightParametres << Qt::endl;
        out << "lastLineDonor = " << lastLineDonor << Qt::endl;
        out << "lastLineRecepient = " << lastLineRecepient << Qt::endl;
        out << "colorColumnRecepint = " << colorColumnRecepint << Qt::endl;
        out << "colorChecked = " << colorChecked << Qt::endl;
    }
    else 
    {
        qWarning("Could not open file");
    }

    file.close();

    statusBar->showMessage("Default parameters was save.", 2000);
}

void Table::writeCurrentinOtherFile()
{
    QString savedFile = QFileDialog::getSaveFileName(0, "Save parameters in other file", "", "*.txt");

    if (savedFile == "") return;

    QFile file(savedFile);
    file.open(QIODevice::WriteOnly);
    QTextStream out(&file); // поток записываемых данных направляем в файл

    // Для записи данных в файл используем оператор <<
    out << "memberWhatFind = " << memberWhatFind << Qt::endl;
    out << "memberRowFromFindDonor = " << memberRowFromFindDonor << Qt::endl;
    out << "memberWhereFind = " << memberWhereFind << Qt::endl;
    out << "memberRowFromFindRecepient = " << memberRowFromFindRecepient << Qt::endl;
    out << "memberwhereDayNightDonor = " << memberwhereDayNightDonor << Qt::endl;
    out << "memberwhereDayNightRecepient = " << memberwhereDayNightRecepient << Qt::endl;
    out << "memberWhatToInsert = " << memberWhatToInsert << Qt::endl;
    out << "memberWhereToInsert = " << memberWhereToInsert << Qt::endl;
    out << "refreshChecked = " << refreshChecked << Qt::endl;
    out << "dayNightParametres = " << dayNightParametres << Qt::endl;
    out << "lastLineDonor = " << lastLineDonor << Qt::endl;
    out << "lastLineRecepient = " << lastLineRecepient << Qt::endl;
    out << "colorColumnRecepint = " << colorColumnRecepint << Qt::endl;
    out << "colorChecked = " << colorChecked << Qt::endl;

    file.close();

    statusBar->showMessage("New file with parameters was save.", 2000);
}

void Table::refreshAllButtons() // обновляет окно программы до начального состояния
{
    if (readyDonor)
    {
        delete Table::table;
        readyDonor = false;
    }

    if (readyRecepient)
    {
        delete Table::table2;
        readyRecepient = false;
    }
}

void Table::funcConvertToXML()
{
    if (!Table::readyDonor)
    {
        statusBar->showMessage("Add Donor first!", 2000);

        return;
    }

    QDate curDate = QDate::currentDate();
    QTime curTime = QTime::currentTime();

    QString fileName = "80020__";

    fileName += (curDate.toString("dd.MM.yyyy")) + "__" +(curTime.toString("hh:mm:ss"));

    for (int i = 0; i < fileName.size(); i++)
    {
        if (fileName[i].isPunct())
            fileName.remove(i, 1);
    }

    QString savedFile = QFileDialog::getSaveFileName(0, "Save XML", fileName, "*.xml"); // В последнем параметре также можно прописать tr("Xml files (*.xml)"). Это будет как приписка с указанием формата. Удобно.

    if (savedFile == "") return;

    QElapsedTimer timer;
    int countTimer = 0;
    timer.start();

    qDebug() << "Wait...";

    QFile file(savedFile);
    file.open(QIODevice::WriteOnly);

    excelDonor = new QAxObject("Excel.Application", 0);// использование самого Excel. При использованиии ActiveX надо полагать что на всех целевыфх машинах будет установлен Excel. В общем указываем с каким приложением будем работать (к примеру могло быть "Outlook.Application")
    workbooksDonor = excelDonor->querySubObject("Workbooks"); // выбираем книгу
    workbookDonor = workbooksDonor->querySubObject("Open(const QString&)", addFileDonor); // выбираем файл с каким работать
    sheetsDonor = workbookDonor->querySubObject("Worksheets"); // обращаемся к листу
    sheetDonor = sheetsDonor->querySubObject("Item(int)", listDonor); // выбираем номер листа
    
    QXmlStreamWriter xmlWriter(&file);
    xmlWriter.setAutoFormatting(true);
    xmlWriter.writeStartDocument();

    QAxObject* xmlAxObject = nullptr;

    checkXml();

    qDebug() << "xmlEsf = " << xmlEsf;
    qDebug() << "xmlZarya" << xmlZarya;

    qDebug() << "check 1";

    if (xmlEsf)
    {
        qDebug() << "check 2";

        xmlWriter.writeStartElement("message"); // отркывает начальный элемент "лестницы" xml
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 1);
        xmlWriter.writeAttribute("class", xmlAxObject->property("Value").toString()); // присваиваем атрибуты внутри открытого первого элемента
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 2);
        xmlWriter.writeAttribute("version", xmlAxObject->property("Value").toString());
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 3);
        xmlWriter.writeAttribute("number", xmlAxObject->property("Value").toString());

        xmlWriter.writeStartElement("datetime"); // отркывает второй элемент и т.д.

        xmlWriter.writeStartElement("timestamp");
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 4);
        xmlWriter.writeCharacters(xmlAxObject->property("Value").toString()); //вставка между открытием и закрытием элемента
        xmlWriter.writeEndElement(); // timestamp

        xmlWriter.writeStartElement("daylightsavingtime");
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 5);
        xmlWriter.writeCharacters(xmlAxObject->property("Value").toString());
        xmlWriter.writeEndElement(); // daylightsavingtime

        xmlWriter.writeStartElement("day");
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 6);
        xmlWriter.writeCharacters(xmlAxObject->property("Value").toString());
        xmlWriter.writeEndElement(); // day

        xmlWriter.writeEndElement(); // datetime

        xmlWriter.writeStartElement("sender");

        xmlWriter.writeStartElement("inn");
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 7);
        xmlWriter.writeCharacters(xmlAxObject->property("Value").toString());
        xmlWriter.writeEndElement(); // inn

        xmlWriter.writeStartElement("name");
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 8);
        xmlWriter.writeCharacters(xmlAxObject->property("Value").toString());
        xmlWriter.writeEndElement(); // name

        xmlWriter.writeEndElement(); // sender

        xmlWriter.writeStartElement("area");

        xmlWriter.writeStartElement("inn");
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 9);
        xmlWriter.writeCharacters(xmlAxObject->property("Value").toString());
        xmlWriter.writeEndElement(); // inn2

        xmlWriter.writeStartElement("name");
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 10);
        xmlWriter.writeCharacters(xmlAxObject->property("Value").toString());
        xmlWriter.writeEndElement(); // name3

        for (int counter = 2; counter <= countRowsDonor; counter++)
        {
            xmlWriter.writeStartElement("measuringpoint");

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 11);
            xmlWriter.writeAttribute("code", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 12);
            xmlWriter.writeAttribute("name", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 13);
            xmlWriter.writeAttribute("serial", xmlAxObject->property("Value").toString());

            for (int internalCounter = 0; internalCounter < 3; internalCounter++)
            {
                xmlWriter.writeStartElement("measuringchannel");

                xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter + internalCounter, 14);
                QString codeStr = xmlAxObject->property("Value").toString();
                if (codeStr == "1") codeStr = "0" + codeStr;
                xmlWriter.writeAttribute("code", codeStr);

                xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter + internalCounter, 15);
                xmlWriter.writeAttribute("desc", xmlAxObject->property("Value").toString());

                xmlWriter.writeStartElement("period");

                xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter + internalCounter, 16);
                QString periodStr = xmlAxObject->property("Value").toString();
                if (periodStr == "0") periodStr = "0000";
                xmlWriter.writeAttribute("start", periodStr);

                xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter + internalCounter, 17);
                xmlWriter.writeAttribute("end", xmlAxObject->property("Value").toString());

                xmlWriter.writeStartElement("timestamp");
                xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter + internalCounter, 18);
                xmlWriter.writeCharacters(xmlAxObject->property("Value").toString());

                xmlWriter.writeEndElement(); // value

                xmlWriter.writeStartElement("value");
                xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter + internalCounter, 19);
                xmlWriter.writeCharacters(xmlAxObject->property("Value").toString());

                xmlWriter.writeEndElement(); // timestamp

                xmlWriter.writeEndElement(); /// period

                xmlWriter.writeEndElement(); // measurechannel
            }

            counter = counter + 2; // делаем переход через две строки чтобы не дублировать строки с тарифами

            xmlWriter.writeEndElement(); // measurepoint
        }

        xmlWriter.writeEndElement(); // area

        xmlWriter.writeEndElement(); // message
    }

    if (xmlZarya)
    {
        qDebug() << "check 3";
        xmlWriter.writeStartElement("message"); // отркывает начальный элемент "лестницы" xml
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 1);
        xmlWriter.writeAttribute("class", xmlAxObject->property("Value").toString()); // присваиваем атрибуты внутри открытого первого элемента
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 2);
        xmlWriter.writeAttribute("version", xmlAxObject->property("Value").toString());
        xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", 2, 3);
        xmlWriter.writeAttribute("datetime", xmlAxObject->property("Value").toString());

        for (int counter = 2; counter <= countRowsDonor; counter++)
        {
            xmlWriter.writeStartElement("account");

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 4);
            xmlWriter.writeAttribute("street", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 5);
            xmlWriter.writeAttribute("house", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 6);
            xmlWriter.writeAttribute("flat", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 7);
            xmlWriter.writeAttribute("contract", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 8);
            xmlWriter.writeAttribute("numberId", xmlAxObject->property("Value").toString());

            xmlWriter.writeStartElement("counter");

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 9);
            xmlWriter.writeAttribute("number", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 10);
            xmlWriter.writeAttribute("typename", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 11);
            xmlWriter.writeAttribute("typeid", xmlAxObject->property("Value").toString());

            xmlWriter.writeStartElement("measure");

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 12);
            xmlWriter.writeAttribute("tariff", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 13);
            xmlWriter.writeAttribute("value", xmlAxObject->property("Value").toString());

            xmlAxObject = sheetDonor->querySubObject("Cells(&int,&int)", counter, 14);
            xmlWriter.writeAttribute("datetime", xmlAxObject->property("Value").toString());

            xmlWriter.writeEndElement(); // measure

            xmlWriter.writeEndElement(); // counter

            xmlWriter.writeEndElement(); // account
        }

        xmlWriter.writeEndElement(); // message
    }

    qDebug() << "check 4";
    
    delete xmlAxObject;

    xmlWriter.writeEndDocument();

    file.close();

    workbookDonor->dynamicCall("Close()");
    excelDonor->dynamicCall("Quit()");
    delete workbookDonor;
    delete excelDonor;  

    bool xmlEsf = false;
    bool xmlZarya = false;

    countTimer = timer.elapsed();
    out << "XLS to XML was convert for = " << (double)countTimer / 1000 << " sec" << Qt::endl;
}

void Table::checkXml()
{
    QAxObject* headOfFile = nullptr;
    QString compareStr;
    int count = 0;

    if (countColsDonor == 14)
    {
        qDebug() << "countColsDonor = " << countColsDonor;
        for (int column = 1; column <= countColsDonor; column++)
        {
            headOfFile = sheetDonor->querySubObject("Cells(&int,&int)", 1, column);
            compareStr = headOfFile->property("Value").toString();

            if (compareStr == "class" && column == 1) count++;
            if (compareStr == "version" && column == 2) count++;
            if (compareStr == "datetime" && column == 3) count++;
            if (compareStr == "street" && column == 4) count++;
            if (compareStr == "house" && column == 5) count++;
            if (compareStr == "flat" && column == 6) count++;
            if (compareStr == "contract" && column == 7) count++;
            if (compareStr == "numberId" && column == 8) count++;
            if (compareStr == "number" && column == 9) count++;
            if (compareStr == "typename" && column == 10) count++;
            if (compareStr == "typeid" && column == 11) count++;
            if (compareStr == "tariff" && column == 12) count++;
            if (compareStr == "value" && column == 13) count++;
            if (compareStr == "datetime2" && column == 14) count++;
        }

        if (count == 14)
        {
            qDebug() << "XLS convert in Zarya format XML";
            xmlZarya = true;
            delete headOfFile;
            return;
        }
        else
        {
            qDebug() << "count = " << count;
            qDebug() << "Incorrect format Zarya XLS file. Try again with correct file";
            delete headOfFile;
            return;
        }

    }

    if (countColsDonor == 19)
    {
        qDebug() << "countColsDonor = " << countColsDonor;
        for (int column = 1; column <= countColsDonor; column++)
        {
            headOfFile = sheetDonor->querySubObject("Cells(&int,&int)", 1, column);
            compareStr = headOfFile->property("Value").toString();

            if (compareStr == "class" && column == 1) count++;
            if (compareStr == "version" && column == 2) count++;
            if (compareStr == "number" && column == 3) count++;
            if (compareStr == "timestamp" && column == 4) count++;
            if (compareStr == "daylightsavingtime" && column == 5) count++;
            if (compareStr == "day" && column == 6) count++;
            if (compareStr == "inn" && column == 7) count++;
            if (compareStr == "name" && column == 8) count++;
            if (compareStr == "inn2" && column == 9) count++;
            if (compareStr == "name3" && column == 10) count++;
            if (compareStr == "code" && column == 11) count++;
            if (compareStr == "name4" && column == 12) count++;
            if (compareStr == "serial" && column == 13) count++;
            if (compareStr == "code5" && column == 14) count++;
            if (compareStr == "desc" && column == 15) count++;
            if (compareStr == "start" && column == 16) count++;
            if (compareStr == "end" && column == 17) count++;
            if (compareStr == "timestamp6" && column == 18) count++;
            if (compareStr == "value" && column == 19) count++;
        }

        if (count == 19)
        {
            qDebug() << "XLS convert in Esf format XML";
            xmlEsf = true;
            delete headOfFile;
            return;
        }
        else
        {
            qDebug() << "Incorrect format Esf XLS file. Try again with correct file";
            qDebug() << "count = " << count;
            delete headOfFile;
            return;
        }
    }

    qDebug() << "Incorrect format XLS file. Try again with correct file";
    return;
}