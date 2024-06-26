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
    pm = new QMenu(paramMenu); // �������������� ���������� ������

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



    VboxButtons->setSpacing(10); // ���������� ����� ��������� ������ ������������� �����
    VboxButtons->addStretch(1); // ������������ �� ���� ��� ���� ����
    VboxButtons->addWidget(cb);
    VboxButtons->addWidget(dayNightCheck);
    VboxButtons->addWidget(colorCheck);
    VboxButtons->addWidget(VPR);
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

    excelDonor = new QAxObject("Excel.Application", 0);// ������������� ������ Excel. ��� �������������� ActiveX ���� �������� ��� �� ���� �������� ������� ����� ���������� Excel. � ����� ��������� � ����� ����������� ����� �������� (� ������� ����� ���� "Outlook.Application")
    workbooksDonor = excelDonor->querySubObject("Workbooks"); // �������� �����
    workbookDonor = workbooksDonor->querySubObject("Open(const QString&)", addFileDonor); // �������� ���� � ����� ��������
    sheetsDonor = workbookDonor->querySubObject("Worksheets"); // ���������� � �����
    sheetDonor = sheetsDonor->querySubObject("Item(int)", listDonor); // �������� ����� �����

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

        QMultiHash<QPair<QString, QString>, QVariant> tabelDonorFindAndDay; // ������� ���

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

                if ((it.key().first == compareRecepient->property("Value").toString()) && (it.key().second == dayRecepient->property("Value").toString())) // ���� ���������� QVariant � ��������� � QString ����� �� ����������.
                {
                    ++countDoingIterationForTime;

                    paste->dynamicCall("SetValue(String)", it.value().toDouble());

                   // tabelDonorFindAndDay.remove(it.key(), it.value()); // �������� ������� �� ���� (��������� �������� �������)

                    qDebug() << "DONE WITH PARAM" << counter; // tabelDonorFindAndDay.count(); - ��� �������� �������� ����� �������� �� ���� �������

                    delete compareRecepient;
                    delete paste;
                    delete dayRecepient;

                    if (colorChecked)
                    {
                        if (negativeValue->property("Value").toDouble() < 0)
                        {
                            // �������� ��������� �� � ���
                            QAxObject* interior = negativeValue->querySubObject("Interior");
                            // ������������� ����
                            interior->setProperty("Color", QColor("red"));
                            // ������������ ������
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
                        // �������� ��������� �� � ���
                        QAxObject* interior = negativeValue->querySubObject("Interior");
                        // ������������� ����
                        interior->setProperty("Color", QColor("red"));
                        // ������������ ������
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
                            // �������� ��������� �� � ���
                            QAxObject* interior = negativeValue->querySubObject("Interior");
                            // ������������� ����
                            interior->setProperty("Color", QColor("red"));
                            // ������������ ������
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
                        // �������� ��������� �� � ���
                        QAxObject* interior = negativeValue->querySubObject("Interior");
                        // ������������� ����
                        interior->setProperty("Color", QColor("red"));
                        // ������������ ������
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

    usedRangeColRecepient = sheetRecepient->querySubObject("UsedRange"); // ��� ����� �������� ���������� �������� � ���������
    columnsRecepient = usedRangeColRecepient->querySubObject("Columns");
    countColsRecepient = columnsRecepient->property("Count").toInt();

    table2 = new QTableWidget(20, countColsRecepient, this);
    Vbox->addWidget(table2);

    QAxObject* cell = nullptr;
    QTableWidgetItem* item = nullptr;

    for (int row = 0; row < 20; ++row) {
        for (int column = 0; column < countColsRecepient; ++column) {

            cell = sheetRecepient->querySubObject("Cells(int,int)", row + 1, column + 1); // ��� ��������� � ����� ������� ��������
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
   
    listDonor = sheetsDonor->property("Count").toInt(); // ��� ����� �������� ���������� ������ � ���������
    
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

    sheetDonor = sheetsDonor->querySubObject("Item(int)", listDonor);// ��� ���������� ���� � ������� ����� �������

    readyDonor = true;

    usedRangeDonor = sheetDonor->querySubObject("UsedRange"); // ��� ����� �������� ���������� ����� � ���������
    rowsDonor = usedRangeDonor->querySubObject("Rows");
    countRowsDonor = rowsDonor->property("Count").toInt();

    usedRangeColDonor = sheetDonor->querySubObject("UsedRange"); // ��� ����� �������� ���������� �������� � ���������
    columnsDonor = usedRangeColDonor->querySubObject("Columns");
    countColsDonor = columnsDonor->property("Count").toInt();

    table = new QTableWidget(20, countColsDonor, this); // ������ �������� �� ������� ��� ������� ��������� � excelDonor

    Vbox->addWidget(table);

    QAxObject* cell = nullptr;
    QTableWidgetItem* item = nullptr;

    for (int row = 0; row < 20; ++row) {
        for (int column = 0; column < countColsDonor; ++column) {

            cell = sheetDonor->querySubObject("Cells(int,int)", row + 1, column + 1); // ��� ��������� � ����� ������� ��������
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

    workbookDonor->dynamicCall("Close()"); // ����������� ���������� � ������ � Excel ����� ��������� ����� �b���� ������ � �������
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

    excelRecepient = new QAxObject("Excel.Application", 0); // ������������� ������ Excel. ��� �������������� ActiveX ���� �������� ��� �� ���� �������� ������� ����� ���������� Excel. � ����� ��������� � ����� ����������� ����� �������� (� ������� ����� ���� "Outlook.Application")
    workbooksRecepient = excelRecepient->querySubObject("Workbooks"); // ������� ��� ����������� API ��� ������ � COM ��������. � ����� ������ � Excel
    workbookRecepient = workbooksRecepient->querySubObject("Open(const QString&)", addFileRecepient); // ��� �������������� �� ������ ������ ����������� ���������������
    sheetsRecepient = workbookRecepient->querySubObject("Worksheets");// ��� �������������� �� ������ ������ ����������� ���������������
   
    listRecepient = sheetsRecepient->property("Count").toInt(); // ��� ����� �������� ���������� ������ � ���������

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
    
    sheetRecepient = sheetsRecepient->querySubObject("Item(int)", listRecepient);// ��� �������������� �� ������ ������ ����������� ���������������

    usedRangeRecepient = sheetRecepient->querySubObject("UsedRange"); // ��� ����� �������� ���������� ����� � ���������
    rowsRecepient = usedRangeRecepient->querySubObject("Rows");
    countRowsRecepient = rowsRecepient->property("Count").toInt();

    usedRangeColRecepient = sheetRecepient->querySubObject("UsedRange"); // ��� ����� �������� ���������� �������� � ���������
    columnsRecepient = usedRangeColRecepient->querySubObject("Columns");
    countColsRecepient = columnsRecepient->property("Count").toInt();

    table2 = new QTableWidget(20, countColsRecepient, this);
    Vbox->addWidget(table2);

    QAxObject* cell = nullptr;
    QTableWidgetItem* item = nullptr;

    // ��������� ������� 2 ���������� �� ����� 2
    for (int row = 0; row < 20; ++row) {
        for (int column = 0; column < countColsRecepient; ++column) {

            cell = sheetRecepient->querySubObject("Cells(int,int)", row + 1, column + 1); // ��� ��������� � ����� ������� ��������
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
    // bool ok �������������� �������� ��� inputDialog.getInt(). ����������� �� ������� �� � Cancel � ���� ����� ������. �������������� ���������� true ��� false � ����������� �� ������� ������. 
    // ����������� ���� � ������ ������ ����� �� �� ���� �������� ����� ��������� ��������� �������� ���������� ��� ������� ������. ���������� � �������� ��������� � ����������. ������� �� ������.
    bool ok = true; 
    QInputDialog inputDialog;
    QString now = "Specify Search Values. Now ";
    now.append(QString::number(memberWhatFind));
    int whatFind = inputDialog.getInt(this, "What find?", now, memberWhatFind, 0, 30, 1, &ok); // ��������������/�������� ��� ������� �����/��� ����/�������� ����� �������� � ����/�in/max/��� ��������� �������� �� ������� �������/bool ������ ������� ���������� ������ (����� ������)
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

    // ��������� ���� ������ �� �������
    while (!in.atEnd())
    { // ����� atEnd() ���������� true, ���� � ������ ������ ��� ������ ��� ������
        QString line = in.readLine(); // ����� readLine() ��������� ���� ������ �� ������
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

	// ��������� ���� ������ �� �������
	while (!in.atEnd())
	{ // ����� atEnd() ���������� true, ���� � ������ ������ ��� ������ ��� ������
		QString line = in.readLine(); // ����� readLine() ��������� ���� ������ �� ������
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
    
    // ��������� ���� � ������ "������ ��� ������"
    if (file.open(QIODevice::WriteOnly)) {
        QTextStream out(&file); // ����� ������������ ������ ���������� � ����

        // ��� ������ ������ � ���� ���������� �������� <<
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
    QTextStream out(&file); // ����� ������������ ������ ���������� � ����

    // ��� ������ ������ � ���� ���������� �������� <<
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

void Table::refreshAllButtons() // ��������� ���� ��������� �� ���������� ���������
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
