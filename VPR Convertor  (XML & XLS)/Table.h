#pragma once

#include <QWidget>
#include <QPushButton>
#include <QTableWidget>
#include <QAxObject>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QFileDialog>
#include <QMenu>
#include <QStatusBar>
#include <QMainWindow>
#include <algorithm>
#include <QCheckBox>

#include <QXmlStreamWriter>
#include <QXmlStreamAttribute>


class Table : public QWidget {

    Q_OBJECT

public:
    Table(QWidget* parent = 0);

private slots:
    void myVPR();
    void addDonor();
    void addRecepient();
    void whatFind();
    void RowDoctor();
    void whereFind();
    void RowRecepient();
    void whereDayNightDonor();
    void whereDayNightRecepient();
    void whatToInsert();
    void whereToInsert();
    void checkStateForRefresh(int state);
    void checkDayNight(int myState);
    void readFileConfig();
    void writeCurrent();
    void readDefaultFileConfig();
    void writeCurrentinOtherFile();
    void refreshAllButtons();
    void lastLineInDonor();
    void lastLineInRecepient();
    void colorColumnRecepientFunc();
    void checkColorRecepient(int myState);
    void funcConvertToXML();
    void checkXml();

private:

    QPushButton* VPR = nullptr;
    QPushButton* buttConvertToXML = nullptr;
    QPushButton* donor = nullptr;
    QPushButton* recepient = nullptr;
    QPushButton* paramMenu = nullptr;
    QPushButton* loadConfig = nullptr;
    QPushButton* savedConfig = nullptr;
    QPushButton* refresh = nullptr;

    QMenu* pm = nullptr;
    QMenu* saveMenu = nullptr;

    QTableWidget* table = nullptr;
    QTableWidget* table2 = nullptr;

    int countRowsDonor = 0;
    int countRowsRecepient = 0;
    int countColsDonor = 0;
    int countColsRecepient = 0;
    int memberWhatFind = 1;
    int memberRowFromFindDonor = 1;
    int memberWhereFind = 1;
    int memberRowFromFindRecepient = 1;
    int memberwhereDayNightDonor = 2;
    int memberwhereDayNightRecepient = 2;
    int memberWhatToInsert = 4;
    int memberWhereToInsert = 4;
    int listDonor = 1;
    int listRecepient = 1;
    int lastLineDonor = 0;
    int lastLineRecepient = 0;
    int colorColumnRecepint = 9;

    bool readyDonor = false;
    bool readyRecepient = false;
    bool refreshChecked = false;
    bool dayNightParametres = false;
    bool colorChecked = false;
    bool xmlEsf = false;
    bool xmlZarya = false;

    QString addFileDonor;
    QString addFileRecepient;
    QString addDefaultConfigFile;

    QAxObject* excelDonor = nullptr;
    QAxObject* workbooksDonor = nullptr;
    QAxObject* workbookDonor = nullptr;
    QAxObject* sheetsDonor = nullptr;
    QAxObject* sheetDonor = nullptr;
    QAxObject* usedRangeDonor = nullptr;
    QAxObject* rowsDonor = nullptr;
    QAxObject* usedRangeColDonor = nullptr;
    QAxObject* columnsDonor = nullptr;

    QAxObject* excelRecepient = nullptr;
    QAxObject* workbooksRecepient = nullptr;
    QAxObject* workbookRecepient = nullptr;
    QAxObject* sheetsRecepient = nullptr;
    QAxObject* sheetRecepient = nullptr;
    QAxObject* usedRangeRecepient = nullptr;
    QAxObject* rowsRecepient = nullptr;
    QAxObject* usedRangeColRecepient = nullptr;
    QAxObject* columnsRecepient = nullptr;
    QAxObject* colorColumnRecepient = nullptr;

    QVBoxLayout* Vbox = nullptr;

    QStatusBar* statusBar;

    QCheckBox* cb;
    QCheckBox* dayNightCheck;
    QCheckBox* colorCheck;

    
    /* // Использовался в ранеем варианте VPR. Изменения не отразились на скорости работы алгоритма
    struct vprStruct
    {
        QVariant whatFindStruct;
        QVariant dayNightStruct;
        QVariant valueStruct;
    };
    */
};