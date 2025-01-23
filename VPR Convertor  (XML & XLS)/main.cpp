#include <QApplication>
#include "table.h"

int main(int argc, char* argv[]) {

    QApplication app(argc, argv);

    Table window;

    window.resize(650, 450);
    window.setWindowIcon(QIcon("icon.png"));
    window.setWindowTitle("VPR Convertor by Solovev");
    window.show();

    return app.exec();
}




