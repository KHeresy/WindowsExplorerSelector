#include "QExplorerFinder.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    QExplorerFinder window;
    
    if (argc > 1) {
        window.setTargetPath(QString::fromLocal8Bit(argv[1]));
        window.setWindowTitle("Finding in: " + QString::fromLocal8Bit(argv[1]));
    }

    window.show();
    return app.exec();
}
