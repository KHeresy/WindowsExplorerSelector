#pragma once

#include <QtWidgets/QWidget>
#include "ui_QExplorerFinder.h"

QT_BEGIN_NAMESPACE
namespace Ui { class QExplorerFinderClass; };
QT_END_NAMESPACE

class QExplorerFinder : public QWidget
{
    Q_OBJECT

public:
    QExplorerFinder(QWidget *parent = nullptr);
    ~QExplorerFinder();

    void setTargetPath(const QString& path);

private slots:
    void on_btnSearch_clicked();
    void on_btnSelectAll_clicked();
    void on_lineEditSearch_returnPressed();
    
    void on_listResults_itemDoubleClicked(QListWidgetItem *item);
    void on_listResults_itemActivated(QListWidgetItem *item);
    void on_listResults_itemSelectionChanged();

    void on_btnListSelectAll_clicked();
    void on_btnListInvert_clicked();
    void on_btnListNone_clicked();
    
    void on_chkAlwaysOnTop_stateChanged(int state);

protected:
    void keyPressEvent(QKeyEvent *event) override;

private:
    Ui::QExplorerFinderClass *ui;
    QString m_targetPath;
    
    void selectFiles(const QStringList& files);
    void updateStatusLabel();
    void loadHistory();
    void saveHistory(const QString& text);
};

