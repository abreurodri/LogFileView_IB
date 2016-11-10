#ifndef FRMPEDIRIP_H
#define FRMPEDIRIP_H

#include <QDialog>

namespace Ui {
class FrmPedirIp;
}

class FrmPedirIp : public QDialog
{
    Q_OBJECT

public:
    explicit FrmPedirIp(QWidget *parent = 0);
    ~FrmPedirIp();

private slots:

    void on_btnDownload_clicked();

private:
    Ui::FrmPedirIp *ui;
};

#endif // FRMPEDIRIP_H
