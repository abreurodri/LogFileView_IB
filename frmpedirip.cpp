#include <QProcess>
#include <QString>
#include <QMessageBox>
#include <QDir>

#include "frmpedirip.h"
#include "ui_frmpedirip.h"

FrmPedirIp::FrmPedirIp(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::FrmPedirIp)
{
    ui->setupUi(this);
}

FrmPedirIp::~FrmPedirIp()
{
    delete ui;
}

void FrmPedirIp::on_btnDownload_clicked()
{
    QString appPath = qApp->applicationDirPath();


    // Borramos el contenido de la actual carpeta de historico
    QDir oQDir(appPath + "/histfiles");
    oQDir.removeRecursively();

    QString programa = "\"" + appPath + "/pscp.exe\" -pw scusac -r scu@" + ui->lineDirIp->text() + ":histfiles \"" + appPath + "\"";

    int nResCommand = QProcess::execute(programa);

    if (nResCommand == 0)
        QDialog::accept();
    else
        QDialog::reject();
}
