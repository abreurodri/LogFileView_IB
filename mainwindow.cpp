#include <QFileDialog>
#include <QTextStream>
#include <QMessageBox>
#include <QDateTime>
#include <QDebug>
#include <QClipboard>
#include <QMimeData>

#include <QtXlsx//xlsxdocument.h>
#include <QtXlsx/xlsxformat.h>

#include "mainwindow.h"
#include "frmpedirip.h"
#include "ui_mainwindow.h"


// Para realizar las ordenaciones
class QTableWidgetItemDate : public QTableWidgetItem {

    public:
        QTableWidgetItemDate(QString const& str)
            : QTableWidgetItem(str)
        {}

    public:
        bool operator <(const QTableWidgetItem &other) const
        {
            QStringList lista      = text().split("/");
            QStringList listaOther = other.text().split("/");

            for (int ii = lista.count(); ii > 0; --ii)
            {
                if (lista.at(ii - 1).toInt() < listaOther.at(ii -1).toInt())
                    return true;
                else if (lista.at(ii -1).toInt() > listaOther.at(ii -1).toInt())
                    return false;
            }

            return false;
        }
};

class QTableWidgetItemTime : public QTableWidgetItem {

    public:
        QTableWidgetItemTime(QString const& str)
            : QTableWidgetItem(str)
        {}

    public:
        bool operator <(const QTableWidgetItem &other) const
        {
            QStringList lista      = text().split(QRegExp("[:.]"));
            QStringList listaOther = other.text().split(QRegExp("[:.]"));

            for (int ii = 0; ii < lista.count(); ++ii)
            {
                if (lista.at(ii).toInt() < listaOther.at(ii).toInt())
                    return true;
                else if (lista.at(ii).toInt() > listaOther.at(ii).toInt())
                    return false;
            }

            return false;
        }
};

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    // Pongo cabeceras
    //QString cabecera = "Id;Segundos;Descripción;Estado;a;b;c;Ref61850;Dia y Hora";
    QString cabecera = "Date;Time;Domain;Signal;Status";
    QStringList lista(cabecera.split(";"));
    ui->tableWidget->setHorizontalHeaderLabels(lista);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_actionAbrir_archivo_triggered()
{
    QString nombreFich = QFileDialog::getOpenFileName(this, "Open", "", "Events file (*.his);; All files (*.*)");
    if (nombreFich == "")
        return;

    // Limpiamos la tabla
    ui->tableWidget->setRowCount(0);

    // Lo incluimos en la tabla
    insert_histfile_on_table(nombreFich, ui->tableWidget->rowCount());
}

void MainWindow::on_actionAbrir_carpeta_triggered()
{
    QString nombreFolder = QFileDialog::getExistingDirectory( this, "Open Folder", "");
    QString nombreFich;

    //Nos quedamos con aquellos que tengan extension .his
    QStringList listaFich = QDir(nombreFolder).entryList().filter(QRegExp(".*\\.his$"));

    // Limpiamos la tabla
    ui->tableWidget->setRowCount(0);

    for (int i = 0; i < listaFich.size(); ++i)
    {
        // Cada Fichero que leemos
        nombreFich = nombreFolder + "/" + listaFich.at(i);

        // Lo incluimos en la tabla
        insert_histfile_on_table(nombreFich, ui->tableWidget->rowCount());
    }
}

void MainWindow::on_actionDownload_Event_Files_triggered()
{
    FrmPedirIp *w;
    w = new FrmPedirIp(this);
    int ret = w->exec();

    if (ret == QDialog::Accepted)
    {
        QString appPath = qApp->applicationDirPath();
        QString histPath = appPath + "/histfiles";
        QString nombreFich;

        //Nos quedamos con aquellos que tengan extension .his
        QStringList listaFich = QDir(histPath).entryList().filter(QRegExp(".*\\.his$"));

        // Limpiamos la tabla
        ui->tableWidget->setRowCount(0);

        for (int i = 0; i < listaFich.size(); ++i)
        {
            // Cada Fichero que leemos
            nombreFich = histPath + "/" + listaFich.at(i);

            // Lo incluimos en la tabla
            insert_histfile_on_table(nombreFich, ui->tableWidget->rowCount());
        }
    }

    delete(w);
}

void MainWindow::insert_histfile_on_table(QString const& filePath, int rowPos)
{
    QFile fichEventos(filePath);
    QString linea;

    if (fichEventos.open(QIODevice::Text | QIODevice::ReadOnly))
    {
        QTextStream streamIn(&fichEventos);
        QStringList campos;

        // Lo hacemos cuando queremos limpiar la tabla
        //ui->tableWidget->setRowCount(0);    // Utilizo setRowCount(0) porque el método clear() me quita las cabeceras
        int indexRow=rowPos;

        while (!streamIn.atEnd())
        {
            linea = streamIn.readLine();
            linea.replace(",",";");
            linea.replace(":", ";");
            campos = linea.split(";");

            /********************************************/
            /*  COMPROBAMOS QUE EL REGISTRO SEA VALIDO  */
            /********************************************/
            // Si la descripcion incorpora "Sin Descripcion"
            // Si el estado es vacio
            // Se considera invalido el registro y se ignora
            if (campos[3].contains("Sin Descripcion"))
                continue;

            if (campos[4].isEmpty())
                continue;


            /********************************************/
            /*   PREPARAMOS LA INFORMACION A INSERTAR   */
            /********************************************/
            // Fichero [IdElem;TimeStamp;QTimeStamp;Descripcion;Estado;Valor;Calidad;Referencia]
            QString sDate, sTime, sDomain;
            bool bTimeQualFailure;

            // Obtenemos la calidad de la marca de tiempo
            if (campos[2] == "1")
            {
                // Calidad invalida de marca de tiempo
                bTimeQualFailure = true;
            }
            else
            {
                bTimeQualFailure = false;
            }

            // Obtenemos Date y Time
            bool bConvertOk = false;
            unsigned long long s = campos[1].toULongLong(&bConvertOk);
            if (bConvertOk)
            {
                const QDateTime dt = QDateTime::fromMSecsSinceEpoch(s);
                sDate = dt.toString("  dd/MM/yyyy  ");
                sTime = dt.toString("  hh:mm:ss.zzz  ");
            }

            // Obtenemos el dominio
            sDomain = campos[7].mid(0, campos[7].indexOf("/")) + "    ";

            /********************************************/
            /*      INSERTAMOS LA FILA EN LA TABLA      */
            /********************************************/
            // Tabla [Date;Time;Domain;Signal;Status]

            // Crea fila (inserta y añade widgets de texto)
            ui->tableWidget->insertRow(indexRow);

            // Insertamos la fecha
            int indexColumn = 0;

            ui->tableWidget->setItem(indexRow, indexColumn, new QTableWidgetItemDate(sDate));
            if (bTimeQualFailure)
                ui->tableWidget->item(indexRow, indexColumn)->setForeground(QColor(255,0,0));

            // Insertamos la marca de tiempo
            indexColumn++;
            ui->tableWidget->setItem(indexRow, indexColumn, new QTableWidgetItemTime(sTime));
            if (bTimeQualFailure)
                ui->tableWidget->item(indexRow, indexColumn)->setForeground(QColor(255,0,0));

            // Insertamos el dominio
            indexColumn++;
            ui->tableWidget->setItem(indexRow, indexColumn, new QTableWidgetItem(sDomain));

            // Insertamos la descripcion
            indexColumn++;
            ui->tableWidget->setItem(indexRow, indexColumn, new QTableWidgetItem(campos[3] + "    "));

            // Insertamos el estado
            indexColumn++;
            ui->tableWidget->setItem(indexRow, indexColumn, new QTableWidgetItem(campos[4] + "    "));

            indexRow++;
        }

        // Autoajuste de columnas
        ui->tableWidget->resizeColumnsToContents();
        ui->tableWidget->resizeRowsToContents();
        fichEventos.close();
    }
    else
    {
        QMessageBox::critical(this, "Error", "No se pudo abrir el archivo");
        return;
    }

    // Ordenamos la tabla por columnas
    int numColumnas = ui->tableWidget->columnCount();
    for (int col = numColumnas; col > 0; --col)
    {
        ui->tableWidget->sortByColumn(col - 1);
    }

    // Eliminamos duplicados (Al estar ordenado un registro tiene que comprobar que es diferente del anterior)
    QTableWidgetItem *itemA, *itemB;

    for (int fila = 1; fila < ui->tableWidget->rowCount(); ++fila)
    {
        bool bElimina = true;

        for (int col = 0; col < numColumnas; ++col)
        {
            itemA = ui->tableWidget->item(fila - 1, col);
            itemB = ui->tableWidget->item(fila, col);

            if (itemA->text().compare(itemB->text()) != 0)
                bElimina = false;
        }

        if (bElimina)
        {
            // Fila repetida, la eliminamos
            ui->tableWidget->removeRow(fila);
        }
    }

}

void MainWindow::on_actionCopiar_triggered()
{
    QClipboard *clipboard = QApplication::clipboard();
    QByteArray ByteArray;
    QTableWidgetItem *item;
    int numFilas = ui->tableWidget->rowCount();
    int numColumnas = ui->tableWidget->columnCount();

    ByteArray.clear();

    // Pone cabecera
    for (int col = 0; col < numColumnas; ++col)
        ByteArray.append(ui->tableWidget->horizontalHeaderItem(col)->text() + ",");
    ByteArray.append("\n");

    // Recorre filas y columnas
    for (int fila = 0; fila < numFilas; ++fila)
    {
        for (int col = 0; col < numColumnas; ++col)
        {
            item = ui->tableWidget->item(fila, col);
            ByteArray.append(item->text() + ",");
        }
        // Sustituye la última coma por nueva línea
        ByteArray.replace(ByteArray.length(), 1, "\n");
    }

    QMimeData *mimeData = new QMimeData();
    mimeData->setData("text/plain", ByteArray);
    clipboard->setMimeData(mimeData);
}

void MainWindow::on_actionExportar_a_Excel_2010_triggered()
{
    // 1. Leer y preparar archivo donde guardar
    QString nombreFich = QFileDialog::getSaveFileName(this, "Guardar como", "", "Archivo de Excel 2010 (*.xlsx);; All files (*.*)");
    if (nombreFich == "")
    {
        QMessageBox::information(this, tr("Exportar a Excel"), tr("Operación cancelada"));
        return;
    }

    QXlsx::Document xlsx;

    // 2. Recorrer lista añadiendo a la hoja de excel
    QTableWidgetItem *item;
    int numFilas = ui->tableWidget->rowCount();
    int numColumnas = ui->tableWidget->columnCount();

    // Pone cabecera
    QXlsx::Format font_bold;
    font_bold.setFontBold(true);
    for (int col = 0; col < numColumnas; ++col)
    {
        xlsx.write(1, col + 1, ui->tableWidget->horizontalHeaderItem(col)->text(), font_bold); // En Xlsx, primera celda 1,1
        xlsx.setColumnWidth(col + 1, (double) ui->tableWidget->columnWidth(col) / 10);
    }

    // Recorre la tabla
    for (int fila = 0; fila < numFilas; ++fila)
    {
        for (int col = 0; col < numColumnas; ++col)
        {
            item = ui->tableWidget->item(fila, col);
            xlsx.write(fila + 2, col + 1, item->text()); // Primera fila en la que escribir es la 2 (la 1 son los títulos)
        }
    }

    // 3. Guardar archivo
    if (xlsx.saveAs(nombreFich))
        QMessageBox::information(this, tr("Exportar a Excel"), tr("Se guardo correctamente"));
    else
        QMessageBox::critical(this, tr("Exportar a Excel"), tr("Error al guardar el archivo"));

}
