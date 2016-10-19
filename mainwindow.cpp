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
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    // Pongo cabeceras
    QString cabecera = "Id;Segundos;Descripción;Estado;a;b;c;Ref61850;Dia y Hora";
    QStringList lista(cabecera.split(";"));
    ui->tableWidget->setHorizontalHeaderLabels(lista);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_actionAbrir_archivo_triggered()
{
    QString nombreFich = QFileDialog::getOpenFileName(this, "Abrir", "", "Histórico de eventos (*.his);; All files (*.*)");
    if (nombreFich == "")
        return;

    QFile fichEventos(nombreFich);
    QString linea;

    if (fichEventos.open(QIODevice::Text | QIODevice::ReadOnly))
    {
        QTextStream streamIn(&fichEventos);
        QStringList campos;
        ui->tableWidget->setRowCount(0);    // Utilizo setRowCount(0) porque el método clear() me quita las cabeceras
        int indexRow=0;

        while (!streamIn.atEnd())
        {
            linea = streamIn.readLine();
            linea.replace(",",";");
            linea.replace(":", ";");
            campos = linea.split(";");
            // Crea fila (inserta y añade widgets de texto)
            ui->tableWidget->insertRow(indexRow);
            int maxColumn = campos.size();
            int indexColumn;
            for(indexColumn = 0; indexColumn < maxColumn; ++indexColumn)
            {
                ui->tableWidget->setItem(indexRow, indexColumn, new QTableWidgetItem(campos[indexColumn]));
            }

            // Añade la marca de tiempo con milisegundos
            bool ok;
            unsigned long long s = campos[1].toULongLong( &ok );
            if ( ok )
            {
                const QDateTime dt = QDateTime::fromMSecsSinceEpoch( s );
                QString marca = dt.toString("yyyy-MM-dd hh:mm:ss.zzz");
                ui->tableWidget->setItem(indexRow, indexColumn++, new QTableWidgetItem(marca));
            }

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
