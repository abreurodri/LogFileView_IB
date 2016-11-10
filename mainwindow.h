#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_actionAbrir_archivo_triggered();

    void on_actionCopiar_triggered();

    void on_actionExportar_a_Excel_2010_triggered();

    void on_actionAbrir_carpeta_triggered();

    void on_actionDownload_Event_Files_triggered();

private:
    void insert_histfile_on_table(QString const& filePath, int rowPos);

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
