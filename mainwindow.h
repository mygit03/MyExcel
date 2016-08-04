#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QTableWidget>
#include <QPushButton>

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = 0);
    ~MainWindow();

    void init();

private slots:
    void slot_writeToExcel();

private:
    QTableWidget *tableWidget;
    QPushButton *btn;
};

#endif // MAINWINDOW_H
