#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include  <QThread>
#include <QWidget>


struct CFG
{
    QString filename_of_config;
    QString sheetname;
    QString excel_path;
};


namespace Ui {
class MainWindow;
class Worker;
}

class Worker:public QThread
{

    Q_OBJECT
public:
    Worker(QWidget *parent=nullptr);
    ~ Worker();
    bool abort=false;
    void stopwork()
    {
        abort=true;
    }
public slots:
    void doWork_save_to_db();
private slots:


private:
    Ui::MainWindow *ui;


};



class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_toolButton_sel_excel_clicked();

    void on_pushButton_read_excel_clicked();

    void on_pushButton_rea_excel_t2_clicked();

    void on_pushButton_search_clicked();

    void on_pushButton_get_column_name_clicked();

    void on_toolButton_select_path_for_db_clicked();

    void on_pushButton_save_excel_to_db_clicked();
    void on_groupBox_search_clicked();
    void on_groupBox_read_clicked();
    void on_comboBox_col_search_currentIndexChanged(int index);

private:
    Ui::MainWindow *ui;
    void load_config();
    void save_config();
    Worker *worker;
    QThread *workerThread;
};

#endif // MAINWINDOW_H
