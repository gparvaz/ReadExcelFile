#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <stdio.h>
#include <QFileDialog>
#include <QDebug>
#include <QTableWidgetItem>
#include <QMessageBox>
#include <QtSql/QSql>
#include <QtSql/QSqlDatabase>
#include <QSqlQuery>
#include <QSqlDriver>
#include <QSql>
#include <QSettings>
#include <QSqlDatabase>
#include <QObject>
#include <QProcess>
#include <QAxObject>
#include <QStringListModel>
#include <QTimer>

static CFG config;
static QStringList dataRow1_for_write_in_db;
static QStringList dataRow2_for_write_in_db;
static QStringList dataRow3_for_write_in_db;
static QString path_for_saving_db;
static QString table_name_of_db;

static QLCDNumber *lcd_scan;
static int i_scan,scan_counter=0;
void updateMCvalues(int t);



Worker::Worker(QWidget *parent)
{

}
Worker::~Worker()
{
    abort=true;
    wait();
}

void Worker::doWork_save_to_db()
{
    forever
    {
        if(abort)
        {
            qDebug() <<"6" ;
            return;
        }
        else if(i_scan <dataRow1_for_write_in_db.size())
        {
            qDebug() <<"7" ;

            // Create database.
            QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "Connection");
            db.setDatabaseName(path_for_saving_db);
            if (!db.open())
            {
                qDebug("Error occurred opening the database.");
            }
            // Insert table in DB.
            QString table_name="CREATE TABLE IF NOT EXISTS "+table_name_of_db+" (Column1 TEXT, Column2 TEXT, Column3 TEXT)" ;
            QSqlQuery query(db);
            query.prepare(table_name);
            if (!query.exec())
            {
                qDebug("Error occurred creating table.");
            }

            // Query DB.
            QString query_table="SELECT * FROM "+table_name_of_db;
            query.prepare(query_table);
            if (!query.exec())
            {
                qDebug("Error occurred querying.");
            }

            // Insert row in created Table.
            for (int iter=0;iter<dataRow1_for_write_in_db.size() ;iter++)
            {
                lcd_scan->display(iter);
                QString query_row="INSERT INTO "+table_name_of_db+" (Column1, Column2, Column3) VALUES (\'"+dataRow1_for_write_in_db[iter]+"\', \'"+dataRow2_for_write_in_db[iter]+"\', \'"+dataRow3_for_write_in_db[iter]+"\')";

                bool inserted = query.exec(query_row);
                if(inserted == false)
                {
                    qDebug() << "Error-RRRROOORRRRRRRRE";
                }
            }
            break;
            abort=true;
        }
        else
        {
            break;
            abort=true;
        }
    }
}


MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    config.filename_of_config=QString("config.ini");
    load_config();
    lcd_scan=ui->lcdNumber_save_to_db;
    //
    ui->lineEdit_table_name->setText("Table"+QString::number(qrand() % 10));
    QString outputpath = ui->lineEdit_excel->text();
    QFile f(outputpath);
    QFileInfo fileInfo(f.fileName());
    QString filename_of_files_with_extention(fileInfo.fileName());
    QString filepath_of_files=outputpath.remove(filename_of_files_with_extention);
    QString filename_of_files=filename_of_files_with_extention.replace(".","_");
    QString file_path_name=filepath_of_files+filename_of_files+".db";
    ui->lineEdit_path_save_db->setText(file_path_name);
    //
    ui->label_sheet_no->setHidden(1);
    ui->spinBox_sheet_no->setHidden(1);
    ui->spinBox_max_empth_row->setHidden(1);
    ui->label_max_empty->setHidden(1);
    ui->comboBox_col3->setHidden(1);
    ui->label_col3->setHidden(1);
    ui->comboBox_col2->setHidden(1);
    ui->label_col2->setHidden(1);
    ui->comboBox_col1->setHidden(1);
    ui->label_col1->setHidden(1);
    //
    ui->label_sheet_name->setHidden(1);
    ui->lineEdit_sheet->setHidden(1);
    ui->pushButton_read_excel->setHidden(1);
    ui->pushButton_rea_excel_t2->setHidden(1);
    ui->line_seprator->setHidden(1);
    //
    ui->groupBox_read->setHidden(1);
    ui->groupBox_search->setHidden(1);
    ui->tableWidget->setHidden(1);
    //
    ui->label_word_search->setHidden(1);
    ui->lineEdit_searchword->setHidden(1);
    ui->label_col_search->setHidden(1);
    ui->comboBox_col_search->setHidden(1);
    ui->pushButton_search->setHidden(1);
    ui->lcdNumber_search->setHidden(1);
    //
    ui->label_db_name->setHidden(1);
    ui->lineEdit_path_save_db->setHidden(1);
    ui->toolButton_select_path_for_db->setHidden(1);
    ui->label_table_db->setHidden(1);
    ui->lineEdit_table_name->setHidden(1);
    ui->pushButton_save_excel_to_db->setHidden(1);
    ui->lcdNumber_save_to_db->setHidden(1);


}

MainWindow::~MainWindow()
{
    delete ui;
    save_config();
}


void MainWindow::load_config()
{
    QString FILENAME(config.filename_of_config);
    QSettings settings(FILENAME,QSettings::IniFormat);
    //read
    config.sheetname=settings.value("Main/Input_Path4",QString()).toString();
    config.excel_path=settings.value("Main/Input_Path5",QString()).toString();
    // //////////////////////////////////////////////////////////////////////////////
    ui->lineEdit_sheet->setText(config.sheetname);
    ui->lineEdit_excel->setText(config.excel_path);

}
void MainWindow::save_config()
{
    QString FILENAME(config.filename_of_config);
    QSettings settings(FILENAME,QSettings::IniFormat);
    settings.setValue("Main/Input_Path4",ui->lineEdit_sheet->text());
    settings.setValue("Main/Input_Path5",ui->lineEdit_excel->text());


}
void MainWindow::on_toolButton_sel_excel_clicked()
{
    QString fileName = QFileDialog::getOpenFileName(this, tr("Select File"),
                                                    "../list/",
                                                    tr("Excel File(*.xls *.xlsx *.xlsm *.xlsb)"));
    ui->lineEdit_excel->setText(fileName);
    //
    ui->lineEdit_table_name->setText("Table"+QString::number(qrand() % 10));
    QString outputpath = fileName;
    QFile f(outputpath);
    QFileInfo fileInfo(f.fileName());
    QString filename_of_files_with_extention(fileInfo.fileName());
    QString filepath_of_files=outputpath.remove(filename_of_files_with_extention);
    QString filename_of_files=filename_of_files_with_extention.replace(".","_");
    QString file_path_name=filepath_of_files+filename_of_files+".db";
    ui->lineEdit_path_save_db->setText(file_path_name);
}

void MainWindow::on_pushButton_get_column_name_clicked()
{
    if(ui->lineEdit_excel->text()!="")
    {
        ui->label_sheet_no->setHidden(0);
        ui->spinBox_sheet_no->setHidden(0);
        ui->spinBox_max_empth_row->setHidden(0);
        ui->label_max_empty->setHidden(0);
        ui->comboBox_col3->setHidden(0);
        ui->label_col3->setHidden(0);
        ui->comboBox_col2->setHidden(0);
        ui->label_col2->setHidden(0);
        ui->comboBox_col1->setHidden(0);
        ui->label_col1->setHidden(0);
        ui->groupBox_read->setHidden(0);
        ui->groupBox_search->setHidden(0);
        ui->tableWidget->setHidden(0);
        QString excel_path_name=ui->lineEdit_excel->text();
        int sheetnumber=ui->spinBox_sheet_no->value();

        QAxObject* excel     = new QAxObject( "Excel.Application");
        QAxObject* workbooks = excel->querySubObject( "Workbooks" );
        QAxObject* workbook  = workbooks->querySubObject( "Open(const QString&)",excel_path_name);
        QAxObject* sheets    = workbook->querySubObject( "Worksheets" );
        QAxObject* sheet     = sheets->querySubObject( "Item( int )", sheetnumber );

        QAxObject* columns = sheet->querySubObject( "Columns" );
        int columnCount    = columns->dynamicCall( "Count()" ).toInt(); //similarly, always returns 65535

        //One of possible ways to get column count
        int currentColumnCount = 0;
        for (int col=1; col<columnCount; col++)
        {
            QAxObject* cell = sheet->querySubObject( "Cells( int, int )", 1, col );
            QVariant value  = cell->dynamicCall( "Value()" );
            if (value.toString().isEmpty())
                break;
            else
                currentColumnCount = col;
        }
        columnCount = currentColumnCount;
        QStringList dataRow;


        for (int column=1; column <= columnCount; column++)
        {
            QAxObject* cell = sheet->querySubObject( "Cells( int, int )", 1, column );
            QString value = cell->dynamicCall( "Value()" ).toString();
            dataRow.append(value);
        }


        QString guid_of_columns_name;
        for (int p=0;p< dataRow.size();p++)
        {
            guid_of_columns_name.append(QString::number(p)+":    ");
            guid_of_columns_name.append(dataRow[p]);
            guid_of_columns_name.append("\n");

        }
        ui->pushButton_read_excel->setToolTip(guid_of_columns_name);
        ui->comboBox_col1->setToolTip(guid_of_columns_name);
        ui->comboBox_col2->setToolTip(guid_of_columns_name);
        ui->comboBox_col3->setToolTip(guid_of_columns_name);
        //Instance of model type to QStringList
        QStringListModel *model = new QStringListModel();
        model->setStringList(dataRow);

        ui->comboBox_col1->setModel(model);
        ui->comboBox_col2->setModel(model);
        ui->comboBox_col3->setModel(model);
        ui->comboBox_col_search->setModel(model);

        workbook->dynamicCall("Close()");
        excel->dynamicCall("Quit()");

    }
    else
    {
        QMessageBox msgBox_war;
        msgBox_war.setWindowTitle("     Status      ");
        msgBox_war.setInformativeText("Fill Inputpath");
        msgBox_war.exec();
    }


}

void MainWindow::on_pushButton_read_excel_clicked()
{
    if(ui->lineEdit_excel->text()!="" && ui->lineEdit_sheet->text()!="" && ui->comboBox_col1->currentIndex()!=-1)
    {
        QString sheetName=ui->lineEdit_sheet->text();
        int col1_for_read=ui->comboBox_col1->currentIndex();
        int col2_for_read=ui->comboBox_col2->currentIndex();
        int col3_for_read=ui->comboBox_col3->currentIndex();
        QString excel_file_path=ui->lineEdit_excel->text();

        qDebug() << "Open excel start" << endl;
        QSqlDatabase db = QSqlDatabase::addDatabase("QODBC","xlsx_connection");
        db.setDatabaseName("DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + excel_file_path);
        if(db.open())
        {
            QString column1;
            QString column2;
            QString column3;
            QStringList column1_list;
            QStringList column2_list;
            QStringList column3_list;


            qDebug() << "Can open file (Excel)" << endl;

            QSqlQuery query2 = *new QSqlQuery(db);
            query2.exec("select * from [" + sheetName + "$]"); // $A1:B5 or $

            while (query2.next())
            {
                column1.append(query2.value(col1_for_read).toString());
                column1.append("\n");
                column2.append(query2.value(col2_for_read).toString());
                column2.append("\n");
                column3.append(query2.value(col3_for_read).toString());
                column3.append("\n");
            }
            db.close();
            //  convert string1 to list1
            QStringList strlistcolumn1= column1.split("\n");
            for (int i=0;i<strlistcolumn1.size();++i)
            {
                column1_list<<strlistcolumn1[i];
            }
            //  convert string2 to list2
            QStringList strlistcolumn2= column2.split("\n");
            for (int i=0;i<strlistcolumn2.size();++i)
            {
                column2_list<< strlistcolumn2[i];
            }
            //  convert string3 to list3
            QStringList strlistcolumn3= column3.split("\n");
            for (int i=0;i<strlistcolumn3.size();++i)
            {
                column3_list <<strlistcolumn3[i];
            }

            // Widget1
            dataRow1_for_write_in_db=column1_list;
            dataRow2_for_write_in_db=column2_list;
            dataRow3_for_write_in_db=column3_list;
            ui->tableWidget->setRowCount(column1_list.size());
            ui->tableWidget->setColumnCount(3);
            for(int i=0; i<column1_list.size();i++ )
            {
                QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(column1_list[i]));
                ui->tableWidget->setItem(i,0,newItem);
            }
            // Widget2
            for(int i=0; i<column2_list.size();i++ )
            {
                QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(column2_list[i]));
                ui->tableWidget->setItem(i,1,newItem);
            }
            // Widget3
            for(int i=0; i<column3_list.size();i++ )
            {
                QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(column3_list[i]));
                ui->tableWidget->setItem(i,2,newItem);
            }


        }
        else
        {
            qDebug() << "Can't open file (Excel)" << endl;
        }
    }
    else
    {
        QMessageBox msgBox_war;
        msgBox_war.setWindowTitle("***********     Status     *****************");
        msgBox_war.setInformativeText("Fill Inputpath and sheetname and column");
        msgBox_war.exec();
    }
}

void MainWindow::on_pushButton_rea_excel_t2_clicked()
{
    if(ui->lineEdit_excel->text()!="" && ui->comboBox_col1->currentIndex()!=-1)
    {
        QString excel_path_name=ui->lineEdit_excel->text();
        int sheetnumber=ui->spinBox_sheet_no->value();

        QAxObject* excel     = new QAxObject( "Excel.Application");
        QAxObject* workbooks = excel->querySubObject( "Workbooks" );
        QAxObject* workbook  = workbooks->querySubObject( "Open(const QString&)",excel_path_name);
        QAxObject* sheets    = workbook->querySubObject( "Worksheets" );
        QAxObject* sheet     = sheets->querySubObject( "Item( int )", sheetnumber );

        qDebug() << sheet;



        qDebug () << "1";
        // convert columnname to int :
        int column_1_for_write_in_table=ui->comboBox_col1->currentIndex()+1;
        int column_2_for_write_in_table=ui->comboBox_col2->currentIndex()+1;
        int column_3_for_write_in_table=ui->comboBox_col3->currentIndex()+1;
        int chance_for_end=0;
        QStringList dataRow1;
        QStringList dataRow2;
        QStringList dataRow3;
        for (int row=1; row <= 65535; row++)
        {
            // Column 1
            QAxObject* cell1 = sheet->querySubObject( "Cells( int, int )", row, column_1_for_write_in_table );
            QString value1   = cell1->dynamicCall( "Value()" ).toString();
            // Column 2
            QAxObject* cell2 = sheet->querySubObject( "Cells( int, int )", row, column_2_for_write_in_table );
            QString value2   = cell2->dynamicCall( "Value()" ).toString();
            // Column 3
            QAxObject* cell3 = sheet->querySubObject( "Cells( int, int )", row, column_3_for_write_in_table );
            QString value3   = cell3->dynamicCall( "Value()" ).toString();
            // write and find end of excel
            if (value1=="" && value2=="" && value3=="")
            {
                chance_for_end++;
            }
            else
            {
                dataRow1 << value1;
                dataRow2 << value2;
                dataRow3 << value3;
            }
            if (chance_for_end>ui->spinBox_max_empth_row->value())
            {
                break;
            }
        }
        dataRow1_for_write_in_db= dataRow1;
        dataRow2_for_write_in_db= dataRow2;
        dataRow3_for_write_in_db= dataRow3;

        qDebug () << "2";

        // Widget1
        ui->tableWidget->setRowCount(dataRow1.size());
        ui->tableWidget->setColumnCount(3);
        for(int i=0; i<dataRow1.size();i++ )
        {
            QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(dataRow1[i]));
            ui->tableWidget->setItem(i,0,newItem);
        }
        // Widget2
        for(int i=0; i<dataRow2.size();i++ )
        {
            QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(dataRow2[i]));
            ui->tableWidget->setItem(i,1,newItem);
        }
        // Widget3
        for(int i=0; i<dataRow3.size();i++ )
        {
            QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(dataRow3[i]));
            ui->tableWidget->setItem(i,2,newItem);
        }


        workbook->dynamicCall("Close()");
        excel->dynamicCall("Quit()");

    }
    else
    {
        QMessageBox msgBox_war;
        msgBox_war.setWindowTitle("************ Status ************");
        msgBox_war.setInformativeText("Fill Inputpath and Select Columns");
        msgBox_war.exec();
    }
}

void MainWindow::on_pushButton_search_clicked()
{
    if(ui->lineEdit_excel->text()!="" && ui->lineEdit_searchword->text()!="")
    {
        QString excel_path_name=ui->lineEdit_excel->text();
        int sheetnumber=ui->spinBox_sheet_no->value();
        // find a word
        QString search_word =ui->lineEdit_searchword->text();
        QString column_number;


        auto excel     = new QAxObject("Excel.Application");
        auto workbooks = excel->querySubObject("Workbooks");
        auto workbook  = workbooks->querySubObject("Open(const QString&)",excel_path_name);
        auto sheets    = workbook->querySubObject("Worksheets");
        auto sheet     = sheets->querySubObject("Item(int)", sheetnumber);
        // find column and conver it to range
        QString column_name=ui->comboBox_col_search->currentText();
        auto range_column_name = sheet->querySubObject("Range(A1,AZ1)");//60 columns
        auto find_column_name = range_column_name->querySubObject("Find(const QString&)",column_name);
        if (nullptr != find_column_name)
        {
            column_number=find_column_name->dynamicCall("Address").toString();
            column_number.remove(2,2);
            column_number.remove("$");
        }

        QStringList rows_contains_search_word;
        int real_row_count_incolumn=0;
        int column_1_for_write_in_table=ui->comboBox_col1->currentIndex()+1;
        int column_2_for_write_in_table=ui->comboBox_col2->currentIndex()+1;
        int column_3_for_write_in_table=ui->comboBox_col3->currentIndex()+1;
        QStringList dataRow1_search;
        QStringList dataRow2_search;
        QStringList dataRow3_search;

        for (int row=1; row <= 65535; row++)
        {
            bool isEmpty = true;
            int column=ui->comboBox_col_search->currentIndex()+1;
            QAxObject* cell = sheet->querySubObject( "Cells( int, int )", row, column );
            QVariant value = cell->dynamicCall( "Value()" );
            if (!value.toString().isEmpty() && isEmpty)
            {
                isEmpty = false;
            }
            real_row_count_incolumn++;
            if (isEmpty) //criteria to get out of cycle
                break;
        }
        for(int t=1 ; t <real_row_count_incolumn-1;t++)
        {
            QString convert_column_number_to_range=column_number;
            convert_column_number_to_range.append(QString::number(t));
            convert_column_number_to_range.append(",");
            convert_column_number_to_range.append(column_number);
            convert_column_number_to_range.append("65535"); //max of rows in a column
            QString Range_For_Search="Range("+convert_column_number_to_range+")";

            QByteArray Range_char = Range_For_Search.toLocal8Bit();
            const char *Range_char_data = Range_char.data();

            auto range     = sheet->querySubObject(Range_char_data);
            auto find      = range->querySubObject("Find(const QString&)",search_word);

            if (nullptr != find)
            {
                QString row_nomber_in_selected_column=find->dynamicCall("Address").toString();
                row_nomber_in_selected_column.remove(0,3);
                //t=row_nomber_in_selected_column.toInt()-1;
                rows_contains_search_word <<(row_nomber_in_selected_column);
            }
        }
        rows_contains_search_word.removeDuplicates();

        ui->lcdNumber_search->display(rows_contains_search_word.size()-2);
        for (int row=0; row <= rows_contains_search_word.size()-2; row++)
        {
            // Column 1
            QAxObject* cell1 = sheet->querySubObject( "Cells( int, int )", rows_contains_search_word[row], column_1_for_write_in_table );
            QString value1   = cell1->dynamicCall( "Value()" ).toString();
            dataRow1_search << value1;
            // Column 2
            QAxObject* cell2 = sheet->querySubObject( "Cells( int, int )", rows_contains_search_word[row], column_2_for_write_in_table );
            QString value2   = cell2->dynamicCall( "Value()" ).toString();
            dataRow2_search << value2;
            // Column 3
            QAxObject* cell3 = sheet->querySubObject( "Cells( int, int )", rows_contains_search_word[row], column_3_for_write_in_table );
            QString value3   = cell3->dynamicCall( "Value()" ).toString();
            dataRow3_search << value3;
        }
        dataRow1_for_write_in_db=dataRow1_search;
        dataRow2_for_write_in_db=dataRow2_search;
        dataRow3_for_write_in_db=dataRow3_search;

        // Widget0
        ui->tableWidget->setRowCount(rows_contains_search_word.size());
        ui->tableWidget->setColumnCount(4);
        for(int i=0; i<rows_contains_search_word.size();i++ )
        {
            QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(rows_contains_search_word[i]));
            ui->tableWidget->setItem(i,0,newItem);
        }
        // Widget1
        for(int i=0; i<dataRow1_search.size();i++ )
        {
            QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(dataRow1_search[i]));
            ui->tableWidget->setItem(i,1,newItem);
        }
        // Widget2
        for(int i=0; i<dataRow2_search.size();i++ )
        {
            QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(dataRow2_search[i]));
            ui->tableWidget->setItem(i,2,newItem);
        }
        // Widget3
        for(int i=0; i<dataRow3_search.size();i++ )
        {
            QTableWidgetItem *newItem = new QTableWidgetItem(tr("%1").arg(dataRow3_search[i]));
            ui->tableWidget->setItem(i,3,newItem);
        }





        // don't forget to quit Excel
        excel->dynamicCall("Quit()");
        delete excel;
    }
    else
    {
        QMessageBox msgBox_war;
        msgBox_war.setWindowTitle("********     Status      ********");
        msgBox_war.setInformativeText("Fill Inputpath and search word");
        msgBox_war.exec();
    }
}


void MainWindow::on_toolButton_select_path_for_db_clicked()
{
    if (ui->lineEdit_excel->text()!="")
    {
        QString dir_db = QFileDialog::getExistingDirectory(this, tr("Open Directory"),
                                                           "../",
                                                           QFileDialog::ShowDirsOnly
                                                           | QFileDialog::DontResolveSymlinks);


        QString outputpath = ui->lineEdit_excel->text();
        QFile f(outputpath);
        QFileInfo fileInfo(f.fileName());
        QString filename_of_files_with_extention(fileInfo.fileName());
        QString filename_of_files=filename_of_files_with_extention.replace(".","_");
        dir_db.append("/");
        dir_db.append(filename_of_files);
        dir_db.append(".db");
        ui->lineEdit_path_save_db->setText(dir_db);
    }
    else
    {
        QMessageBox msgBox_war;
        msgBox_war.setWindowTitle    ("*       Status      *");
        msgBox_war.setInformativeText("Fill Inputpath ");
        msgBox_war.exec();
    }
}

void MainWindow::on_pushButton_save_excel_to_db_clicked()
{
    if(ui->lineEdit_excel->text()!="" && ui->lineEdit_path_save_db->text()!="" && ui->lineEdit_table_name->text()!=""&&ui->tableWidget->columnCount()!=0)
    {
        path_for_saving_db=ui->lineEdit_path_save_db->text();
        table_name_of_db=ui->lineEdit_table_name->text();

        worker =new Worker();
        workerThread=new QThread(this);

        connect(workerThread,SIGNAL(started()),worker,SLOT(doWork_save_to_db()));
        connect(workerThread,SIGNAL(finished()),worker,SLOT(deleteLater()));

        worker->moveToThread(workerThread);
        workerThread->start();

    }
    else
    {
        QMessageBox msgBox_war;
        msgBox_war.setWindowTitle    ("******************* Status ******************");
        msgBox_war.setInformativeText("Fill Inputpath and DatabaseName and TableName");
        msgBox_war.exec();



    }
}

void MainWindow::on_groupBox_search_clicked()
{
    if(ui->groupBox_search->isChecked() ==1)
    {
        ui->label_word_search->setHidden(0);
        ui->lineEdit_searchword->setHidden(0);
        ui->label_col_search->setHidden(0);
        ui->comboBox_col_search->setHidden(0);
        ui->pushButton_search->setHidden(0);
        ui->lcdNumber_search->setHidden(0);
        //
        ui->label_db_name->setHidden(0);
        ui->lineEdit_path_save_db->setHidden(0);
        ui->toolButton_select_path_for_db->setHidden(0);
        ui->label_table_db->setHidden(0);
        ui->lineEdit_table_name->setHidden(0);
        ui->pushButton_save_excel_to_db->setHidden(0);
        ui->lcdNumber_save_to_db->setHidden(0);
    }
    else
    {
        ui->label_word_search->setHidden(1);
        ui->lineEdit_searchword->setHidden(1);
        ui->label_col_search->setHidden(1);
        ui->comboBox_col_search->setHidden(1);
        ui->pushButton_search->setHidden(1);
        ui->lcdNumber_search->setHidden(1);
        //
        ui->label_db_name->setHidden(1);
        ui->lineEdit_path_save_db->setHidden(1);
        ui->toolButton_select_path_for_db->setHidden(1);
        ui->label_table_db->setHidden(1);
        ui->lineEdit_table_name->setHidden(1);
        ui->pushButton_save_excel_to_db->setHidden(1);
        ui->lcdNumber_save_to_db->setHidden(1);
    }
}
void MainWindow::on_groupBox_read_clicked()
{
    if(ui->groupBox_read->isChecked() ==1)
    {
        ui->label_sheet_name->setHidden(0);
        ui->lineEdit_sheet->setHidden(0);
        ui->pushButton_read_excel->setHidden(0);
        ui->pushButton_rea_excel_t2->setHidden(0);
        ui->line_seprator->setHidden(0);
    }
    else
    {
        ui->label_sheet_name->setHidden(1);
        ui->lineEdit_sheet->setHidden(1);
        ui->pushButton_read_excel->setHidden(1);
        ui->pushButton_rea_excel_t2->setHidden(1);
        ui->line_seprator->setHidden(1);
    }

}
void MainWindow::on_comboBox_col_search_currentIndexChanged(int index)
{
    ui->comboBox_col1->setCurrentIndex(index);
    ui->comboBox_col2->setCurrentIndex(index+2);
    ui->comboBox_col3->setCurrentIndex(index+3);
    ui->lcdNumber_search->display(0);

}

// *******************************************************************************************
// *******************************************************************************************
// *****************************      Don't Delete         ***********************************
// *******************************************************************************************
// *******************************************************************************************
//        QAxObject* columns = sheet->querySubObject( "Columns" );
//        int columnCount = columns->dynamicCall( "Count()" ).toInt(); //similarly, always returns 65535

//        //One of possible ways to get column count
//        int currentColumnCount = 0;
//        for (int col=1; col<columnCount; col++)
//        {
//            QAxObject* cell = sheet->querySubObject( "Cells( int, int )", 1, col );
//            QVariant value = cell->dynamicCall( "Value()" );
//            if (value.toString().isEmpty())
//                break;
//            else
//                currentColumnCount = col;
//        }
//        columnCount = currentColumnCount;




// // Create database.
//    QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "Connection");
//    db.setDatabaseName(ui->lineEdit_db->text());
//    if (!db.open())
//    {
//        qDebug("Error occurred opening the database.");
//        //qDebug("%s.", qPrintable(db.lastError().text()));

//    }

//    // Insert table.
//    QSqlQuery query(db);
//    query.prepare("CREATE TABLE IF NOT EXISTS test (id INTEGER PRIMARY KEY, text TEXT)");
//    if (!query.exec())
//    {
//        qDebug("Error occurred creating table.");

//    }

//    // Insert row.
//    query.prepare("INSERT INTO test VALUES (null, ?)");
//    query.addBindValue("Some text");
//    if (!query.exec())
//    {
//        qDebug("Error occurred inserting.");

//    }

//    // Query.
//    query.prepare("SELECT * FROM test");
//    if (!query.exec())
//    {
//        qDebug("Error occurred querying.");


//    }
//    while (query.next())
//    {
//        qDebug("id = %d, text = %s.",
//               query.value(0).toInt(),
//               qPrintable(query.value(1).toString()));
//    }






