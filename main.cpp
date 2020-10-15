#include <QCoreApplication>
#include <BasicExcel.hpp>
#include <stack>
#include <sstream>
#include <iostream>
#include <QFile>
#include <QDebug>
#include <windows.h>

bool makeChanges(string & filename, bool debugCOUT){
    //Makes needed changes in "xls" file
    //Returns true if successful; otherwise returns false.

    using namespace YExcel;
    stack<string> stack_;

    BasicExcel e_from;
    e_from.Load(&filename[0]);
    vector<size_t> sheets_to_proceed;//Листы, которые нужно обработать
    //Так как есть скрытые листы с другим количеством столбцов, их трогать не нужно
    {
        if (debugCOUT) cout<<"TotalWorkSheets: "<<e_from.GetTotalWorkSheets()<<endl;
        for (size_t i=0; i<e_from.GetTotalWorkSheets(); i++){
            if (debugCOUT) cout<<"Sheet "<<i<<" size: "<<e_from.GetWorksheet(i)->GetTotalRows()<<" rows, "<<e_from.GetWorksheet(i)->GetTotalCols()<<" columns ";
            if (e_from.GetWorksheet(i)->GetTotalCols()>=18){
                sheets_to_proceed.push_back(i);
                if (debugCOUT) cout<<"OK"<<endl;
            } else {
                if (debugCOUT) cout<<"BAD"<<endl;
            }
        }
    }

    if (sheets_to_proceed.empty()){
        if (debugCOUT) cout<<"Processing failed. No worksheets with columns>=18"<<endl;
        return false;
    }

    BasicExcel e_to;
    e_to.New(sheets_to_proceed.size());

    for (size_t i=0; i<sheets_to_proceed.size(); i++){
        BasicExcelWorksheet* sheet_from = e_from.GetWorksheet(sheets_to_proceed[i]);
        BasicExcelWorksheet* sheet_to = e_to.GetWorksheet(i);
        if (!(sheet_from && sheet_to)){
            if (debugCOUT) cout<<"Cannot open worksheet "<<sheets_to_proceed[i]<<endl;
            return false;
        }
        size_t current_row_from = 3, current_row_to = 0;

        //Копирование первых трёх строчек
        for (size_t row=0; row<3; row++){
            for (size_t col=0; col<sheet_from->GetTotalCols(); col++){
                BasicExcelCell* cell = sheet_from->Cell(row, col);
                switch (cell->Type())
                {
                    case BasicExcelCell::INT:
                        sheet_to->Cell(row,col)->SetInteger(cell->GetInteger());
                        break;
                    case BasicExcelCell::DOUBLE:
                        sheet_to->Cell(row,col)->SetDouble(cell->GetDouble());
                        break;
                    case BasicExcelCell::STRING:
                        sheet_to->Cell(row,col)->SetString(cell->GetString());
                        break;
                    case BasicExcelCell::WSTRING:
                        sheet_to->Cell(row,col)->SetWString(cell->GetWString());
                        break;
                }
            }
            current_row_to++;
        }


        for (; current_row_from<sheet_from->GetTotalRows() ;current_row_from++){

            vector<int> columns = {0,1,2,3,4};
            //Копирование единожды
            for (int col : columns)
            {
                BasicExcelCell* cell = sheet_from->Cell(current_row_from, col);
                switch (cell->Type())
                {
                    case BasicExcelCell::INT:
                        sheet_to->Cell(current_row_to,col)->SetInteger(cell->GetInteger());
                        break;
                    case BasicExcelCell::DOUBLE:
                        sheet_to->Cell(current_row_to,col)->SetDouble(cell->GetDouble());
                        break;
                    case BasicExcelCell::STRING:
                        sheet_to->Cell(current_row_to,col)->SetString(cell->GetString());
                        break;
                    case BasicExcelCell::WSTRING:
                        sheet_to->Cell(current_row_to,col)->SetWString(cell->GetWString());
                        break;
                }
            }

            string channel_names_10;//Номер канала (в соответствии со стандартом), 10 столбец
            string freq_names_11;//Частоты ПРД РЭС/ПРМ РЭС, 11 столбец

            //Запись информации в channel_names_10 и freq_names_11
            {
                BasicExcelCell* cell = sheet_from->Cell(current_row_from, 10);
                if (cell->Type()!=BasicExcelCell::STRING){
                    if (debugCOUT) cout<<"cell->Type()!=BasicExcelCell::STRING for ("<<current_row_from<<','<<10<<")"<<endl;
                } else {
                    channel_names_10 = cell->GetString();
                }
            }
            {
                BasicExcelCell* cell = sheet_from->Cell(current_row_from, 11);
                if (cell->Type()!=BasicExcelCell::STRING){
                    if (debugCOUT) cout<<"cell->Type()!=BasicExcelCell::STRING for ("<<current_row_from<<','<<11<<")"<<endl;
                } else {
                    freq_names_11 = cell->GetString();
                }
            }

            //Разбивка информации
            istringstream channel_names_10_ss(channel_names_10);
            istringstream freq_names_11_ss(freq_names_11);
            string channel_10, freq_11;
            while (true){
                channel_names_10_ss>>channel_10;
                freq_names_11_ss>>freq_11;
                if (channel_10.empty() || freq_11.empty()) break;

                stack_.push(move(channel_10));
                sheet_to->Cell(current_row_to,10)->SetString(&stack_.top()[0]);
                stack_.push(move(freq_11));
                sheet_to->Cell(current_row_to,11)->SetString(&stack_.top()[0]);

                columns = {5,6,7,8,9,12,13,14,15,16,17};
                for (int col : columns)
                {
                    BasicExcelCell* cell = sheet_from->Cell(current_row_from, col);
                    switch (cell->Type())
                    {
                        case BasicExcelCell::INT:
                            sheet_to->Cell(current_row_to,col)->SetInteger(cell->GetInteger());
                            break;
                        case BasicExcelCell::DOUBLE:
                            sheet_to->Cell(current_row_to,col)->SetDouble(cell->GetDouble());
                            break;
                        case BasicExcelCell::STRING:
                            sheet_to->Cell(current_row_to,col)->SetString(cell->GetString());
                            break;
                        case BasicExcelCell::WSTRING:
                            sheet_to->Cell(current_row_to,col)->SetWString(cell->GetWString());
                            break;
                    }
                }

                current_row_to++;
            }
        }
    }

    e_to.SaveAs(&filename[0]);

    return true;
}
void qDebugLastError(){
    LPWSTR bufPtr = NULL;
    DWORD err = GetLastError();
    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER |
                   FORMAT_MESSAGE_FROM_SYSTEM |
                   FORMAT_MESSAGE_IGNORE_INSERTS,
                   NULL,err,0,(LPWSTR)&bufPtr,0,NULL);
    const QString result = (bufPtr) ?
                QString::fromUtf16((const ushort*)bufPtr).trimmed() :
                QString("Unknown hz %1").arg(err);
    LocalFree(bufPtr);
    qDebug()<<result;
}
void qDebugCurrentDirectory(){
    QString dir;
    TCHAR path[MAX_PATH];
    GetCurrentDirectory(sizeof (path), path);
    for (int i=0; i<105; i++){
        dir.push_back(static_cast<char>(path[i]));
    }
    qDebug()<<dir;
}
bool changeCurrentDirectory(QString dir){
    LPCWSTR array = (const wchar_t*)dir.utf16();
    return SetCurrentDirectory(array);
}

int main(int argc, char *argv[])
{
    if (argc != 2) return 0;

    bool debugCOUT = false;//Вывод этапов работы главной программы в консоль
    bool debugCOUT_xls = true;//Вывод этапов работы функции изменения xls в консоль
    string filename = "tempfile_213.xls";

    QCoreApplication a(argc, argv);

    //Путь к исходному файлу
        QString filename_to_copy_from = a.arguments().at(1);
        if (debugCOUT) cout<<"filename_to_copy_from[old]: "<<filename_to_copy_from.toStdString()<<endl;
        {
            QString tempQString;
            for (auto qCh : filename_to_copy_from){
                tempQString.push_back(qCh);
                if (qCh == '\\')
                    tempQString.push_back(qCh);
            }
            filename_to_copy_from = tempQString;
            if (debugCOUT) cout<<"filename_to_copy_from[new]: "<<filename_to_copy_from.toStdString()<<endl;
        }

        //Изменение текущей директории на ту, где лежит исходный файл
        {
            if (debugCOUT) qDebugCurrentDirectory();

            QString directory_with_xls;
            int position = filename_to_copy_from.lastIndexOf('\\');
            for (int i=0; i<=position; i++){
                directory_with_xls.push_back(filename_to_copy_from.at(i));
            }
            if (debugCOUT) cout<<"directory_with_exe: "<<directory_with_xls.toStdString()<<endl;

            if (!changeCurrentDirectory(directory_with_xls)){
                if (debugCOUT) qDebugLastError();
                int as;cin>>as;return 0;
            }

            if (debugCOUT) qDebugCurrentDirectory();
        }

        //Путь к файлу с добавкой " разд"
        QString filename_to_copy_to = filename_to_copy_from;
        {
            int position = filename_to_copy_to.lastIndexOf('.');
            QString adding(" разд");
            filename_to_copy_to.insert(position, adding);
            if (debugCOUT) cout<<"filename_to_copy_to: "<<filename_to_copy_to.toStdString()<<endl;
        }

        //Путь к временному файлу
        QString filename_temp;
        {
            int position = filename_to_copy_to.lastIndexOf('\\')+1;
            for (int i=0; i<position; i++){
                filename_temp.push_back(filename_to_copy_to.at(i));
            }
            //filename_temp += "input.xls";//QString::fromStdString(filename);
            filename_temp += QString::fromStdString(filename);//QString::fromStdString(filename);
            if (debugCOUT) cout<<"filename_temp: "<<filename_temp.toStdString()<<endl;
        }

        //Создание временного файла
        if (QFile::exists(filename_temp)){
            QFile::remove(filename_temp);
        }
        if(!QFile::copy(filename_to_copy_from ,filename_temp)){
            if (debugCOUT) cout<<"First copy failed"<<endl;
            if (debugCOUT) cout<<"From "<<filename_to_copy_from.toStdString()<<endl;
            if (debugCOUT) cout<<"To   "<<filename_temp.toStdString()<<endl;
            int as;cin>>as; return 0;
        }

        //Обработка временного файла
        if (makeChanges(filename, debugCOUT_xls)){
            if (debugCOUT) cout<<"makeChanges OK"<<endl;
        } else {
            if (debugCOUT) cout<<"makeChanges FAILED!"<<endl;
            int as;cin>>as; return 0;
        }

        //Создание выходного файла
        if (QFile::exists(filename_to_copy_to)){
            QFile::remove(filename_to_copy_to);
        }
        if(!QFile::copy(filename_temp ,filename_to_copy_to)){
            if (debugCOUT) cout<<"Second copy failed"<<endl;
            if (debugCOUT) cout<<"From "<<filename_temp.toStdString()<<endl;
            if (debugCOUT) cout<<"To   "<<filename_to_copy_to.toStdString()<<endl;
            int as;cin>>as; return 0;
        }
        QFile::remove(filename_temp);

        if (debugCOUT) cout<<"EVERYTHING IS OK"<<endl;

        return 0;
}
