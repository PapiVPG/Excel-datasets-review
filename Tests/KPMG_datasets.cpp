#include <iostream>
#include "KPMG_datasets.hpp"
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <regex>
using namespace std;

Customer_Demographic::Customer_Demographic():
    length(4002)
{
    check_id();
    check_first_name();
    check_last_name();
    check_gender();
    check_past_3_years();
    check_date();
    check_job_tittle();
    checK_job_industry_category();
    check_wealth_segment();
    check_deceased_indicator();
    check_defalute();
    check_own_cars();
    check_tenure();
}

Transactions::Transactions():
    length(5082)    
{
    check_id_trans();
    check_id_product();
    check_id_customer();
    check_date();
    check_online_order();
    check_order_status();
    check_brand();
    check_produkt_line();
    check_produkt_class();
    check_produkt_size();
    check_list_price();
    check_standard_cost();
    check_first_date();
}

CustomerAddress::CustomerAddress():
    length(4001)
{
    check_id();
    check_address();
    check_post_code();
    check_state();
    check_country();
    check_valuations();
}


bool czy_wylosowana(int licznik,int wyl_liczba,int tab[]){
    if(licznik<=0){
        return false;
    }
    for(int i = 0; i<licznik;i++){
        if(wyl_liczba==tab[i]){
        return true;
    }
    }
    return false;
}

void Customer_Demographic::check_id()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto id = demo_sheet.cell("A2").to_string();
    this->id=id;
    int r;
    int counter = 0;
    file.open("All_errors (Customer_Demographic).txt",ios::out);
    file<<this->id<<" :"<<endl;
    for(int i = 3; i<=this->length;i++){
        auto b =demo_sheet.cell(xlnt::cell_reference("A", i)).value<int>();
        r=i-2;
        if(b!=r)
        {
            counter++;
            //cout<<"Something wrong in cell: A"<<b<<endl;
            file<<"Check cell: A"<<b<<endl;
        }
        //cout<<b<<" ";
    }
    if(counter==0)
    {
        file<<this->id<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
   
    file<<"----------------------------------------------------------------------------"<<endl;
        
    file.close();
}

void Customer_Demographic::check_first_name()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto fname = demo_sheet.cell("B2").to_string();
    this->first_name=fname;
    int counter = 0;
    regex key("^[a-zA-Z' -]+$");
    regex empty("^$");
    smatch match;
    
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->first_name<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =demo_sheet.cell(xlnt::cell_reference("B", i)).to_string();
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell B"<<i<<endl;
            }
           
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("B", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[B"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->first_name<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
        file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::check_last_name()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto fname = demo_sheet.cell("C2").to_string();
    this->last_name=fname;
    int counter = 0;
    regex key("^[a-zA-Z' -]+$");
    regex empty("^$");
    smatch match;
    
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->last_name<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =demo_sheet.cell(xlnt::cell_reference("C", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell C"<<i<<endl;
            }
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("C", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[C"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->last_name<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
        file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::check_gender()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto fname = demo_sheet.cell("D2").to_string();
    this->gender=fname;
    int counter = 0;
    regex key("\\b(Male|Female|U)\\w?\\b");
    regex key2("^[a-zA-Z]+$");
    regex empty("^$");
    smatch match;
    
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->gender<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =demo_sheet.cell(xlnt::cell_reference("D", i)).to_string();
            if(regex_search(b,match,empty))
            {
                continue;
            }
            if(!regex_search(b,match,key2))
            {
                ++counter;
                file<<"Check cell D"<<i<<endl;
                continue;
            }
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell D"<<i<<endl;
            }
            
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("D", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[D"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->gender<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}

void Customer_Demographic::check_past_3_years()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto id = demo_sheet.cell("E2").to_string();
    this->past_3_years=id;
    int counter = 0;
    smatch match;
    regex key("^[0-9]+$");
    regex empty("^$");
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->past_3_years<<" :"<<endl;
    for(int i = 3; i<=this->length;i++){
        auto b =demo_sheet.cell(xlnt::cell_reference("E", i)).to_string();
         if(regex_search(b,match,empty))
            {
                continue;
            }
           
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell E"<<i<<endl;
            }
            
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("E", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[E"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->past_3_years<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::check_date()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto id = demo_sheet.cell("F2").to_string();
    this->date=id;
    int counter = 0;
    smatch match;
    regex key("^[0-9]+$");
    regex empty("^$");
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->date<<" :"<<endl;
    for(int i = 3; i<=this->length;i++){
        auto b =demo_sheet.cell(xlnt::cell_reference("F", i)).to_string();
         if(regex_search(b,match,empty))
             continue;
           
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell F"<<i<<endl;
            }
            
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("F", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[F"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->date<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}

void Customer_Demographic::check_job_tittle()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto id = demo_sheet.cell("G2").to_string();
    this->job_tittle=id;
    int counter = 0;
    smatch match;
    regex key("^[a-z /A-Z]+$");
    regex empty("^$");
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->job_tittle<<" :"<<endl;
    for(int i = 3; i<=this->length;i++){
        auto b =demo_sheet.cell(xlnt::cell_reference("G", i)).to_string();
         if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell G"<<i<<endl;
            }
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("G", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[G"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->job_tittle<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::checK_job_industry_category()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto id = demo_sheet.cell("H2").to_string();
    this->job_industry_category=id;
    int counter = 0;
    smatch match;
    regex key("^[a-z /A-Z]+$");
    regex empty("^$");
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->job_industry_category<<" :"<<endl;
    for(int i = 3; i<=this->length;i++){
        auto b =demo_sheet.cell(xlnt::cell_reference("H", i)).to_string();
         if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell H"<<i<<endl;
            }
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("H", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[H"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->job_industry_category<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::check_wealth_segment()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto id = demo_sheet.cell("I2").to_string();
    this->wealth_segment=id;
    int counter = 0;
    smatch match;
    regex key("^[a-z /A-Z]+$");
    regex empty("^$");
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->wealth_segment<<" :"<<endl;
    for(int i = 3; i<=this->length;i++){
        auto b =demo_sheet.cell(xlnt::cell_reference("I", i)).to_string();
         if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell I"<<i<<endl;
            }
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("I", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[I"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->wealth_segment<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::check_deceased_indicator()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto fname = demo_sheet.cell("J2").to_string();
    this->deceased_indicator=fname;
    int counter = 0;
    regex key("^[Y|N]{1}+$");
    //regex key2("^[a-zA-Z]{1}+$");
    regex empty("^$");
    smatch match;
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->deceased_indicator<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =demo_sheet.cell(xlnt::cell_reference("J", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell J"<<i<<endl;
            }
            
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("J", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[J"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->deceased_indicator<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::check_defalute()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto id = demo_sheet.cell("K2").to_string();
    this->defaulte=id;
    int counter = 0;
    smatch match;
    regex key("^.+$");
    regex empty("^$");
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->defaulte<<" :"<<endl;
    for(int i = 3; i<=this->length;i++){
        auto b =demo_sheet.cell(xlnt::cell_reference("K", i)).to_string();
         if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell K"<<i<<endl;
            }
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("K", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[K"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->defaulte<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::check_own_cars()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto fname = demo_sheet.cell("L2").to_string();
    this->own_cars=fname;
    int counter = 0;
    regex key("\\b([Yes]{3}|[No]{2})\\w?\\b");
    regex empty("^$");
    smatch match;
    
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->own_cars<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =demo_sheet.cell(xlnt::cell_reference("L", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell L"<<i<<endl;
            }
            
    }
    file<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("L", j)).to_string();
        if(regex_search(cell,match,empty))
        {
            file<<"[L"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->own_cars<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}
void Customer_Demographic::check_tenure()
{
    fstream file;
    xlnt::workbook workbook_one;
    workbook_one.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet demo_sheet = workbook_one.sheet_by_title("CustomerDemographic");
    auto fname = demo_sheet.cell("M2").to_string();
    this->tenure=fname;
    int counter = 0;
    regex key("^[0-9]{1,2}+$");
    regex empty("^$");
    smatch match;
    
    file.open("All_errors (Customer_Demographic).txt",ios::out|ios::app);
    file<<this->tenure<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =demo_sheet.cell(xlnt::cell_reference("M", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file<<"Check cell M"<<i<<endl;
            }
    }
    file<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =demo_sheet.cell(xlnt::cell_reference("M", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file<<"[M"<<j<<"] ";
        }
    }
    file<<endl;
    if(counter==0)
    {
        file<<this->tenure<<" : Correct";
        file<<" (O Errors) "<<endl;
    }
    file<<"----------------------------------------------------------------------------"<<endl;
    file.close();
}


void Transactions::check_id_trans()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("A2").to_string();
    this->id_trans=fname;
    int counter = 0;
    int tab[length];
    regex empty("^$");
    smatch match;
    int o = 0;
    int x=0;
    
    file_two.open("All_errors (Transactions).txt",ios::out);
    file_two<<this->id_trans<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("A", i)).value<int>();
        tab[o]=b;
        if(czy_wylosowana(o,b,tab))
        {
            if(x==0){
            file_two<<"Duplicats in cell : ";
            x++;
            }
            file_two<<"[A"<<i<<" -> value: "<<b<<"] ";

        }
        ++o;
        
    }
    file_two<<endl;
    
    file_two<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("A", j)).to_string();
        if(regex_search(cell,match,empty))
        {
            file_two<<"[A"<<j<<"] ";
        }
    }
        file_two<<endl;
    if(counter==0)
    {
        file_two<<this->id_trans<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_id_product()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("B2").to_string();
    this->id_product=fname;
    int counter = 0;
    regex empty("^$");
    smatch match;
   
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->id_product<<" :"<<endl;
   
    file_two<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("B", j)).to_string();
        if(regex_search(cell,match,empty))
        {
            file_two<<"[B"<<j<<"] ";
        }
    }
        file_two<<endl;
    if(counter==0)
    {
        file_two<<this->id_product<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_id_customer()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("C2").to_string();
    this->id_customer=fname;
    int counter = 0;
    regex empty("^$");
    smatch match;
   
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->id_customer<<" :"<<endl;
   
    file_two<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("C", j)).to_string();
        if(regex_search(cell,match,empty))
        {
            file_two<<"[C]"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->id_customer<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_date()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto id = trans_sheet.cell("D2").to_string();
    this->date=id;
    int counter = 0;
    smatch match;
    regex key("^[0-9]+$");
    regex empty("^$");
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->date<<" :"<<endl;
    for(int i = 3; i<=this->length;i++){
        auto b =trans_sheet.cell(xlnt::cell_reference("D", i)).to_string();
         if(regex_search(b,match,empty))
             continue;
           
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell D"<<i<<endl;
            }
            
    }
    file_two<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("D", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[D"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->date<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_online_order()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("E2").to_string();
    this->online_order=fname;
    int counter = 0;
    regex key("^[0-1]{1}+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->online_order<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("E", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell E"<<i<<endl;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("E", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[E"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->online_order<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_order_status()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("F2").to_string();
    this->order_status=fname;
    int counter = 0;
    regex key("\\b(Cancelled|Approved)\\w?\\b");
    regex key2("^[a-zA-Z]+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->order_status<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("F", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell F"<<i<<endl;
            }
            if(!regex_search(b,match,key2))
            {
                ++counter;
                file_two<<"Check cell F"<<i<<endl;
                continue;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("F", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[F"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->order_status<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_brand()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("G2").to_string();
    this->brand=fname;
    int counter = 0;
    regex key("^[a-zA-Z 0-9]+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->brand<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("G", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell G"<<i<<endl;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("G", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[G"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->brand<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_produkt_line()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("H2").to_string();
    this->produkt_line=fname;
    int counter = 0;
    regex key("\\b(Mountain|Standard|Road|Touring)\\w?\\b");
    regex key2("^[a-zA-Z]+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->produkt_line<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("H", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell H"<<i<<endl;
            }
            if(!regex_search(b,match,key2))
            {
                ++counter;
                file_two<<"Check cell H"<<i<<endl;
                continue;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("H", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[H"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->produkt_line<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_produkt_class()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("I2").to_string();
    this->produkt_class=fname;
    int counter = 0;
    regex key("\\b(low|medium|high)\\w?\\b");
    regex key2("^[a-zA-Z]+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->produkt_class<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("I", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell I"<<i<<endl;
            }
            if(!regex_search(b,match,key2))
            {
                ++counter;
                file_two<<"Check cell I"<<i<<endl;
                continue;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("I", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[I"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->produkt_class<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_produkt_size()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("J2").to_string();
    this->produkt_size=fname;
    int counter = 0;
    regex key("\\b(small|medium|large)\\w?\\b");
    regex key2("^[a-zA-Z]+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->produkt_size<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("J", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell J"<<i<<endl;
            }
            if(!regex_search(b,match,key2))
            {
                ++counter;
                file_two<<"Check cell J"<<i<<endl;
                continue;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("J", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[J"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->produkt_size<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_list_price()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("K2").to_string();
    this->list_price=fname;
    int counter = 0;
    regex key("^[0-9.]+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->list_price<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("K", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell K"<<i<<endl;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("K", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[K"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->list_price<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_standard_cost()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("L2").to_string();
    this->standard_cost=fname;
    int counter = 0;
    regex key("^[0-9.$]+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->standard_cost<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("L", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell L"<<i<<endl;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("L", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[L"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->standard_cost<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}
void Transactions::check_first_date()
{
    fstream file_two;
    xlnt::workbook workbook_two;
    workbook_two.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet trans_sheet = workbook_two.sheet_by_title("Transactions");
    auto fname = trans_sheet.cell("M2").to_string();
    this->first_date=fname;
    int counter = 0;
    regex key("^[0-9]{5}+$");
    regex empty("^$");
    smatch match;
    
    file_two.open("All_errors (Transactions).txt",ios::out|ios::app);
    file_two<<this->first_date<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =trans_sheet.cell(xlnt::cell_reference("M", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_two<<"Check cell M"<<i<<endl;
            }
    }
    file_two<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =trans_sheet.cell(xlnt::cell_reference("M", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_two<<"[M"<<j<<"] ";
        }
    }
    file_two<<endl;
    if(counter==0)
    {
        file_two<<this->first_date<<" : Correct";
        file_two<<" (O Errors) "<<endl;
    }
    file_two<<"----------------------------------------------------------------------------"<<endl;
    file_two.close();
}



void CustomerAddress::check_id()
{
    fstream file_three;
    xlnt::workbook workbook_three;
    workbook_three.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet address_sheet = workbook_three.sheet_by_title("CustomerAddress");
    auto fname = address_sheet.cell("A2").to_string();
    this->id=fname;
    int counter = 0;
    regex empty("^$");
    smatch match;
    int x=0;
    int q=1;
    int h;
    
    file_three.open("All_errors (CustomerAddress).txt",ios::out);
    file_three<<this->id<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =address_sheet.cell(xlnt::cell_reference("A", i)).value<int>();
        h = b-q;
        if(q!=b)
        {
            
            if(x==0){
            file_three<<"No addres id: "<<endl;
            x++;
            }
            for(int z = 0; z<h;z++)
            {
                file_three<<"id"<<q+z<<endl;
                counter++;
            }
            q=b;

        }
        ++q;
    }
    file_three<<"Empty cells : ";
    for( int j = 3; j<=this->length;j++)
    {
        auto cell =address_sheet.cell(xlnt::cell_reference("A", j)).to_string();
        if(regex_search(cell,match,empty))
        {
            file_three<<"[A"<<j<<"] ";
        }
    }
        file_three<<endl;
    if(counter==0)
    {
        file_three<<this->id<<" : Correct";
        file_three<<" (O Errors) "<<endl;
    }
    file_three<<"----------------------------------------------------------------------------"<<endl;
    file_three.close();
}
void CustomerAddress::check_address()
{
    fstream file_three;
    xlnt::workbook workbook_three;
    workbook_three.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet address_sheet = workbook_three.sheet_by_title("CustomerAddress");
    auto fname = address_sheet.cell("B2").to_string();
    this->address=fname;
    int counter = 0;
    regex key("^[a-zA-Z 0-9]+$");
    regex empty("^$");
    smatch match;
    
    file_three.open("All_errors (CustomerAddress).txt",ios::out|ios::app);
    file_three<<this->address<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =address_sheet.cell(xlnt::cell_reference("B", i)).to_string();
            if(regex_search(b,match,empty))
                continue;

            
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_three<<"Check cell B"<<i<<endl;
            }
    }
    file_three<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =address_sheet.cell(xlnt::cell_reference("B", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_three<<"[B"<<j<<"] ";
        }
    }
    file_three<<endl;
    if(counter==0)
    {
        file_three<<this->address<<" : Correct";
        file_three<<" (O Errors) "<<endl;
    }
    file_three<<"----------------------------------------------------------------------------"<<endl;
    file_three.close();
}
void CustomerAddress::check_post_code()
{
    fstream file_three;
    xlnt::workbook workbook_three;
    workbook_three.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet address_sheet = workbook_three.sheet_by_title("CustomerAddress");
    auto fname = address_sheet.cell("C2").to_string();
    this->post_code=fname;
    int counter = 0;
    regex key("^[0-9]{4}+$");
    regex empty("^$");
    smatch match;
    
    file_three.open("All_errors (CustomerAddress).txt",ios::out|ios::app);
    file_three<<this->post_code<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =address_sheet.cell(xlnt::cell_reference("C", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_three<<"Check cell C"<<i<<endl;
            }
    }
    file_three<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =address_sheet.cell(xlnt::cell_reference("C", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_three<<"[C"<<j<<"] ";
        }
    }
    file_three<<endl;
    if(counter==0)
    {
        file_three<<this->post_code<<" : Correct";
        file_three<<" (O Errors) "<<endl;
    }
    file_three<<"----------------------------------------------------------------------------"<<endl;
    file_three.close();
}
void CustomerAddress::check_state()
{
    fstream file_three;
    xlnt::workbook workbook_three;
    workbook_three.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet address_sheet = workbook_three.sheet_by_title("CustomerAddress");
    auto fname = address_sheet.cell("D2").to_string();
    this->state=fname;
    int counter = 0;
    regex key("^[a-z A-Z]+$");
    regex empty("^$");
    smatch match;
    
    file_three.open("All_errors (CustomerAddress).txt",ios::out|ios::app);
    file_three<<this->state<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =address_sheet.cell(xlnt::cell_reference("D", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_three<<"Check cell D"<<i<<endl;
            }
    }
    file_three<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =address_sheet.cell(xlnt::cell_reference("D", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_three<<"[D"<<j<<"] ";
        }
    }
    file_three<<endl;
    if(counter==0)
    {
        file_three<<this->state<<" : Correct";
        file_three<<" (O Errors) "<<endl;
    }
    file_three<<"----------------------------------------------------------------------------"<<endl;
    file_three.close();
}
void CustomerAddress::check_country()
{
    fstream file_three;
    xlnt::workbook workbook_three;
    workbook_three.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet address_sheet = workbook_three.sheet_by_title("CustomerAddress");
    auto fname = address_sheet.cell("E2").to_string();
    this->country=fname;
    int counter = 0;
    regex key("\\b(Australia)\\w?\\b");
    regex key2("^[a-zA-Z]+$");
    regex empty("^$");
    smatch match;
    
    file_three.open("All_errors (CustomerAddress).txt",ios::out|ios::app);
    file_three<<this->country<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =address_sheet.cell(xlnt::cell_reference("E", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key2))
            {
                ++counter;
                file_three<<"Check cell E"<<i<<endl;
                continue;
            }
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_three<<"Check cell E"<<i<<endl;
            }
    }
    file_three<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =address_sheet.cell(xlnt::cell_reference("E", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_three<<"[E"<<j<<"] ";
        }
    }
    file_three<<endl;
    if(counter==0)
    {
        file_three<<this->country<<" : Correct";
        file_three<<" (O Errors) "<<endl;
    }
    file_three<<"----------------------------------------------------------------------------"<<endl;
    file_three.close();
}

void CustomerAddress::check_valuations()
{
    fstream file_three;
    xlnt::workbook workbook_three;
    workbook_three.load("KPMG_VI_New_raw_data_update_final.xlsx");
    xlnt::worksheet address_sheet = workbook_three.sheet_by_title("CustomerAddress");
    auto fname = address_sheet.cell("F2").to_string();
    this->valuations=fname;
    int counter = 0;
    regex key("^[0-9]{1,2}+$");
    regex empty("^$");
    smatch match;
    
    file_three.open("All_errors (CustomerAddress).txt",ios::out|ios::app);
    file_three<<this->valuations<<" :"<<endl;
    for(int i = 3; i<=this->length;i++)
    {
        auto b =address_sheet.cell(xlnt::cell_reference("F", i)).to_string();
            if(regex_search(b,match,empty))
                continue;
            if(!regex_search(b,match,key))
            {
                ++counter;
                file_three<<"Check cell F"<<i<<endl;
            }
    }
    file_three<<"Empty cells : ";
        for( int j = 3; j<=this->length;j++)
    {
        auto cell =address_sheet.cell(xlnt::cell_reference("F", j)).to_string();
        if(regex_search(cell,match,empty))
        {
        file_three<<"[F"<<j<<"] ";
        }
    }
    file_three<<endl;
    if(counter==0)
    {
        file_three<<this->valuations<<" : Correct";
        file_three<<" (O Errors) "<<endl;
    }
    file_three<<"----------------------------------------------------------------------------"<<endl;
    file_three.close();
}

