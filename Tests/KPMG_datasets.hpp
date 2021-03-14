#ifndef KPMG_DATASETS_HPP
#define KPMG_DATASETS_HPP

using namespace std;

class Customer_Demographic
{
public:
    int length;
    string id;
    string first_name;
    string last_name;
    string gender;
    string past_3_years;
    string date;
    string job_tittle;
    string job_industry_category;
    string wealth_segment;
    string deceased_indicator;
    string defaulte;
    string own_cars;
    string tenure;
    void check_id();
    void check_first_name();
    void check_last_name();
    void check_gender();
    void check_past_3_years();
    void check_date();
    void check_job_tittle();
    void checK_job_industry_category();
    void check_wealth_segment();
    void check_deceased_indicator();
    void check_defalute();
    void check_own_cars();
    void check_tenure();
    Customer_Demographic();
};

class Transactions
{
public:
    string id_trans;
    string id_product;
    string id_customer;
    string date;
    string online_order;
    string order_status;
    string brand;
    string produkt_line;
    string produkt_class;
    string produkt_size;
    string list_price;
    string standard_cost;
    string first_date;
    
    int length;
    void check_id_trans();
    void check_id_product();
    void check_id_customer();
    void check_date();
    void check_online_order();
    void check_order_status();
    void check_brand();
    void check_produkt_line();
    void check_produkt_class();
    void check_produkt_size();
    void check_list_price();
    void check_standard_cost();
    void check_first_date();
    Transactions();
};

class CustomerAddress
{
public:
    string id;
    string address;
    string post_code;
    string state;
    string country;
    string valuations;
    int length;
    void check_id();
    void check_address();
    void check_post_code();
    void check_state();
    void check_country();
    void check_valuations();
    CustomerAddress();
};

#endif