using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Windows.Forms;

namespace HospiceExcelTool
{
    class Conncetion
    {
        private MySqlConnection connection;
        private string server, database, uid, password;
        bool open = false;

        public Conncetion(bool open)
        {
            initialize();
            this.open = open;
        }
        private void initialize()
        {
            server = "77.71.152.127";
            database = "dbhospice";
            uid = "johndebono";
            password = "123456abc!";

            string connectionString = "SERVER=" + server + ";" + "DATABASE=" + database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";
            connection = new MySqlConnection(connectionString);
            //this.open = openConncetion();
            
        }
        public bool openConncetion()
        {
            try
              {
                connection.Open();
                return true;  
              }

            catch(MySqlException e)
            {
                switch (e.Number)
                {
                    case 0:
                        MessageBox.Show("Cannot connect to the Server. Contact Administrator");
                        break;
                    case 1045:
                        MessageBox.Show("Invalid Username/Password, Please try again");
                        break;
                    default:
                        MessageBox.Show(e.Message);
                        break;
                }
                return false;
            }
        }
        public bool closeConnection()
        {
            try
            {
                connection.Close();
                return true;
            }
            catch (MySqlException e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
        }
        public void insert(string name,string surname,string address,string street,string locality,string postCode,string idCard,int gender, long landline,long mobile,string email,int inContact,string month,string posted,int receiptNo,string duration)
        {
            if (connection.State == System.Data.ConnectionState.Open)
            {
                string insertQuery = "INSERT INTO members(Name,Surname,Address,Street,Locality,Postcode,IDCard,Gender,Landline,Mobile,Email,InContact,Month,Posted,Receitno,Duration_FK) VALUES ('" + name + "','" + surname + "','" + address + "','" + street + "','" + locality + "','" + postCode + "','" + idCard + "', " + gender + "," + landline + "," + mobile + ",'" + email + "'," + inContact + ",'" + month + "','" + posted + "'," + receiptNo + "," + duration + ")";
                MySqlCommand insertCommand = new MySqlCommand(insertQuery, connection);
                insertCommand.ExecuteNonQuery();
            }

        }

    }
}
