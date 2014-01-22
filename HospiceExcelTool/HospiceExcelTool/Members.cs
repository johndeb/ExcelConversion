using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace HospiceExcelTool
{
    public class Members
    {
        private int number;
        private string name;
        private string surname;
        private string id;
        private string dob;
        private string town;
        private string email;
        private string title;
        private string address;
        private string street;
        private string postCode;
        private int gender;
        private long landLine;
        private long mobile;
        private int inContact;
        private string monthStarted;
        private int receiptNumber;
        private string duration;


        public Members(int number, string name, string surname, string id, string dob, string email, string town,string title, string address,string street,string postCode,string gender, string landLine,string mobile,string inContact,string monthStarted,int reciptNumber,string duration)
        {
            this.number = number;
            this.name = name;
            this.surname = surname;
            this.id = id;
            this.dob = dob;
            this.town = town;
            this.email = email;
            this.title = title;
            this.address = address;
            this.street = street;
            this.postCode = postCode;
            if (gender.ToLower() == "male" || gender.ToLower() == "m")
            {
                this.gender = 1;
            }
            else if (gender.ToLower() == "female" || gender.ToLower() == "f")
            {
                this.gender = 2;
            }
            else
            {
                this.gender = 0;
            }
            if (landLine == "")
            {
                
            }
            else
            {
                this.landLine = Convert.ToInt64(landLine);
            }
            if (mobile == "")
            {
                
            }
            else
            {
                this.mobile = Convert.ToInt64(mobile);
            }

            if (inContact.ToLower() == "yes" || inContact.ToLower() == "y")
            {
                this.inContact = 1;
            }
            else if (inContact.ToLower() == "no" || inContact.ToLower() == "n")
            {
                this.inContact = 2;
            }
            else
            {
                this.inContact = 0;
            }
            if (monthStarted.ToString() == "")
            {
            }
            else
            {
               // this.monthStarted = DateTime.ParseExact(monthStarted.ToString().Trim(), "yyyy/MM/dd", CultureInfo.InvariantCulture);
                string[] date = monthStarted.Split('/');
                DateTime dt = new DateTime(Convert.ToInt16(date[2]),Convert.ToInt16(date[1]),Convert.ToInt16(date[0]));
                CultureInfo cf = new CultureInfo("ja-JP");
                this.monthStarted = dt.ToString("d", cf);
                
            }
            this.receiptNumber = reciptNumber;
            this.duration = duration;
        }

        public int Number
        {
            get { return number; }
            set { number = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Surname
        {
            get { return surname; }
            set { surname = value; }
        }

        public string Id
        {
            get { return id; }
            set { id = value; }
        }

        public string Dob
        {
            get { return dob; }
            set { dob = value; }
        }

        public string Town
        {
            get { return town; }
            set { town = value; }
        }

        public string Email
        {
            get { return email; }
            set { email = value; }
        }

        public string Title
        {
            get { return title; }
            set { title = value; }
        }

        public string Address
        {
            get { return address; }
            set { address = value; }
        }

        public string Street
        {
            get { return street; }
            set { street = value; }
        }

        public string PostCode
        {
            get { return postCode; }
            set { postCode = value; }
        }

        public int Gender
        {
            get { return gender; }
            set { gender = value; }
        }

        public long LandLine
        {
            get { return landLine; }
            set { landLine = value; }
        }

        public long Mobile
        {
            get { return mobile; }
            set { mobile = value; }
        }

        public int InContact
        {
            get { return inContact; }
            set { inContact = value; }
        }

        public string MonthStarted
        {
            get { return monthStarted; }
            set { monthStarted = value; }
        }

        public int ReceiptNumber
        {
            get { return receiptNumber; }
            set { receiptNumber = value; }
        }

        public string Duration
        {
            get { return duration; }
            set { duration = value; }
        } 
    }
}

