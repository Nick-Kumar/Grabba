using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.Data.SqlClient;

namespace Grabba
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Outlook.MAPIFolder inbox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            inbox.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(InboxFolderItemAdded);
        }
        private void InboxFolderItemAdded(object Item)
        {
            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mailItem = (Item as Outlook.MailItem);
                //lets start with the sender
                Outlook.MailItem mailSender = (Item as Outlook.MailItem);
                if (mailSender.SenderEmailAddress == "nikhilanmolkumar@gmail.com")
                {
                    Regex expedia = new Regex(@"Expedia Booking # ");
                    Match machE = expedia.Match(mailItem.Subject);
                    if (machE.Success)       //subject Match i.e. Expedia or Booking
                    {
                        ExpediaMethod(mailItem);
                       
                    }
                    else
                    {
                        MessageBox.Show("NADA");
                    }
                }        
    

            }
        }

        private void ExpediaMethod(Outlook.MailItem mailItem)
        {
            String CheckInDateFinal = "";
            String CheckOutDateFinal = "";
            String BookedOnDateFinal = "";
            String roomTypeFinal = "";
            String CollectFinal = "";
            String guestNameFinal = "";
            int proper_resID = 0;
            int adultsFinal = 0;
            int NightsFinal = 0;
            int guestPhoneFinal = 0;

            

            //Reservation ID
            Regex resregex = new Regex(@"Reservation ID .: \d{9}");
            Match resmatch = resregex.Match(mailItem.Body);
            if (resmatch.Success)       //Reservation Id Match
            {
                //convert to string
                String resID = resmatch.ToString();
                Regex actual_res = new Regex(@"\d{9}");
                Match actual_resmatch = actual_res.Match(resID);
                if (actual_resmatch.Success)
                {
                    String actual_resID = actual_resmatch.ToString();
                     proper_resID = Int32.Parse(actual_resID);//Convert Reservation ID to integer

                }

            }

            //Booked On

            Regex bookedOnRegex = new Regex(@"Booked On ......: \d[1-31]\/\d[1-12]\/\d{4}");
            Match bookedOnmatch = bookedOnRegex.Match(mailItem.Body);
            if (bookedOnmatch.Success)       //Reservation Id Match
            {
                //convert to string
                String BookedOnString = bookedOnmatch.ToString();
                Regex BookedOnDateRegex = new Regex(@"\d[1-31]\/\d[1-12]\/\d{4}");
                Match BookedOnDatematch = BookedOnDateRegex.Match(BookedOnString);
                if (BookedOnDatematch.Success)
                {
                    BookedOnDateFinal = BookedOnDatematch.ToString();

                }

            }

            //Check In

            Regex checkInRegex = new Regex(@"Check In .......: \d[1-31]\/\d[1-12]\/\d{4}");
            Match checkInmatch = checkInRegex.Match(mailItem.Body);
            if (checkInmatch.Success)       //Check In Match
            {
                //convert to string
                String CheckInString = checkInmatch.ToString();
                Regex CheckInDateRegex = new Regex(@"\d[1-31]\/\d[1-12]\/\d{4}");
                Match CheckInDatematch = CheckInDateRegex.Match(CheckInString);
                if (CheckInDatematch.Success)
                {
                    CheckInDateFinal = CheckInDatematch.ToString();
                }

            }


            //Check Out

            Regex checkOutRegex = new Regex(@"Check Out ......: \d[1-31]\/\d[1-12]\/\d{4}");
            Match checkOutmatch = checkOutRegex.Match(mailItem.Body);
            if (checkOutmatch.Success)       //Check Out Match
            {
                //convert to string
                String CheckOutString = checkOutmatch.ToString();
                Regex CheckOutDateRegex = new Regex(@"\d[1-31]\/\d[1-12]\/\d{4}");
                Match CheckOutDatematch = CheckOutDateRegex.Match(CheckOutString);
                if (CheckOutDatematch.Success)
                {
                    CheckOutDateFinal = CheckOutDatematch.ToString();
                }

            }


            //Adults

            Regex AdultsRegex = new Regex(@"Adults .........: \d");
            Match Adultsmatch = AdultsRegex.Match(mailItem.Body);
            if (Adultsmatch.Success)       //Check Out Match
            {
                //convert to string
                String AdultsString = Adultsmatch.ToString();
                Regex AdultsNumRegex = new Regex(@"\d");
                Match AdultsNummatch = AdultsNumRegex.Match(AdultsString);
                if (AdultsNummatch.Success)
                {
                    String AdultsTempString = AdultsNummatch.ToString();
                    adultsFinal = Int32.Parse(AdultsTempString);
                }

            }


            //Nights

            Regex NightsRegex = new Regex(@"Nights ....: \d");
            Match Nightsmatch = NightsRegex.Match(mailItem.Body);
            if (Nightsmatch.Success)       //Check Out Match
            {
                //convert to string
                String NightsString = Nightsmatch.ToString();
                Regex NightsNumRegex = new Regex(@"\d");
                Match NightsNummatch = NightsNumRegex.Match(NightsString);
                if (NightsNummatch.Success)
                {
                    String NightsTempString = NightsNummatch.ToString();
                    NightsFinal = Int32.Parse(NightsTempString);
                }

            }

            //Room Type

            Regex deluxeRoomRegex = new Regex(@"Room Type Name .: Deluxe Room");
            Match deluxeRoomMatch = deluxeRoomRegex.Match(mailItem.Body);

            Regex OneBedroomAptRegex = new Regex(@"Room Type Name .: One Bedroom Apartment");
            Match OneBedroomAptMatch = OneBedroomAptRegex.Match(mailItem.Body);

            Regex TwoBedroomAptRegex = new Regex(@"Room Type Name .: Two Bedroom Apartment");
            Match TwoBedroomAptMatch = TwoBedroomAptRegex.Match(mailItem.Body);

            Regex SpaSuiteRegex = new Regex(@"Room Type Name .: Spa Suite");
            Match SpaSuiteMatch = SpaSuiteRegex.Match(mailItem.Body);

            if (deluxeRoomMatch.Success){
                roomTypeFinal = "DLXE";
            }
            else if(OneBedroomAptMatch.Success){
                roomTypeFinal = "A1";
            }
            else if (TwoBedroomAptMatch.Success)
            {
                roomTypeFinal = "A2K";
            }
            else if (SpaSuiteMatch.Success)
            {
                roomTypeFinal = "SPA";
            }

            //Collect Booking
            Regex ExpediaCollectRegex = new Regex(@"Expedia Collect Booking");
            Match ExpediaCollectMatch = ExpediaCollectRegex.Match(mailItem.Body);


            Regex HotelCollectRegex = new Regex(@"Hotel Collect Booking");
            Match HotelCollectMatch = HotelCollectRegex.Match(mailItem.Body);

            if (ExpediaCollectMatch.Success)
            {
                CollectFinal = "Expedia";
            }
            else if (HotelCollectMatch.Success)
            {
                CollectFinal = "Hotel";
            }

            //Guest Name

            Regex guestNameRegex = new Regex(@"Guest Name .....: [a-zA-Z]+\, [a-zA-Z]+");
            Match guestNamematch = guestNameRegex.Match(mailItem.Body);
            if (guestNamematch.Success)       //Check Out Match
            {
                //convert to string
                String guestNameString = guestNamematch.ToString();
                Regex guestNameStringRegex = new Regex(@"[a-zA-Z]+\, [a-zA-Z]+");
                Match guestNameStringmatch = guestNameStringRegex.Match(guestNameString);
                if (guestNameStringmatch.Success)
                {
                    guestNameFinal = guestNameStringmatch.ToString();
                    MessageBox.Show = ("Found" + guestNameFinal);
                }

            }

            //Guest Phone
            Regex guestPhoneregex = new Regex(@"Guest Phone ....: \d{9}");
            Match guestPhonematch = guestPhoneregex.Match(mailItem.Body);
            if (guestPhonematch.Success)       //Reservation Id Match
            {
                //convert to string
                String guestPhoneString = guestPhonematch.ToString();
                Regex guestPhoneStringRegex = new Regex(@"\d{9}");
                Match guestPhoneStringmatch = guestPhoneStringRegex.Match(guestPhoneString);
                if (guestPhoneStringmatch.Success)
                {
                    String Gphone = guestPhoneStringmatch.ToString();
                    guestPhoneFinal = Int32.Parse(Gphone);//Convert Reservation ID to integer

                }

            }




            //inserting into database
            SqlConnection testConnection = new SqlConnection("Data Source=localhost\\SQLEXPRESS;Database=RES;Trusted_Connection=True;");
            testConnection.Open();
            SqlCommand set = new SqlCommand("SET DATEFORMAT dmy", testConnection);
            SqlCommand cmd = new SqlCommand("INSERT INTO testTable(Reservation, GuestName, GuestPhone, BookedOn, CheckIn, CheckOut, Adults, Nights, RoomType, CollectBooking) VALUES(@resID,@GName, @GPhone, @bookedOn,@checkIn,@checkOut,@adults,@nights, @roomType, @CB)", testConnection);
            
            cmd.Parameters.AddWithValue("@resID", proper_resID);
            cmd.Parameters.AddWithValue("@bookedOn", BookedOnDateFinal);
            cmd.Parameters.AddWithValue("@checkIn", CheckInDateFinal);
            cmd.Parameters.AddWithValue("@checkOut", CheckOutDateFinal);
            cmd.Parameters.AddWithValue("@adults", adultsFinal);
            cmd.Parameters.AddWithValue("@nights", NightsFinal);
            cmd.Parameters.AddWithValue("@roomType", roomTypeFinal);
            cmd.Parameters.AddWithValue("@CB", CollectFinal);
            cmd.Parameters.AddWithValue("@GName", guestNameFinal);
            cmd.Parameters.AddWithValue("@GPhone", guestPhoneFinal);
            set.ExecuteNonQuery();
            cmd.ExecuteNonQuery();
            testConnection.Close();
            MessageBox.Show("Entered new reservation: " + proper_resID + " in database");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support – do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
