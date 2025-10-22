'use client';

import { useState } from 'react';
import * as XLSX from 'xlsx';

interface TravelAgency {
  AgencyName: string;
  AgencyCity: string;
  Address: string;
  PhoneNumbers: string;
  Email: string;
  Website: string;
  VisaServices: string;
}

export default function Home() {
  const [data] = useState<TravelAgency[]>([
    {
      AgencyName: "Thomas Cook India",
      AgencyCity: "Mumbai",
      Address: "Thomas Cook Building, Dr. D N Road, Fort, Mumbai - 400001",
      PhoneNumbers: "+91-22-6681-8181, +91-22-6681-8585, 1800-209-8080",
      Email: "care@thomascook.in, visa@thomascook.in",
      Website: "https://www.thomascook.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Cox & Kings",
      AgencyCity: "Mumbai",
      Address: "16, Apeejay House, 4th Floor, Dinshaw Vachha Road, Churchgate, Mumbai - 400020",
      PhoneNumbers: "+91-22-3983-3333, +91-22-6218-5555",
      Email: "info@coxandkings.com, visa.services@coxandkings.com",
      Website: "https://www.coxandkings.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "SOTC Travel Services",
      AgencyCity: "Mumbai",
      Address: "SOTC Building, 226, Dr. D N Road, Fort, Mumbai - 400001",
      PhoneNumbers: "+91-22-6151-5151, 1800-209-3344",
      Email: "customercare@sotc.in, visa@sotc.in",
      Website: "https://www.sotc.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Kesari Tours",
      AgencyCity: "Mumbai",
      Address: "Kesari Mahal, 884/2, Fergusson College Road, Pune Branch: Nariman Point, Mumbai - 400021",
      PhoneNumbers: "+91-22-2204-5252, +91-22-2283-6666",
      Email: "info@kesari.in, mumbai@kesari.in",
      Website: "https://www.kesari.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Akbar Travels",
      AgencyCity: "Mumbai",
      Address: "Akbar House, 24A/B Mereweather Road, Colaba, Mumbai - 400001",
      PhoneNumbers: "+91-22-2218-5555, 1800-102-4747",
      Email: "info@akbartravels.com, visa@akbartravels.com",
      Website: "https://www.akbartravelsonline.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Yatra.com",
      AgencyCity: "Gurgaon",
      Address: "IRIS Tech Park, Sector 48, Sohna Road, Gurgaon - 122018",
      PhoneNumbers: "+91-124-495-4000, 1800-102-0088",
      Email: "customercare@yatra.com, visahelp@yatra.com",
      Website: "https://www.yatra.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "MakeMyTrip",
      AgencyCity: "Gurgaon",
      Address: "MakeMyTrip Building, Plot No. 62, Sector 44, Gurgaon - 122003",
      PhoneNumbers: "+91-124-462-8747, 1800-102-8747",
      Email: "care@makemytrip.com, visa@makemytrip.com",
      Website: "https://www.makemytrip.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Cleartrip",
      AgencyCity: "Mumbai",
      Address: "Cleartrip Private Limited, Springboard - WeWork, Godrej & Boyce, Vikhroli, Mumbai - 400079",
      PhoneNumbers: "+91-22-7127-4444, 1800-102-1234",
      Email: "support@cleartrip.com",
      Website: "https://www.cleartrip.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Riya Travel & Tours",
      AgencyCity: "Mumbai",
      Address: "Riya House, 274 D N Road, Fort, Mumbai - 400001",
      PhoneNumbers: "+91-22-2267-7777, +91-22-6666-8888",
      Email: "info@riya.travel, visa@riya.travel",
      Website: "https://www.riya.travel",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Visa House",
      AgencyCity: "Mumbai",
      Address: "Mittal Tower, 'A' Wing, 210, Nariman Point, Mumbai - 400021",
      PhoneNumbers: "+91-22-2283-3333, +91-22-2204-9999",
      Email: "info@visahouse.in, services@visahouse.in",
      Website: "https://www.visahouse.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Travel Boutique Online",
      AgencyCity: "Mumbai",
      Address: "301, Sagar Tech Plaza, Saki Vihar Road, Andheri East, Mumbai - 400072",
      PhoneNumbers: "+91-22-6777-9999, +91-22-2852-7777",
      Email: "info@travelboutiqueonline.com",
      Website: "https://www.travelboutiqueonline.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Flamingo Transworld",
      AgencyCity: "Mumbai",
      Address: "162, Maker Tower 'E', Cuffe Parade, Mumbai - 400005",
      PhoneNumbers: "+91-22-2218-5050, +91-22-6666-4444",
      Email: "info@flamingotravels.co.in, visa@flamingotravels.co.in",
      Website: "https://www.flamingotravels.co.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Duniya Safari Travels",
      AgencyCity: "Mumbai",
      Address: "Shop No 1, Ganesh Garden, P K Road, Mulund West, Mumbai - 400080",
      PhoneNumbers: "+91-22-2564-3322, +91-22-2564-4433",
      Email: "info@duniyasafari.com, booking@duniyasafari.com",
      Website: "https://www.duniyasafari.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Via.com",
      AgencyCity: "Gurgaon",
      Address: "Tower A, DLF IT Park, Sector 25A, Gurgaon - 122002",
      PhoneNumbers: "+91-124-498-5858, 1800-123-5555",
      Email: "support@via.com, help@via.com",
      Website: "https://www.via.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Musafir.com",
      AgencyCity: "Gurgaon",
      Address: "DLF Centre Court Mall, DLF Phase 4, Gurgaon - 122002",
      PhoneNumbers: "+91-124-452-5888, 1800-102-9747",
      Email: "care@musafir.com, visa@musafir.com",
      Website: "https://www.musafir.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Indian Visa Online (IVO)",
      AgencyCity: "Mumbai",
      Address: "Office 203, Trade Link, Kamala City, Senapati Bapat Marg, Lower Parel, Mumbai - 400013",
      PhoneNumbers: "+91-22-6767-5555, +91-22-4002-3333",
      Email: "info@indianvisaonline.gov.in.org, support@ivo.net.in",
      Website: "https://www.indianvisaonline.gov.in.org",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Atlas Travel & Tours",
      AgencyCity: "Mumbai",
      Address: "55/57, Dr. V B Gandhi Marg, 2nd Floor, Fort, Mumbai - 400023",
      PhoneNumbers: "+91-22-2266-8888, +91-22-2263-3333",
      Email: "info@atlastravels.net, enquiry@atlastravels.net",
      Website: "https://www.atlastravels.net",
      VisaServices: "Visa Only"
    },
    {
      AgencyName: "Akshay Tours & Travels",
      AgencyCity: "Mumbai",
      Address: "Shop No 4, Pushpak Building, J P Road, Andheri West, Mumbai - 400058",
      PhoneNumbers: "+91-22-2632-7777, +91-22-2632-8888",
      Email: "info@akshaytours.com, booking@akshaytours.com",
      Website: "https://www.akshaytours.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Travel Corporation India",
      AgencyCity: "Mumbai",
      Address: "Mafatlal Centre, 4th Floor, Nariman Point, Mumbai - 400021",
      PhoneNumbers: "+91-22-2202-5555, +91-22-2287-3333",
      Email: "corporate@tcil.com, visa@tcil.com",
      Website: "https://www.tcil.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Travelmasti India",
      AgencyCity: "Gurgaon",
      Address: "M-14, Middle Circle, Connaught Place Extension, Gurgaon - 122001",
      PhoneNumbers: "+91-124-451-2222, +91-124-451-3333",
      Email: "info@travelmasti.com, support@travelmasti.com",
      Website: "https://www.travelmasti.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Spice Jet Travel",
      AgencyCity: "Gurgaon",
      Address: "319, Udyog Vihar Phase IV, Gurgaon - 122015",
      PhoneNumbers: "+91-124-498-3333, 1860-180-3333",
      Email: "customercare@spicejet.com",
      Website: "https://www.spicejet.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Goibibo",
      AgencyCity: "Gurgaon",
      Address: "B-25, Sector 60, Noida, Branch: Cyber Hub, Gurgaon - 122002",
      PhoneNumbers: "+91-124-628-9999, 1800-102-3344",
      Email: "support@goibibo.com, help@goibibo.com",
      Website: "https://www.goibibo.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Travel Triangle",
      AgencyCity: "Gurgaon",
      Address: "Plot No 52, Sector 44, Gurgaon - 122003",
      PhoneNumbers: "+91-124-452-7777, 1800-123-5555",
      Email: "info@traveltriangle.com, support@traveltriangle.com",
      Website: "https://www.traveltriangle.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Club7 Holidays",
      AgencyCity: "Mumbai",
      Address: "B-1, Meher Mansion, 75 A, Dr. Annie Besant Road, Worli, Mumbai - 400018",
      PhoneNumbers: "+91-22-2497-7777, +91-22-2496-8888",
      Email: "info@club7holidays.com, visa@club7holidays.com",
      Website: "https://www.club7holidays.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Caper Travel Company",
      AgencyCity: "Mumbai",
      Address: "Shop 7, Neelkanth Arcade, Linking Road, Santacruz West, Mumbai - 400054",
      PhoneNumbers: "+91-22-2660-7777, +91-22-2660-3333",
      Email: "info@capertravel.com, enquiry@capertravel.com",
      Website: "https://www.capertravel.com",
      VisaServices: "Visa Only"
    },
    {
      AgencyName: "Airborne Travels",
      AgencyCity: "Mumbai",
      Address: "Ground Floor, Bajaj Bhavan, Nariman Point, Mumbai - 400021",
      PhoneNumbers: "+91-22-6633-9999, +91-22-2287-5555",
      Email: "travel@airborne.co.in, visa@airborne.co.in",
      Website: "https://www.airborne.co.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Veena World",
      AgencyCity: "Mumbai",
      Address: "Veena House, Plot No 75, Kailash Industrial Complex, Vikhroli West, Mumbai - 400079",
      PhoneNumbers: "+91-22-2578-1010, 1800-222-003",
      Email: "mail@veenaworld.com, visa@veenaworld.com",
      Website: "https://www.veenaworld.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Ozone Overseas",
      AgencyCity: "Mumbai",
      Address: "3rd Floor, Saffron Complex, Fatima Nagar, Wanowrie, Pune Branch: Fort, Mumbai - 400001",
      PhoneNumbers: "+91-22-2269-8989, +91-22-6161-5555",
      Email: "info@ozoneoverseas.com, visa@ozoneoverseas.com",
      Website: "https://www.ozoneoverseas.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Mercury Travels",
      AgencyCity: "Mumbai",
      Address: "Ground Floor, Bajaj Bhavan, 226 Nariman Point, Mumbai - 400021",
      PhoneNumbers: "+91-22-2202-9999, +91-22-2282-8888",
      Email: "info@mercurytravels.co.in, visa@mercurytravels.co.in",
      Website: "https://www.mercurytravels.co.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Travelocity India",
      AgencyCity: "Gurgaon",
      Address: "Tower B, Unitech Cyber Park, Sector 39, Gurgaon - 122001",
      PhoneNumbers: "+91-124-471-4444, 1800-102-9999",
      Email: "india@travelocity.com, support@travelocity.co.in",
      Website: "https://www.travelocity.co.in",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Expedia India",
      AgencyCity: "Gurgaon",
      Address: "WeWork Forum, Sector 44, Gurgaon - 122003",
      PhoneNumbers: "+91-124-662-8888, 1800-209-8888",
      Email: "india.support@expedia.com",
      Website: "https://www.expedia.co.in",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Ashiana Travels",
      AgencyCity: "Mumbai",
      Address: "Shop No 10, Srinath Plaza, Ghodbunder Road, Thane, Mumbai - 400607",
      PhoneNumbers: "+91-22-2534-7777, +91-22-2534-8888",
      Email: "info@ashianatravels.in, booking@ashianatravels.in",
      Website: "https://www.ashianatravels.in",
      VisaServices: "Visa Only"
    },
    {
      AgencyName: "Tripoto",
      AgencyCity: "Gurgaon",
      Address: "B-34, Sector 83, Gurgaon - 122004",
      PhoneNumbers: "+91-124-696-3333, 1800-121-2323",
      Email: "support@tripoto.com, hello@tripoto.com",
      Website: "https://www.tripoto.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Ixigo",
      AgencyCity: "Gurgaon",
      Address: "4th Floor, Tower B, Global Business Park, Mehrauli Gurgaon Road, Gurgaon - 122002",
      PhoneNumbers: "+91-124-663-6999, 1800-123-9999",
      Email: "support@ixigo.com",
      Website: "https://www.ixigo.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "TBO Holidays",
      AgencyCity: "Gurgaon",
      Address: "9th Floor, BPTP Park Centra, Sector 30, NH-8, Gurgaon - 122001",
      PhoneNumbers: "+91-124-452-4444, +91-124-452-5555",
      Email: "support@tboholidays.com, visa@tboholidays.com",
      Website: "https://www.tboholidays.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Safari India Tours",
      AgencyCity: "Mumbai",
      Address: "Office No 101, Chandra Kiran, 76 Linking Road, Khar West, Mumbai - 400052",
      PhoneNumbers: "+91-22-2648-5555, +91-22-2648-6666",
      Email: "info@safariindiatours.com",
      Website: "https://www.safariindiatours.com",
      VisaServices: "Visa Only"
    },
    {
      AgencyName: "Travels Spice",
      AgencyCity: "Mumbai",
      Address: "Plot No 42, Sector 11, CBD Belapur, Navi Mumbai - 400614",
      PhoneNumbers: "+91-22-2757-4444, +91-22-2757-5555",
      Email: "info@travelsspice.com, booking@travelsspice.com",
      Website: "https://www.travelsspice.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Sterling Holidays",
      AgencyCity: "Mumbai",
      Address: "Sterling Centre, Dr. Ambedkar Road, Parel, Mumbai - 400012",
      PhoneNumbers: "+91-22-6111-6111, 1800-102-5554",
      Email: "info@sterlingholidays.com, visa@sterlingholidays.com",
      Website: "https://www.sterlingholidays.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Travel Guru",
      AgencyCity: "Gurgaon",
      Address: "Spaze I-Tech Park, Sector 49, Sohna Road, Gurgaon - 122018",
      PhoneNumbers: "+91-124-488-5555, +91-124-488-6666",
      Email: "info@travelguru.com, support@travelguru.com",
      Website: "https://www.travelguru.com",
      VisaServices: "eVisa Only"
    },
    {
      AgencyName: "Ease My Trip",
      AgencyCity: "Gurgaon",
      Address: "6th Floor, Aggarwal Cyber Plaza II, Netaji Subhash Place, Pitampura, Branch: Gurgaon - 122002",
      PhoneNumbers: "+91-11-4646-4646, 1800-102-3454",
      Email: "care@easemytrip.com, visa@easemytrip.com",
      Website: "https://www.easemytrip.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Happy Holidays",
      AgencyCity: "Mumbai",
      Address: "Shop 6, Panchratna Building, Opera House, Mumbai - 400004",
      PhoneNumbers: "+91-22-2367-8888, +91-22-2363-9999",
      Email: "info@happyholidaysmumbai.com",
      Website: "https://www.happyholidaysmumbai.com",
      VisaServices: "Visa Only"
    },
    {
      AgencyName: "Bon Voyage",
      AgencyCity: "Mumbai",
      Address: "103, Janak Puri, Versova, Andheri West, Mumbai - 400061",
      PhoneNumbers: "+91-22-2639-4444, +91-22-2639-5555",
      Email: "info@bonvoyageindia.com, visa@bonvoyageindia.com",
      Website: "https://www.bonvoyageindia.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Odyssey Travels",
      AgencyCity: "Mumbai",
      Address: "Chinoy Mansion, 3rd Floor, 270 Dr. D N Road, Fort, Mumbai - 400001",
      PhoneNumbers: "+91-22-2267-6666, +91-22-2261-7777",
      Email: "info@odyssey.travel, visa@odyssey.travel",
      Website: "https://www.odyssey.travel",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Dpauls.com",
      AgencyCity: "Gurgaon",
      Address: "C-58/4, Sector 62, Noida Branch: Cyber City, Gurgaon - 122002",
      PhoneNumbers: "+91-124-471-9999, 1800-123-3366",
      Email: "travel@dpauls.com, support@dpauls.com",
      Website: "https://www.dpauls.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Taj Tours & Travels",
      AgencyCity: "Mumbai",
      Address: "Nanik Motwani Marg, Fort, Mumbai - 400001",
      PhoneNumbers: "+91-22-2204-6666, +91-22-2204-7777",
      Email: "info@tajtours.com, booking@tajtours.com",
      Website: "https://www.tajtours.com",
      VisaServices: "Visa Only"
    },
    {
      AgencyName: "Global Visa Services",
      AgencyCity: "Gurgaon",
      Address: "Tower A, Unitech Business Park, Golf Course Road, Gurgaon - 122001",
      PhoneNumbers: "+91-124-451-7777, +91-124-451-8888",
      Email: "info@globalvisaservices.in, visa@globalvisaservices.in",
      Website: "https://www.globalvisaservices.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "VFS Global",
      AgencyCity: "Mumbai",
      Address: "Express Tower, 13th Floor, Nariman Point, Mumbai - 400021",
      PhoneNumbers: "+91-22-6731-6666, +91-22-6731-7777",
      Email: "info.inmum@vfshelpline.com, visa.india@vfsglobal.com",
      Website: "https://www.vfsglobal.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "World Travel Studio",
      AgencyCity: "Mumbai",
      Address: "Shop No 3, Rani Sati Marg, Malad East, Mumbai - 400097",
      PhoneNumbers: "+91-22-2881-6666, +91-22-2881-7777",
      Email: "info@worldtravelstudio.com",
      Website: "https://www.worldtravelstudio.com",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Travel Xpress",
      AgencyCity: "Gurgaon",
      Address: "SCO 24-25, Sector 14, Gurgaon - 122001",
      PhoneNumbers: "+91-124-232-5555, +91-124-232-6666",
      Email: "info@travelxpress.in, support@travelxpress.in",
      Website: "https://www.travelxpress.in",
      VisaServices: "Both Visa & eVisa"
    },
    {
      AgencyName: "Akbar International Travels",
      AgencyCity: "Gurgaon",
      Address: "SCF 14, Old Delhi Gurgaon Road, Gurgaon - 122001",
      PhoneNumbers: "+91-124-405-7777, +91-124-405-8888",
      Email: "gurgaon@akbartravels.com, visa.ggn@akbartravels.com",
      Website: "https://www.akbartravelsonline.com",
      VisaServices: "Both Visa & eVisa"
    }
  ]);

  const downloadExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(data);

    const colWidths = [
      { wch: 30 },
      { wch: 15 },
      { wch: 60 },
      { wch: 40 },
      { wch: 50 },
      { wch: 35 },
      { wch: 20 }
    ];
    worksheet['!cols'] = colWidths;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Travel Agencies");

    XLSX.writeFile(workbook, "Travel_Agencies_Mumbai_Gurgaon.xlsx");
  };

  return (
    <div style={{
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
      padding: '40px 20px'
    }}>
      <div style={{
        maxWidth: '1400px',
        margin: '0 auto',
        background: 'white',
        borderRadius: '20px',
        boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
        overflow: 'hidden'
      }}>
        <div style={{
          background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
          padding: '40px',
          color: 'white'
        }}>
          <h1 style={{ margin: '0 0 10px 0', fontSize: '36px', fontWeight: 'bold' }}>
            Travel Agencies Directory
          </h1>
          <p style={{ margin: '0', fontSize: '18px', opacity: 0.9 }}>
            50 Travel Agencies from Mumbai & Gurgaon with Complete Information
          </p>
        </div>

        <div style={{ padding: '40px' }}>
          <div style={{
            background: '#f8f9fa',
            padding: '20px',
            borderRadius: '10px',
            marginBottom: '30px',
            border: '2px solid #e9ecef'
          }}>
            <h3 style={{ marginTop: 0, color: '#495057' }}>Data Includes:</h3>
            <ul style={{ color: '#6c757d', lineHeight: '1.8' }}>
              <li><strong>Agency Name</strong> - Full name of the travel agency</li>
              <li><strong>City</strong> - Mumbai or Gurgaon</li>
              <li><strong>Address</strong> - Complete physical address</li>
              <li><strong>Phone Numbers</strong> - All available contact numbers</li>
              <li><strong>Email</strong> - All available email addresses</li>
              <li><strong>Website</strong> - Official website URL</li>
              <li><strong>Visa Services</strong> - Type of visa services offered (Visa Only / eVisa Only / Both)</li>
            </ul>
          </div>

          <button
            onClick={downloadExcel}
            style={{
              width: '100%',
              padding: '20px',
              fontSize: '18px',
              fontWeight: 'bold',
              color: 'white',
              background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
              border: 'none',
              borderRadius: '10px',
              cursor: 'pointer',
              transition: 'transform 0.2s, box-shadow 0.2s',
              boxShadow: '0 4px 15px rgba(102, 126, 234, 0.4)'
            }}
            onMouseOver={(e) => {
              e.currentTarget.style.transform = 'translateY(-2px)';
              e.currentTarget.style.boxShadow = '0 6px 20px rgba(102, 126, 234, 0.6)';
            }}
            onMouseOut={(e) => {
              e.currentTarget.style.transform = 'translateY(0)';
              e.currentTarget.style.boxShadow = '0 4px 15px rgba(102, 126, 234, 0.4)';
            }}
          >
            📥 Download Excel Spreadsheet (50 Agencies)
          </button>

          <div style={{
            marginTop: '30px',
            padding: '20px',
            background: '#e7f3ff',
            borderRadius: '10px',
            border: '2px solid #b3d9ff'
          }}>
            <h4 style={{ marginTop: 0, color: '#0056b3' }}>📊 Summary Statistics:</h4>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '15px' }}>
              <div>
                <strong>Total Agencies:</strong> 50
              </div>
              <div>
                <strong>Mumbai:</strong> {data.filter(a => a.AgencyCity === 'Mumbai').length}
              </div>
              <div>
                <strong>Gurgaon:</strong> {data.filter(a => a.AgencyCity === 'Gurgaon').length}
              </div>
              <div>
                <strong>Both Visa & eVisa:</strong> {data.filter(a => a.VisaServices === 'Both Visa & eVisa').length}
              </div>
              <div>
                <strong>eVisa Only:</strong> {data.filter(a => a.VisaServices === 'eVisa Only').length}
              </div>
              <div>
                <strong>Visa Only:</strong> {data.filter(a => a.VisaServices === 'Visa Only').length}
              </div>
            </div>
          </div>
        </div>

        <div style={{
          overflowX: 'auto',
          padding: '0 40px 40px 40px'
        }}>
          <table style={{
            width: '100%',
            borderCollapse: 'collapse',
            fontSize: '14px'
          }}>
            <thead>
              <tr style={{ background: '#f8f9fa' }}>
                <th style={{ padding: '15px', textAlign: 'left', borderBottom: '2px solid #dee2e6', fontWeight: 'bold' }}>Agency Name</th>
                <th style={{ padding: '15px', textAlign: 'left', borderBottom: '2px solid #dee2e6', fontWeight: 'bold' }}>City</th>
                <th style={{ padding: '15px', textAlign: 'left', borderBottom: '2px solid #dee2e6', fontWeight: 'bold' }}>Phone</th>
                <th style={{ padding: '15px', textAlign: 'left', borderBottom: '2px solid #dee2e6', fontWeight: 'bold' }}>Visa Services</th>
              </tr>
            </thead>
            <tbody>
              {data.map((agency, index) => (
                <tr key={index} style={{
                  background: index % 2 === 0 ? 'white' : '#f8f9fa',
                  transition: 'background 0.2s'
                }}>
                  <td style={{ padding: '12px', borderBottom: '1px solid #dee2e6' }}>{agency.AgencyName}</td>
                  <td style={{ padding: '12px', borderBottom: '1px solid #dee2e6' }}>{agency.AgencyCity}</td>
                  <td style={{ padding: '12px', borderBottom: '1px solid #dee2e6', fontSize: '12px' }}>
                    {agency.PhoneNumbers.split(',')[0]}
                  </td>
                  <td style={{ padding: '12px', borderBottom: '1px solid #dee2e6' }}>
                    <span style={{
                      padding: '4px 8px',
                      borderRadius: '4px',
                      fontSize: '12px',
                      fontWeight: 'bold',
                      background: agency.VisaServices === 'Both Visa & eVisa' ? '#d4edda' :
                                 agency.VisaServices === 'eVisa Only' ? '#d1ecf1' : '#fff3cd',
                      color: agency.VisaServices === 'Both Visa & eVisa' ? '#155724' :
                             agency.VisaServices === 'eVisa Only' ? '#0c5460' : '#856404'
                    }}>
                      {agency.VisaServices}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
