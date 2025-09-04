# UiPath RPA – Stock Price Trend Comparison  

## 📌 Introduction  
This project is part of the **UiPath RPA Specialization Capstone**.  
The goal is to automate the **stock price trend comparison** for two companies by extracting data from a website, storing it in Excel, generating a graph, and emailing the report to the customer.  

---

## 🎯 Project Background  
Organizations often need to track **daily stock price movements** of competitor companies. Doing this manually every 30 minutes during trading hours is inefficient and prone to human error.  
This automation project solves the problem by scheduling **automatic stock price collection, graph generation, and report delivery**.  

---

## 📝 Problem Statement  
- Stock prices of **Exxon RPA Corp.** and **WEX Academy Inc.** must be tracked every 30 minutes.  
- Prices need to be stored with **date and timestamp** in an Excel file.  
- At the end of trading hours, a **comparison graph** must be generated.  
- A final Excel report (with data + chart) should be **emailed automatically** to the customer.  

---

## ✅ Expected Output  
- Automated stock price extraction at 30-minute intervals.  
- Excel file (`Data.xlsx`) containing:  
  - System Date  
  - System Timestamp  
  - Exxon RPA Corp. stock price  
  - WEX Academy Inc. stock price  
- Auto-generated **double line graph** showing both companies’ stock price trends.  
- Automated email sent to customer with Excel file attached.  

---

## ⚙️ Environment Prerequisites  

### 1. Applications / Software  
- **UiPath Studio** (Academic Alliance Version 2021.10)  
- **Microsoft Excel** (to store data & create graphs)  
- **Microsoft Outlook** (to send reports)  
- **Microsoft Edge** (browser for stock market website)  

### 2. UiPath Packages  
- UiPath.Excel.Activities (v2.11.4)  
- UiPath.System.Activities (v22.4.1)  
- UiPath.UiAutomation.Activities (v21.10.5)  
- UiPath.Mail.Activities (v1.12.3)  

### 3. Initial Data Files  
- **Data.xlsx** → Pre-defined with headers:  
  - A1: System Date  
  - B1: System Timestamp  
  - C1: Exxon RPA Corp  
  - D1: WEX Academy Inc  
  *(also includes an empty chart ready to be populated)*  
- **Config.xlsx** → Pre-defined with:  
  - Receiver’s Email Address  
  - Body Message  
  - Mail Subject  

### 4. Other Requirements  
- Stable internet connection (for accessing the stock market site).  
- Ensure the system remains ON during trading hours (9:00 AM – 4:00 PM).  

---

## 🔄 Process Overview  
1. **Check Trading Time** → Only run if current time is between 9:30 AM – 4:00 PM.  
2. **Calculate Execution Frequency** → Number of runs until 4 PM (every 30 minutes).  
3. **Extract Date & Timestamp**.  
4. **Open Stock Market Website** → [RPA Stock Market](https://www.rpachallenge.com/assets/rpaStockMarket/index.html).  
5. **Scrape Stock Prices** → Get values for Exxon RPA Corp. and WEX Academy Inc.  
6. **Update Excel File** → Append data to `Data.xlsx`.  
7. **Repeat Until End of Day** → Run every 30 minutes until trading closes.  
8. **Generate Line Graph** → Double line chart showing stock price comparison.  
9. **Send Email** → Attach updated Excel file with chart to customer.  
10. **End Process**.  

---

## 📊 Example Output (Excel)  
- **Columns**: Date | Timestamp | Exxon RPA Corp | WEX Academy Inc  
- **Chart**: Double line graph showing price trends of both companies.  

---

## 🚀 How to Run  
1. Clone the repository:  
   ```bash
   git clone https://github.com/Ixzahlutf/uipath-rpa-stock-price-trend-comparison.git
2. Open the project in **UiPath Studio**.
3. Restore dependencies.
4. Configure Outlook mailbox + update `Config.xlsx` with recipient details.
5. Ensure `Data.xlsx` is blank with only headers before running.
6. Run `Main.xaml`.

---

## 🏆 Outcome

* Automated, scheduled stock price tracking
* Excel-based line graph for visualization
* Automatic daily report delivery via email

---

## 👩‍💻 Author

**Nurul Izzah Luthfiah Nur**
📍 Jakarta, Indonesia
📧 [izzahluthfiah@gmail.com](mailto:izzahluthfiah@gmail.com)
🔗 [LinkedIn](https://linkedin.com/in/izzahluthfiah) | [GitHub](https://github.com/Ixzahlutf)

---




