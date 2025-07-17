# 📅 Excel-7

![Status](https://img.shields.io/badge/status-active-brightgreen.svg)
![Excel](https://img.shields.io/badge/Microsoft-Excel-blue.svg)

## ✨ Project Description

**Excel-7** is a guide to working with dates and time in Microsoft Excel. Here you'll find practical tips, instructions, and illustrations about useful functions.

> 📚 **Goal:** Help you master Excel's date and time functions for everyday use—whether you're a beginner or advanced user!

---

## 📚 Table of Contents

- [Date Functions](#-date-functions)
- [Current Date and Time](#-current-date-and-time)
- [DATEDIF and Days Calculation](#-datedif-and-days-calculation)
- [Calculate Age](#-calculate-age)
- [WEEKDAY and Day Names](#-weekday-and-day-names)
- [NETWORKDAYS and Workdays](#-networkdays-and-workdays)
- [Quarter Calculation](#-quarter-calculation)
- [Screenshots](#-screenshots)
- [Requirements](#-requirements)
- [Author](#-author)

---

## 📆 Date Functions

To enter a date in Excel, use the "/" or "-" characters. To enter a time, use ":" (colon). You can also enter a date and a time in one cell.

> ℹ️ Note: Polish date format used here: day.month.year

To extract the year of a date, use the `YEAR` function.

![Date](Screenshots/Date.png)

> 💡 Tip: Use `MONTH` and `DAY` functions to extract month and day.

To add a number of years, months, and/or days, use the `DATE` function.

![Change Date](Screenshots/ChangeDate.png)

If you only need to add days you can do it like this:

![Add Days](Screenshots/AddDays.png)

---

## ⏰ Current Date and Time

To get the current date and time, use the `NOW` function.

![Current Date and Time](Screenshots/CurrentDateAndTime.png)

> ℹ️ Note: Use the `TODAY` function to enter today's date in Excel.  
> ℹ️ The result of the `NOW` function updates automatically whenever the sheet is recalculated.

To extract hour from date use the `HOUR` function.

![Hour](Screenshots/Hour.png)

> 💡 Tip: Use the `MINUTE` and `SECOND` functions to return minute and second.

To add a number of hours, minutes, and/or seconds, use the `TIME` function.

![Change Time](Screenshots/ChangeTime.png)

---

## 📏 DATEDIF and Days Calculation

The `DATEDIF` function in Excel calculates the number of days, months, or years between two dates. It has 3 arguments.

![DATEDIF](Screenshots/DateDif.png)

> ℹ️ Note:  
> - Type `"d"` for the third argument to get days between two dates  
> - `"m"` for months  
> - `"y"` for years  
> - `"yd"` to ignore years  
> - `"md"` to ignore months and years

You can also use the `DAYS` function to achieve the same result:

![DAYS](Screenshots/Days.png)

---

## 🎂 Calculate Age

We can use `DATEDIF` to calculate the age of a person:

![Age](Screenshots/Age.png)

Use this formula to calculate the exact age:

![Exact Age](Screenshots/ExactAge.png)

---

## 📅 WEEKDAY and Day Names

The `WEEKDAY` function in Excel returns a number from 1 (Sunday) to 7 (Saturday) representing the day of the week of a date.

![Weekday](Screenshots/Weekday.png)

We can use the `TEXT` function with `"dddd"` format to display that day:

![Text](Screenshots/Text.png)

> ℹ️ Note: "sobota" means Saturday in Polish.

---

## 🏢 NETWORKDAYS and Workdays

The `NETWORKDAYS` function in Excel returns the number of workdays between two dates, excluding weekends (Saturday and Sunday).

![NetWorkDays](Screenshots/NetWorkDays.png)

---

## 🗓️ Quarter Calculation

Formula that returns the quarter for a given date:

![RoundUp](Screenshots/RoundUp.png)

> ℹ️ Note: There's no built-in function in Excel to do this.  
> ℹ️ `ROUNDUP(x,0)` always rounds x up to the nearest integer.

---

## 📸 Screenshots

You can find all screenshots in the `/Screenshots` folder.

---

## ℹ️ Requirements

- Microsoft Excel (recommended: 2021/365 for modern functions)
- Windows OS

---

## 👨‍💻 Author

Project and documentation by **Kuba27x**  
Repository: [Kuba27x/Excel-7](https://github.com/Kuba27x/Excel-7)

---
