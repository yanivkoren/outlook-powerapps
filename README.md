# Outlook → Power Apps Integration

## 📌 תיאור

פתרון המאפשר פתיחת אפליקציית Power Apps מתוך מייל ב־Outlook, תוך העברת נתוני המייל (msgId, convId, subject) כפרמטרים.

---

## 📂 מבנה הפרויקט

### 1. manifest.xml

קובץ הגדרה של Outlook Add-in
משמש להתקנה ב־Outlook (Custom Add-in)

תפקיד:

* הוספת כפתור “פתח ב-Power Apps” לריבון
* הגדרת הרשאות
* טעינת קובץ ה־HTML

---

### 2. openPowerApps.html

קובץ Web (HTML + JavaScript + Office.js)

מיקום:
https://yanivkoren.github.io/outlook-powerapps/openPowerApps.html

תפקיד:

* שליפת נתוני המייל מתוך Outlook
* בניית URL עם פרמטרים
* פתיחת Power Apps

---

## 🔗 קישור האפליקציה (Power Apps)

https://apps.powerapps.com/play/e/default-89e846c1-57c1-4a4c-b65a-6d2107dbdb64/a/71905a42-b9c2-42d1-b242-d63f52f2b0cf?tenantId=89e846c1-57c1-4a4c-b65a-6d2107dbdb64

האפליקציה מקבלת:

* msgId
* convId
* subject

---

## 📥 קליטת פרמטרים ב-Power Apps

```powerfx
Set(varMsgId, Param("msgId"));
Set(varConvId, Param("convId"));
Set(varSubject, Param("subject"));
```

---

## 🚀 תהליך עבודה

1. התקנת ה־Add-in דרך manifest.xml ב־Outlook
2. פתיחת מייל
3. לחיצה על הכפתור
4. פתיחת Power Apps עם נתוני המייל

---

## 🧠 זרימה

Outlook → Add-in → HTML → Power Apps

---

## ⚠️ הערות

* הקבצים חייבים להיות נגישים ב־HTTPS
* ה־repository ציבורי
* שינוי במניפסט דורש התקנה מחדש
* שינוי ב־HTML מתעדכן ללא התקנה מחדש (ייתכן cache)

---
