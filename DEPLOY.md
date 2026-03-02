# 🚀 خطوات الـ Deploy — مجاني 100%

## الخدمات المطلوبة (كلها مجانية)
| الخدمة | الدور | الرابط |
|--------|-------|--------|
| GitHub | حفظ الكود | github.com |
| Supabase | قاعدة البيانات | supabase.com |
| Render | تشغيل البرنامج | render.com |
| UptimeRobot | يصحّي السيرفر | uptimerobot.com |

---

## الخطوة 1 — GitHub (رفع الكود)

1. اعمل حساب على github.com
2. اضغط **New Repository**
3. اسمه: `dcr-web-app`  
4. اضغط **Create Repository**
5. افتح CMD في فولدر البرنامج وانفذ:

```bash
git init
git add .
git commit -m "Initial DCR app"
git branch -M main
git remote add origin https://github.com/اسمك/dcr-web-app.git
git push -u origin main
```

---

## الخطوة 2 — Supabase (قاعدة البيانات)

1. اعمل حساب على supabase.com
2. اضغط **New Project**
3. اختار:
   - **Name:** dcr-database
   - **Password:** اختار باسورد قوي (احتفظ بيه)
   - **Region:** Frankfurt (أقرب لمصر)
4. استنى ~2 دقيقة لحد ما يتنشأ

5. روح **Settings → Database**
6. في قسم **Connection string** اختار **URI**
7. انسخ الـ connection string — بيبدأ بـ:
   ```
   postgresql://postgres:[PASSWORD]@...supabase.co:5432/postgres
   ```
8. **احتفظ بيه** هتحتاجه في الخطوة الجاية

---

## الخطوة 3 — Render (تشغيل البرنامج)

1. اعمل حساب على render.com
2. اضغط **New → Web Service**
3. اختار **Connect GitHub** وابحث عن `dcr-web-app`
4. اضبط الإعدادات:
   - **Name:** dcr-web-app
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python server.py`
   - **Plan:** Free

5. قبل ما تضغط Deploy، روح **Environment Variables** وأضف:
   ```
   DATABASE_URL = postgresql://postgres:[PASSWORD]@...supabase.co:5432/postgres
   RENDER       = true
   ```

6. اضغط **Create Web Service**
7. استنى 3-5 دقايق للأول

8. هتلاقي الرابط جاهز زي:
   ```
   https://dcr-web-app.onrender.com
   ```

---

## الخطوة 4 — UptimeRobot (يصحّي السيرفر)

1. اعمل حساب على uptimerobot.com
2. اضغط **Add New Monitor**
3. اضبط:
   - **Monitor Type:** HTTP(s)
   - **Friendly Name:** DCR Web App
   - **URL:** `https://dcr-web-app.onrender.com/ping`
   - **Monitoring Interval:** Every 5 minutes
4. اضغط **Create Monitor**

✅ **خلاص! السيرفر هيفضل صاحي 24/7**

---

## تحديث البرنامج (لو عملت تعديلات)

```bash
git add .
git commit -m "Update app"
git push
```
Render هيعمل deploy تلقائي في ~2 دقيقة

---

## لو حصل أي مشكلة

- **البرنامج مش بيفتح:** افتح Render → Logs وشوف الخطأ
- **البيانات اتمسحت:** مستحيل مع Supabase إلا لو مسحتها إنت
- **الدخول بطيء:** أول مرة بعد نوم السيرفر بتستغرق 30 ثانية

---

## المميزات

- ✅ مجاني 100% — مش محتاج بطاقة بنك
- ✅ يشتغل 24/7 بفضل UptimeRobot
- ✅ البيانات محفوظة دايماً على Supabase
- ✅ أي حد يفتح الرابط من أي مكان في العالم
- ✅ تحديث تلقائي لما تعمل git push
