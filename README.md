# BPSC Question Studio

বাংলাদেশ সরকারি কর্ম কমিশনের জন্য Windows x64-এ চলা একটি অফলাইন ডেস্কটপ অ্যাপ। **Set Generation** থেকে প্রশ্নের সেট, ম্যাপিং এবং DOCX/PDF তৈরি করা যায়। **Room Print Packs** থেকে Excel-এর রুম তালিকা ও সেটের PDF ব্যবহার করে সেন্টার ও রুমভিত্তিক প্রিন্ট ফাইল তৈরি করা যায়।

অ্যাপ ব্যবহারের বিস্তারিত ধাপ [ব্যবহার নির্দেশিকা](docs/guides/USER-GUIDE.md)-তে আছে।

## বিল্ডের প্রস্তুতি

- Windows 10 বা 11, 64-bit এবং Python 3.12 x64 লাগবে। Python ইনস্টল করার সময় **Python Launcher** চালু রাখুন। Command Prompt-এ `py -3.12 --version` দিয়ে যাচাই করুন।
- প্রথম বিল্ডে compiler, embedded Python ও Python dependencies নামাতে ইন্টারনেট লাগবে।
- Set Generation থেকে প্রশ্নের PDF বানাতে LibreOffice ও প্রশ্নপত্রে ব্যবহৃত ফন্ট লাগবে। প্রয়োজনে অ্যাপের **Common Settings → Renderer**-এ LibreOffice-এর `program/soffice.com` বাছুন।
- [Python](https://www.python.org/downloads/windows/) · [LibreOffice](https://www.libreoffice.org/download/)

## সোর্স থেকে Windows Portable বিল্ড

1. Git রিপোজিটরি clone করুন অথবা সম্পূর্ণ Source ZIP Extract করুন। যে ফোল্ডারে `Build-Portable.bat` আছে, সেটিই প্রজেক্টের মূল ফোল্ডার।
2. মূল ফোল্ডার থেকে `Build-Portable.bat` চালান। এটি Windows launcher কম্পাইল করে এবং `dist\BPSC-Question-Studio`-তে অ্যাপ ও রানটাইম সাজায়।
3. বিল্ড শেষ হলে `dist\BPSC-Question-Studio\BPSC Question Studio.exe` চালু করুন।
4. অন্য Windows PC-তে চালাতে পুরো `dist\BPSC-Question-Studio` ফোল্ডার কপি বা ZIP করুন। Portable ফোল্ডারে Python runtime অন্তর্ভুক্ত থাকে। EXE-টি পাশের `Runtime`, `mcq`, `assets` ও অন্যান্য ফাইল ব্যবহার করে।
5. আবার বিল্ড করার আগে পুরোনো `dist\BPSC-Question-Studio` ফোল্ডারটি সরিয়ে বা নাম বদলে রাখুন। বিল্ড স্ক্রিপ্ট নতুন ফোল্ডার তৈরি করবে।

সোর্স থেকে সরাসরি চালাতে `Run-Source.bat` ব্যবহার করুন। হাতে বিল্ড, টেস্ট ও ডায়াগনস্টিকের ধাপ [Windows Build Guide](docs/guides/BUILD-WINDOWS.md)-এ আছে।

## Git-এ সোর্স আপডেট

1. Source ZIP-এর মূল ফোল্ডারের ভেতরের সোর্স, টেস্ট, ডকুমেন্টেশন ও বিল্ড ফাইলগুলো আপনার Git working tree-তে কপি করুন।
2. `git status --short` দিয়ে পরিবর্তন দেখুন। সোর্সের সঙ্গে `Config/Template`-এর ডিফল্ট টেমপ্লেট দুটি রাখুন। স্থানীয় `.studio/`, ব্যক্তিগত কনফিগারেশন, পরীক্ষার ডেটা, `.venv/`, `.build/` ও `dist/` stage করার বাইরে রাখুন।
3. প্রয়োজনীয় ফাইল stage করে `git diff --cached --stat` যাচাই করুন। তারপর আপনার রিপোজিটরির নিয়মে commit ও push করুন।

এই README স্থায়ী পরিচিতি ও বিল্ড নির্দেশিকার জন্য। রিলিজভিত্তিক পরিবর্তন [রিলিজ ইতিহাসে](docs/version-history/CHANGELOG.md) রাখুন।
