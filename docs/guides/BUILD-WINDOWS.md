# Windows-এ BPSC Question Studio বিল্ড করুন

Windows 10 বা 11 x64-এর জন্য। সোর্স চালানোর জন্য Python 3.12 x64 দরকার। তৈরি Portable ফোল্ডারে Python থাকে। প্রশ্নপত্র PDF তৈরি ও পুরো PDF টেস্ট চালাতে LibreOffice ও প্রশ্নপত্রে ব্যবহৃত বাংলা ফন্ট ইনস্টল করুন। ম্যাপিং PDF তৈরিতে LibreOffice লাগে না।

- [Python 3.12.10 Windows installer (64-bit)](https://www.python.org/downloads/release/python-31210/)
- [LibreOffice Windows x86-64](https://www.libreoffice.org/download/)

## ১. সোর্স থেকে সরাসরি চালান

1. Source ZIP পুরোটা লেখার অনুমতি আছে এমন ফোল্ডারে Extract করুন, যেমন `D:\BPSC\QuestionStudio-Source`।
2. Python ইনস্টল করার সময় Python Launcher চালু রাখুন। Command Prompt-এ `py -3.12 --version` দিয়ে যাচাই করুন।
3. Extract করা ফোল্ডারে **Run-Source.bat** ডাবল-ক্লিক করুন। প্রথমবার এটি একটি স্থানীয় Python পরিবেশ বানিয়ে dependencies নামাবে। এরপর সাইন-ইন উইন্ডো খুলবে।

একই কাজ Command Prompt থেকে করতে পারেন।

```bat
cd /d "D:\BPSC\QuestionStudio-Source"
py -3.12 -m venv .venv
".venv\Scripts\python.exe" -m pip install -r requirements.txt
".venv\Scripts\python.exe" main.py
```

## ২. Windows x64 Portable EXE বানান

1. Source ZIP-এর ভেতরে **Build-Portable.bat** ডাবল-ক্লিক করুন। প্রথম বিল্ডের জন্য ইন্টারনেট লাগবে।
2. স্ক্রিপ্ট Zig compiler, Windows x64 embedded Python এবং নির্দিষ্ট সংস্করণের Python wheels নামাবে। এরপর অ্যাপের আইকন ও সংস্করণসহ Windows launcher কম্পাইল করবে।
3. শেষ হলে `dist\BPSC-Question-Studio\BPSC Question Studio.exe` খুলুন। অন্য PC-তে দিতে হলে পুরো `dist\BPSC-Question-Studio` ফোল্ডার ZIP করুন। `.exe` তার পাশের `Runtime`, `main.py`, `mcq`, এবং `assets` ব্যবহার করে।
4. আবার বিল্ড করার আগে আগের `dist\BPSC-Question-Studio` ফোল্ডারটি Rename বা Move করুন। এতে ওই ফোল্ডারে পরীক্ষা করে তৈরি করা ফাইলগুলোও সংরক্ষিত থাকবে।

### একই বিল্ড Command Prompt থেকে হাতে করতে চাইলে

নিচের কমান্ডগুলো একে একে চালান। Source ZIP যে ফোল্ডারে Extract করেছেন, প্রথম লাইনে সেই পাথ দিন।

```bat
cd /d "D:\BPSC\QuestionStudio-Source"
mkdir .build\manual
py -3.12 -m venv .build\manual\tools
set "BUILD_PY=.build\manual\tools\Scripts\python.exe"
"%BUILD_PY%" -m pip install --only-binary=:all: ziglang==0.17.1
curl.exe --fail --location "https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip" --output ".build\manual\python-3.12.10-embed-amd64.zip"
"%BUILD_PY%" -m pip download --only-binary=:all: --platform win_amd64 --implementation cp --python-version 312 --abi cp312 --dest ".build\manual\wheels" -r requirements.txt -c packaging\portable-constraints.txt
"%BUILD_PY%" -m ziglang rc /FO.build\manual\branding.res packaging\branding.rc
"%BUILD_PY%" -m ziglang cc -target x86_64-windows-gnu -O2 -municode -Wl,--subsystem,windows packaging\launcher.c ".build\manual\branding.res" -o ".build\manual\Launcher.exe"
"%BUILD_PY%" packaging\assemble_portable.py --python-zip ".build\manual\python-3.12.10-embed-amd64.zip" --wheels ".build\manual\wheels" --launcher ".build\manual\Launcher.exe" --output "dist\BPSC-Question-Studio"
start "" "dist\BPSC-Question-Studio\BPSC Question Studio.exe"
```

## ৩. PDF ও টেস্ট চালান

- **Common Settings → Renderer**-এ প্রয়োজন হলে `C:\Program Files\LibreOffice\program\soffice.com` সিলেক্ট করুন।
- Portable EXE থেকে উইন্ডো না খুললে একই ফোল্ডারের **Run-Diagnostics.bat** চালিয়ে Python-এর বার্তা দেখুন।
- সোর্সের পূর্ণ টেস্ট চালাতে আগে **Run-Source.bat** দিয়ে dependencies ইনস্টল করুন। তারপর Source ফোল্ডারের Command Prompt-এ নিচের কমান্ড দিন। LibreOffice অন্য পাথে থাকলে প্রথম কমান্ডে সেই পাথ দিন।

```bat
set "MCQ_TEST_RENDERER=C:\Program Files\LibreOffice\program\soffice.com"
".venv\Scripts\python.exe" -m unittest discover -s tests -v
```

গাইডে থাকা ইনস্টলার ও dependencies বিল্ড করার PC-র জন্য। যে Windows PC-তে তৈরি Portable ফোল্ডার শুধু চালাবেন, সেখানে Python ও Zig আলাদা করে ইনস্টল করতে হবে না।
