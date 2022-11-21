Video link - https://youtu.be/xhFz3KugXd0

Website link - https://buzz2daytech.blogspot.com/2022/04/activate-office-2021-using-cmd.html



Ms Office 2021 Activate करने के लिए नीचे दी गई Commands का आपको उपयोग लेना है, Office 2021 को डाउनलोड करने से लेकर Install करके Activate करने का पूरा प्रोसेस नीचे दिए गए वीडियो में बताया हुआ है जिसे आपको देख कर Steps को फॉलो करना है।  

### Step 1: Open cmd program with administrator rights.

### Step 2:  Get into the Office directory in cmd.

cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

cd /d %ProgramFiles%\Microsoft Office\Office16

### Step 3: Install Office 2021 volume license.

for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

This step is required. You can not install the KMS client product key of Office without a volume license.

### Step 4: Activate your Office using the KMS key.

Make sure your device is connected to the internet, then run the following commands.

cscript ospp.vbs /setprt:1688

cscript ospp.vbs /unpkey:6F7TH >nul

cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

cscript ospp.vbs /sethst:s8.uk.to

cscript ospp.vbs /act

If you see the error 0xC004F074, it means that your internet connection is unstable or the server is busy. Please make sure your device is online and try the command “act” again until you succeed.
