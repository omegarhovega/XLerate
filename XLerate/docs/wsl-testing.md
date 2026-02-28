# WSL2 Testing Guide (Excel Add-in)

`npm start` uses `office-addin-debugging` desktop sideload.  
That path is not supported on Linux/WSL (`Platform not supported: linux`).

Use one of these workflows instead.

## 1) WSL dev server + manual sideload (recommended in WSL)

1. Start the add-in web server in WSL:

```bash
cd /home/dmin/excel-plugin/XLerate/XLerate
npm run start:wsl
```

2. In Windows Excel Desktop or Excel Web, sideload the manifest manually:
   - Manifest path: `\\wsl.localhost\Ubuntu\home\dmin\excel-plugin\XLerate\XLerate\manifest.xml`

3. Open task pane and test buttons.

If port `3000` is busy:

```bash
npm run stop:wsl
```

## 2) Automatic desktop sideload (Windows shell only)

Run from PowerShell/CMD (not WSL):

```powershell
cd \\wsl.localhost\Ubuntu\home\dmin\excel-plugin\XLerate\XLerate
npm start
```

This uses `office-addin-debugging` on Windows (`win32`) and can sideload automatically.

## Certificate note

The dev server runs on `https://localhost:3000`.  
If Windows blocks it, trust the WSL dev cert in Windows certificate store:

- `\\wsl.localhost\Ubuntu\home\dmin\.office-addin-dev-certs\localhost.crt`
