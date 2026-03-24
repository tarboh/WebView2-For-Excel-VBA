## 🚀 VBA WebView2 Integration Project

<img width="1916" height="1028" alt="image" src="https://github.com/user-attachments/assets/a2bcff2c-3443-42c7-9315-f2555433a61a" />

### Overview
This project is a next-generation VBA development framework that aims to directly place and control **Microsoft Edge WebView2** on UserForms—without requiring external library installations or setting ActiveX references from Excel VBA.

With the end of support for Internet Explorer (IE), browser control in VBA has reached a major turning point. This repository hacks the `WebView2Loader.dll` natively embedded in modern Office installations, bypassing standard VBA limitations to provide a modern UI/UX via low-level COM manipulation using **DispCallFunc**.

---

### 🔥 Project Breakthroughs
Previously, using WebView2 from VBA required building a .NET COM wrapper and registering it in the Windows Registry (which demands Administrator privileges)—an impractical barrier for corporate environments.

However, this project achieves zero-dependency operation based on the following **discoveries**:

* **Leveraging Pre-installed Assets**: 
  By directly tapping into the `WebView2Loader.dll` found inside:
  `C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin`
  using standard `Declare` statements, we can spin up a WebView2 Environment without altering the host OS environment.
* **Pure vTable Hacking**: 
  Once the `WebView2Environment` is captured, every single operation is executed via function pointer calls using `DispCallFunc`. This achieves a complete **Zero-Reference-Setting** architecture.

---

### 🛠 Setup & Initialization Flow
The precise, under-the-hood process to capture the WebView2 object is as follows:

1. **Create Environment**: Call `CreateCoreWebView2EnvironmentWithOptions` from `WebView2Loader.dll`.
2. **Await Callback**: Intercept the completion notification via a standard module.
3. **Generate Controller**: Execute `CreateCoreWebView2Controller` to reserve the browser viewport.
4. **Capture WebView2**: Trap the core `ICoreWebView2` instance using the `GetWebView2` method.
5. **Window Fitting**: Use Windows APIs to find the `Chrome_WidgetWin_0` class window and fit it pixel-perfect onto the VBA UserForm.

---

### 🌟 The Future We Unlock
By completing this project, VBA developers will unlock:

* **Modern GUI Integration**: Build rich user interfaces inside UserForms using modern HTML5, CSS3, and JavaScript.
* **Moving Away from Legacy ActiveX**: Replace notoriously unstable TreeViews and ListViews with high-performance, web-based components.
* **Full Control via Event Handlers**: Complete mastery over web scraping and automated browser manipulation by self-implementing event handlers (such as `NavigationCompleted`).

---

## ⚠️ Troubleshooting: Excel Crashes on Form Startup (Memory Address Table Issue)

### Problem
When you open the UserForm, Excel might crash immediately with an **"Automation error / Exception occurred"** or a silent termination. 

This happens because the VBA compiler sometimes fails to statically evaluate or initialize the memory address of standard module functions using `AddressOf` if they are not explicitly called or held elsewhere. If VBA doesn't map these addresses correctly, calling low-level Win32 APIs (like `DispCallFunc` or `CreateCoreWebView2EnvironmentWithOptions`) will lead to an access violation (Null/Invalid pointer crash).

### Solution
To force the VBA compiler to evaluate `AddressOf` and keep the vTable memory address valid, add a dummy Sub at the bottom of your standard module:

```vba
' Dummy Sub to prevent VBA compiler optimization / address loss
Public Sub RegisterNavigationCompleted_()
    Static vTable As LongPtr
    vTable = GetAddr(AddressOf Handler_QueryInterface)
End Sub

---

### 📚 謝辞
本プロジェクトの着想にあたり、重要な技術的知見を共有してくださった **もっくん** 様（[参考記事](https://eschamali.github.io/StarterWebScrapingKit/#userform-powershell)）に深く感謝いたします。
