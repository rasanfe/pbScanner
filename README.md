# pbScanner — Manejar el escáner desde PowerBuilder (WIA) 🖨️

![PowerBuilder](https://img.shields.io/badge/PowerBuilder-2025-2D6CDF?style=flat-square)
![.NET](https://img.shields.io/badge/.NET-10-512BD4?style=flat-square&logo=dotnet&logoColor=white)
![WIA](https://img.shields.io/badge/Windows-WIA%20(COM)-0078D6?style=flat-square&logo=windows&logoColor=white)
![Blog](https://img.shields.io/badge/blog-rsrsystem-FF5722?style=flat-square&logo=blogger&logoColor=white)

## 📋 ¿Qué es esto?

Un ejemplo PowerBuilder para **manejar un escáner** de verdad: listar los escáneres que hay
conectados, lanzar el escaneo y traerte la imagen escaneada al disco. Como me inspiré en una app de
**Windows Forms** (ved los créditos), he **recreado su interfaz en PowerBuilder**.

¿Y cómo habla PowerBuilder con el escáner? A través de **WIA** (*Windows Image Acquisition*), que es
la API de Windows para dispositivos de captura de imagen. Pero WIA es **COM**, y eso desde
PowerScript es incómodo de tocar a pelo. Así que metemos por medio una librería .NET (`ScannerWia`)
que envuelve WIA, y la cargamos como `dotnetobject` con el **.NET DLL Importer** de PB. Eso nos
genera un proxy que desde PowerScript se usa como un objeto nativo. Fijaos:

- **Listar escáneres** → `ListScanners()` devuelve los dispositivos disponibles.
- **Escanear** → `Scan(escaner, formato, rutaSalida, nombreFichero)` lanza la captura y guarda la
  imagen.
- Y de propina, `ConvertImageToTxt(...)` para pasarle **OCR (Tesseract)** a lo escaneado,
  `GetPageCount()` y `GetErrorText()` para recoger el fallo como texto.

## 🔗 Motor .NET

El trabajo lo hace la librería .NET **`ScannerWia`** (clase `ScannerWia`), que envuelve el COM de
WIA:

- Se **despliega** ya compilada en la carpeta `DotNet\ScannerWia\` de este propio ejemplo, para que
  clones, compiles y funcione sin tocar nada más.
- Se **consume** desde PowerBuilder como `dotnetobject` (proxy creado con el .NET DLL Importer).
- El **código fuente** vive en `Blog\Net10\ScannerWia` (antes estaba en `Net8`) y se
  recompila/despliega con el script **`desplegar_dotnet.bat`** (hace `dotnet publish` y espeja las
  DLLs a la carpeta `DotNet` de cada ejemplo).
- Repo del proyecto .NET (Visual Studio 2022): <https://github.com/rasanfe/ScannerWia>

> 🆕 **Nota didáctica (migración a .NET 10):** WIA es COM, y **antes** el proyecto usaba
> `<COMReference>`, que obliga a `tlbimp` en cada build → `dotnet build` fallaba con **`MSB4803`** y
> solo compilaba desde Visual Studio. **Ya está resuelto:** se genera **una sola vez** el interop
> **`Interop.WIA.dll`** con `tlbimp` (desde `wiaaut.dll`, namespace `WIA`) y se referencia como
> `<Reference>` con `EmbedInteropTypes=true`. Así **compila y publica por CLI** (`dotnet build`),
> como el resto de ejemplos, sin Visual Studio y sin `MSB4803`. El código fuente no cambia (sigue
> `using WIA;`).

## 🛠️ Requisitos

- **PowerBuilder 2025** para abrir y compilar la solución.
- **.NET 10 Runtime** instalado en la máquina → <https://dotnet.microsoft.com/en-us/download/dotnet/10.0>
- La carpeta `DotNet\ScannerWia\` con las DLLs desplegadas (ya viene en el repo).
- Un **escáner compatible con WIA** conectado (en Windows, lo normal).

## ▶️ Cómo probarlo

1. Clona el repo y abre `pbScanner.pbsln` con PowerBuilder 2025.
2. Compila (Full Build) y ejecuta.
3. Pulsa para **listar los escáneres**, elige uno y lanza el **escaneo**.
4. La imagen escaneada se guarda en la ruta indicada; opcionalmente puedes pasarle OCR.

## 🔗 Repo PowerBuilder

<https://github.com/rasanfe/pbScanner>

## 🙌 Créditos

Inspirado en el artículo y el repositorio de **Our Code World**:

- Artículo: <https://ourcodeworld.com/articles/read/382/creating-a-scanning-application-in-winforms-with-csharp>
- Repositorio original (C# WinForms): <https://github.com/ourcodeworld/csharp-scanner-wia>

Como el repositorio original es una aplicación de Windows Forms, he recreado su interfaz en
PowerBuilder.

---

> ¡Nos vemos en el próximo artículo! Y recuerda: en PowerBuilder, los límites solo están en nuestra imaginación. 🚀

📨 **Blog:** <https://rsrsystem.blogspot.com/>
