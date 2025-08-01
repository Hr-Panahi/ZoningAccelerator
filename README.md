# 🏗️ Zoning Accelerator CLI

A command-line tool for comparing and extracting zoning data from Excel files, including dwelling types, ancillary types, and permitted uses.

---

## 📦 How to Use

### 1. **Download**
- Copy or extract the `ZoningAccelerator.exe` file (or full `publish` folder) to a folder on your computer.

---

### 2. **Run the Tool**
- Double-click `ZoningAccelerator.exe`, or
- Open a terminal or command prompt and run:
  ```bash
  ZoningAccelerator.exe
  ```

---

### 3. **Choose an Option**

When launched, you'll see a menu like this:

```
========== ZONING ACCELERATOR ==========

Select an option:
  1. Compare DwellingTypes
  2. Compare AncillaryTypes
  3. Compare PermittedUses
  4. Compare TypeOfUses
  5. Run All Comparisons
  6. Get Unique Dwellings
  7. Get Unique Ancillaries
  8. Get Unique Permitted Uses
  q. Quit

Choice:
```

---

## 📁 Output Location

All results are saved to a folder named:

```
Saved Files/
```

inside the application's directory.

Each file is named with the city Excel filename and auto-numbered to avoid overwriting:
```
UniqueDwellings_<CityName>.txt
AllComparisonResults_<CityName>.xlsx
...
```

---

## ✅ Requirements

No installation required — this is a **self-contained executable**. The tool works on:
- Windows 10/11 (64-bit)

---

## 💬 Questions?

If you run into issues or have suggestions, you know who to contact!

---
